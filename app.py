"""
Customs Shipment CSV/Excel Automation Tool
Flask web application for cleaning and formatting customs shipment reports.
"""

import os
import io
import re
import logging
import uuid
from datetime import datetime

import pandas as pd
from flask import (
    Flask, request, render_template, send_file,
    jsonify, flash, redirect, url_for
)
from werkzeug.utils import secure_filename

# ── App Configuration ────────────────────────────────────────────────────────

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "customs-automation-secret-2024")

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"csv", "xlsx", "xls"}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16 MB

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Constants & State ────────────────────────────────────────────────────────
REQUIRED_COLUMNS = ["Manifested Description", "CE Commodity Description", "CE Item HSCode"]

# In-memory store for uploaded dataframes during a session
SESSIONS = {}

# ── Helper Functions ─────────────────────────────────────────────────────────

def allowed_file(filename: str) -> bool:
    """Check if the file extension is permitted."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_first_hs_code(raw_hs: str) -> str:
    """
    CE Item HSCode field looks like '1|6204120010' or '1 2|6210321422 6210321400'.
    Extract the first valid HS code (10-digit numeric string).
    """
    if pd.isna(raw_hs) or not str(raw_hs).strip():
        return ""
    raw = str(raw_hs).strip()
    # After the pipe, codes are space-separated — take first
    if "|" in raw:
        after_pipe = raw.split("|", 1)[1].strip()
        codes = after_pipe.split()
        if codes:
            # Normalise: remove dots and spaces
            code = re.sub(r"[\.\s]", "", codes[0])
            return code
    # Fallback: find any 6-10 digit sequence
    match = re.search(r"\d{6,10}", raw)
    return match.group() if match else ""


def clean_mid_code(mid_raw: str) -> str:
    """
    Remove ALL non-letter characters from a MID code, keeping only A-Z letters.
    If 'MID' is present, it is removed.
    If the resulting string does not start with 'PK', 'PK' is prepended.
    """
    mid = str(mid_raw).strip()
    # Remove ALL non-letter characters
    cleaned = re.sub(r"[^A-Za-z]", "", mid).upper()
    
    # Remove the word 'MID' if it exists anywhere
    cleaned = cleaned.replace("MID", "")
    
    # Replace SKT with SIA if it is at the end of the alphabetic part
    if cleaned.endswith("SKT"):
        cleaned = cleaned[:-3] + "SIA"
        
    # Prepend PK only if ends with SIA (Sialkot company code — shipper forgot PK prefix)
    if cleaned and not cleaned.startswith("PK") and cleaned.endswith("SIA"):
        cleaned = "PK" + cleaned
        
    return cleaned


def is_valid_mid_code(text: str, raw_mid: str) -> bool:
    """
    Validate a CLEANED MID code.
    Also takes raw_mid to check for explicit 'NIC'.
    """
    text = text.strip()
    raw_upper = str(raw_mid).strip().upper()
    
    if not text:
        return False
        
    # NIC is invalid only if cleaned MID starts with or equals "NIC"
    if text.startswith("NIC") or text == "NIC":
        return False
        
    # Must contain only uppercase letters
    if not re.fullmatch(r"[A-Z]+", text):
        return False
    # Must start with PK (Pakistan country code)
    if not text.startswith("PK"):
        return False
    # Must be at least 7 characters
    if len(text) < 7:
        return False
    
    return True


def strip_known_prefixes(text: str) -> str:
    """
    Remove known garbage prefixes from the start of a Manifested Description.
    """
    # Strip PSWSHIPMENT prefix (with optional spaces/dashes between PSW and SHIPMENT)
    text = re.sub(r'^PSW[\s\-_]*SHIPMENT\s*:\s*', '', text, flags=re.IGNORECASE)
    # No longer strip MID: here as it is handled generally
    return text.strip()


def parse_manifested_description(md: str, fallback_hs: str) -> dict:
    """
    Parse a single Manifested Description string into its three components.

    Returns a dict with keys:
      mid_code, hs_code, product_desc, cleaned, has_mid_format, has_hs_code
    """
    md = str(md).strip() if not pd.isna(md) else ""
    md = strip_known_prefixes(md)
    
    # Check if we should convert alternative formats to slash format
    # Skip if already has '/' AND first segment looks like a clean MID (no spaces)
    has_natural_slash = "/" in md
    first_seg = md.split("/", 1)[0] if has_natural_slash else ""
    first_seg_clean = bool(first_seg) and not re.search(r"\s", first_seg) and re.match(r"^[A-Za-z]{2,}", first_seg)

    if not has_natural_slash or not first_seg_clean:
        # Match: MID(HSCODE)PRODUCT or MID HSCODE PRODUCT or MID:HSCODE:PRODUCT or MID.HSCODE.PRODUCT
        pattern = r"^(PK[A-Za-z0-9]*)\s*[\(\.\:\s]\s*([\d\.\s]{6,15})\s*[\)\.\:\s]\s*(.*)$"
        match = re.match(pattern, md, re.IGNORECASE)
        if match:
            md = f"{match.group(1)}/{match.group(2)}/{match.group(3)}"

    # If no slash at all after conversion — not a MID-code style row; return as-is
    if "/" not in md:
        return {
            "mid_code": "",
            "hs_code": fallback_hs,
            "product_desc": md,
            "cleaned": md,
            "has_mid_format": False,
            "has_hs_code": bool(fallback_hs),
            "mid_modified": False,
        }

    parts = md.split("/")
    raw_mid = parts[0].strip()

    if len(parts) > 1:
        # Agar second part digit se start hai → HS code hai
        if re.match(r"^\s*\d", parts[1]):
            raw_hs_field = parts[1].strip()
            raw_desc = "/".join(parts[2:]).strip()
        else:
            # Second part digit se start nahi → yeh product description hai
            # Original "/" preserve karte huye sab ko join karo
            raw_hs_field = ""
            raw_desc = "/".join(parts[1:]).strip()
    else:
        raw_hs_field = ""
        raw_desc = ""

    # If MID is missing at the start (e.g. " /123456/DESC" or "123456/DESC" where 1st part is just digits)
    # Or if explicit "NIC" is in raw_mid
    if not raw_mid or re.fullmatch(r"\d+", raw_mid):
        return {
            "mid_code": "",
            "hs_code": fallback_hs,
            "product_desc": md,
            "cleaned": md,
            "has_mid_format": False,
            "has_hs_code": bool(fallback_hs),
            "mid_modified": False,
        }

    mid_code = clean_mid_code(raw_mid)

    if not is_valid_mid_code(mid_code, raw_mid):
        return {
            "mid_code": "",
            "hs_code": fallback_hs,
            "product_desc": md,
            "cleaned": md,
            "has_mid_format": False,
            "has_hs_code": bool(fallback_hs),
            "mid_modified": False,
        }

    # ── HS Code ──────────────────────────────────────────────────────────────
    # product_desc default is whatever followed the second slash (if present)
    product_desc = raw_desc

    # If raw_hs_field doesn't start with a digit, it's not an HS code —
    # it's the product description (no HS code in Manifested Description)
    if raw_hs_field and not re.match(r"^\s*\d", raw_hs_field):
        product_desc = (raw_hs_field + " " + product_desc).strip() if product_desc else raw_hs_field
        raw_hs_field = ""

    hs_code = ""

    # If the HS field contains both HS digits and trailing product text
    # e.g. "4203101000 UNISEX SHORT M/O KNITTED" (no slash before product)
    # try to split the leading HS part and use the remainder as product_desc.
    if raw_hs_field:
        m = re.match(r"^\s*([\d\.\s]{10,15})(.*)$", raw_hs_field)
        if m:
            candidate = re.sub(r"[\.\s]", "", m.group(1))
            if re.fullmatch(r"\d{10}", candidate):
                hs_code = candidate
                tail = m.group(2).strip()
                if tail:
                    # If we already had a product part after additional slashes, keep it
                    # otherwise use the tail as the product description.
                    if product_desc:
                        product_desc = (product_desc + " " + tail).strip()
                    else:
                        product_desc = tail

    # If we still don't have an HS code, try normalising the whole field
    if not hs_code and raw_hs_field:
        cand = re.sub(r"[\.\s]", "", raw_hs_field)
        if re.fullmatch(r"\d{10}", cand):
            hs_code = cand

    # Fallback to CE Item HSCode column if HS still missing
    if not hs_code and fallback_hs:
        hs_code = fallback_hs

    # Build final cleaned string without altering product text.
    # If HS code is present include it, otherwise omit the HS segment.
    if hs_code:
        if product_desc:
            cleaned = f"{mid_code}/{hs_code}/{product_desc}"
        else:
            cleaned = f"{mid_code}/{hs_code}"
    else:
        if product_desc:
            cleaned = f"{mid_code}/{product_desc}"
        else:
            cleaned = mid_code

    cleaned = str(cleaned).strip()

    # Was the MID code itself structurally modified? (MID: prefix, SKT→SIA, PK prepend — NOT digit/non-letter stripping)
    raw_upper = str(raw_mid).strip().upper()
    raw_letters = re.sub(r"[^A-Za-z]", "", raw_upper)
    had_mid_removed = bool(mid_code) and ("MID" in raw_upper) and ("MID" not in mid_code)
    had_skt_to_sia = bool(mid_code) and raw_letters.endswith("SKT") and mid_code.endswith("SIA")
    had_pk_prepend = bool(mid_code) and not raw_letters.startswith("PK") and mid_code.startswith("PK")
    mid_modified = had_mid_removed or had_skt_to_sia or had_pk_prepend

    return {
        "mid_code": mid_code,
        "hs_code": hs_code,
        "product_desc": product_desc,
        "cleaned": cleaned,
        "has_mid_format": True,
        "has_hs_code": bool(hs_code),
        "mid_modified": mid_modified,
    }


def process_multi_item_row(manifested: str, ce_hs_code: str) -> tuple:
    """
    Some rows have multiple items separated by commas in Manifested Description.
    Process each item individually and rejoin with commas.
    Returns (cleaned_string, has_any_mid_format, has_any_hs_code, any_mid_modified)

    has_any_hs_code = True if ANY item produced a valid HS code
                      (either from description or fallback column).
    """
    items = str(manifested).split(",")
    fallback_hs = extract_first_hs_code(ce_hs_code)

    cleaned_items = []
    any_mid_found = False
    any_hs_found = False
    any_mid_modified = False
    for item in items:
        item = item.strip()
        if not item:
            continue
        result = parse_manifested_description(item, fallback_hs)
        cleaned_items.append(result["cleaned"])
        if result["has_mid_format"]:
            any_mid_found = True
        if result["has_hs_code"]:
            any_hs_found = True
        if result.get("mid_modified"):
            any_mid_modified = True

    return ", ".join(cleaned_items), any_mid_found, any_hs_found, any_mid_modified


def process_dataframe(df: pd.DataFrame, target_country: str = "US", remove_invalid: bool = True) -> tuple[pd.DataFrame, dict]:
    """
    Apply all cleaning steps to the dataframe.
    Returns (cleaned_df, stats).
    """
    stats = {
        "original_rows": len(df),
        "removed_rows": 0,
        "processed_rows": 0,
        "mid_formatted": 0,
        "invalid_kept": 0,
        "envelopes_removed": 0,
    }

    # ── STEP 0: Filter by Target Country ────────────────────────────────────
    before_country = len(df)
    if "Recip Cntry" in df.columns and target_country != "ALL":
        # Convert both to uppercase and strip whitespace for robust matching
        df = df[df["Recip Cntry"].astype(str).str.strip().str.upper() == target_country.strip().upper()]
    df = df.reset_index(drop=True)
    stats["filtered_country_rows"] = before_country - len(df)

    if df.empty:
        raise ValueError(f"No shipments found for country: {target_country}")

    # ── STEP 0.5: Remove Envelope Shipments ─────────────────────────────────
    before_env = len(df)
    if "Service Type" in df.columns:
        df = df[~df["Service Type"].astype(str).str.contains("Envelope", case=False, na=False)]
    df = df.reset_index(drop=True)
    stats["envelopes_removed"] = before_env - len(df)

    # ── STEP 1: Handle invalid rows (blank CE Commodity Description) ────────
    before_blank = len(df)
    
    # Identify invalid rows
    is_valid = df["CE Commodity Description"].notna() & (df["CE Commodity Description"].astype(str).str.strip() != "")
    
    if remove_invalid:
        df = df[is_valid]
        df = df.reset_index(drop=True)
        stats["removed_rows"] = before_blank - len(df)
    else:
        # Kept invalid rows
        stats["invalid_kept"] = before_blank - is_valid.sum()
        stats["removed_rows"] = 0
        
    stats["processed_rows"] = len(df)

    # ── STEPS 2-6: Process Manifested Description ────────────────────────────
    cleaned_values = []
    valid_indices = []
    # Flagged rows: no valid MID code, no HS code, or both missing
    no_mid_rows = []

    for idx, row in df.iterrows():
        manifested = str(row.get("Manifested Description", ""))
        ce_hs_raw = str(row.get("CE Item HSCode", ""))

        cleaned, has_mid, has_hs, mid_modified = process_multi_item_row(manifested, ce_hs_raw)

        # Flag this row if MID is missing OR HS code is missing
        if not has_mid or not has_hs:
            tracking = str(row.get("Tracking Number", "")).strip()
            # Build a human-readable reason
            reasons = []
            if not has_mid:
                reasons.append("No valid MID code")
            if not has_hs:
                reasons.append("No HS code found")
            no_mid_rows.append({
                "Tracking Number": tracking if tracking else f"Row {idx+1}",
                "Manifested Description": manifested,
                "Reason": " | ".join(reasons),
            })
        else:
            cleaned_values.append(cleaned)
            valid_indices.append(idx)
            if has_mid:
                stats["mid_formatted"] += 1

    # Keep only valid shipments in the cleaned dataframe!
    cleaned_df = df.iloc[valid_indices].copy()
    # Preserve original Manifested Description for before/after comparison
    cleaned_df["_orig_manifested"] = cleaned_df["Manifested Description"]
    cleaned_df["Manifested Description"] = cleaned_values
    cleaned_df = cleaned_df.reset_index(drop=True)

    stats["no_mid_rows"] = no_mid_rows
    stats["no_mid_count"] = len(no_mid_rows)
    
    # Update stats
    stats["processed_rows"] = len(cleaned_df)
    stats["removed_rows"] += len(no_mid_rows)  # Add excluded shipments to removed_rows count
    
    return cleaned_df, stats


def read_uploaded_file(file_path: str, filename: str) -> pd.DataFrame:
    """Read CSV or Excel file into a DataFrame."""
    ext = filename.rsplit(".", 1)[1].lower()

    # Try multiple encodings for CSV
    if ext == "csv":
        encodings = ["utf-16", "utf-8-sig", "utf-8", "latin-1", "cp1252"]
        separators = ["\t", ",", ";"]
        for enc in encodings:
            for sep in separators:
                try:
                    df = pd.read_csv(file_path, encoding=enc, sep=sep)
                    if len(df.columns) > 1:
                        return df
                except Exception:
                    continue
        # Last resort
        return pd.read_csv(file_path, encoding="utf-16", sep="\t")
    else:
        return pd.read_excel(file_path, engine="openpyxl" if ext == "xlsx" else "xlrd")


def validate_dataframe(df: pd.DataFrame) -> list[str]:
    """Return a list of validation error messages (empty = OK)."""
    errors = []
    if df.empty:
        errors.append("The uploaded file is empty.")
        return errors

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        errors.append(
            f"Missing required column(s): {', '.join(missing)}. "
            f"Available columns: {', '.join(df.columns.tolist())}"
        )
    return errors


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    """Analyze the uploaded file and return stats before processing."""
    target_country = request.form.get("country", "US").strip().upper()

    if "file" not in request.files:
        return jsonify({"success": False, "error": "No file part."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"success": False, "error": "No file selected."}), 400

    if not allowed_file(file.filename):
        return jsonify({"success": False, "error": "Invalid file type."}), 400

    filename = secure_filename(file.filename)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    saved_name = f"analyze_{timestamp}_{filename}"
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], saved_name)

    try:
        file.save(file_path)
        df = read_uploaded_file(file_path, filename)
    except Exception as e:
        logger.exception("Failed to read uploaded file for analysis.")
        return jsonify({"success": False, "error": f"Could not read file: {str(e)}"}), 400
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)

    errors = validate_dataframe(df)
    if errors:
        return jsonify({"success": False, "error": " | ".join(errors)}), 422

    # Filter by country to get accurate preview
    if "Recip Cntry" in df.columns and target_country != "ALL":
        df = df[df["Recip Cntry"].astype(str).str.strip().str.upper() == target_country]
    
    if df.empty:
        return jsonify({"success": False, "error": f"No shipments found for country: {target_country}"}), 422

    # Filter out envelopes for the preview counts
    if "Service Type" in df.columns:
        is_envelope = df["Service Type"].astype(str).str.contains("Envelope", case=False, na=False)
        envelope_count = int(is_envelope.sum())
        df = df[~is_envelope]
    else:
        envelope_count = 0

    # Remove rows with blank 'CE Commodity Description' from the preview
    # (these will still be removed during processing as well)
    before_blank = len(df)
    is_not_blank = df["CE Commodity Description"].notna() & (df["CE Commodity Description"].astype(str).str.strip() != "")
    blank_removed = int((~is_not_blank).sum())
    df = df[is_not_blank].reset_index(drop=True)

    total_country_rows = len(df)

    # Pre-calculate MID/HS flags and blank descriptions
    has_mid_list = []
    has_hs_list = []
    blank_desc_list = []
    cleaned_desc_list = []
    changed_list = []
    spaces_found_list = []
    spaces_detail_list = []
    mid_modified_list = []
    # Preserve original manifested descriptions so UI can show before/after
    orig_manifest_list = df["Manifested Description"].fillna("").astype(str).tolist()
    
    for idx, row in df.iterrows():
        # Use the original value we preserved earlier
        manifested = orig_manifest_list[idx] if idx < len(orig_manifest_list) else str(row.get("Manifested Description", ""))
        ce_hs_raw = str(row.get("CE Item HSCode", ""))
        
        is_blank = pd.isna(row.get("CE Commodity Description")) or str(row.get("CE Commodity Description")).strip() == ""
        blank_desc_list.append(is_blank)

        cleaned, has_mid, has_hs, mid_modified = process_multi_item_row(manifested, ce_hs_raw)
        has_mid_list.append(has_mid)
        has_hs_list.append(has_hs)
        cleaned_desc_list.append(cleaned)
        # Track whether MID code was structurally modified
        changed_list.append(str(cleaned).strip() != str(manifested).strip())
        mid_modified_list.append(mid_modified)

        # Detect spacing issues — only MID/HS/slash areas, NOT product description
        spaces_issues = []
        raw = str(manifested)
        if raw.strip() != raw:
            spaces_issues.append("Leading/trailing space")
        if "/" in raw:
            parts = raw.split("/")
            mid_seg = parts[0].strip()
            if re.search(r"\s", mid_seg):
                spaces_issues.append("Space inside MID segment")
            # Check around each slash
            for i in range(len(parts) - 1):
                before_slash = parts[i]
                after_slash = parts[i + 1]
                if re.search(r"\s$", before_slash):
                    spaces_issues.append(f"Space before slash #{i+1}")
                if re.search(r"^\s", after_slash):
                    spaces_issues.append(f"Space after slash #{i+1}")

        if spaces_issues and has_mid and has_hs:
            spaces_found_list.append(True)
            spaces_detail_list.append("; ".join(spaces_issues))
        else:
            spaces_found_list.append(False)
            spaces_detail_list.append("")
        
    df = df.copy()  # Avoid SettingWithCopyWarning
    df["_has_mid"] = has_mid_list
    df["_has_hs"] = has_hs_list
    df["_is_blank_desc"] = blank_desc_list
    df["_spaces_found"] = spaces_found_list
    df["_mid_modified"] = mid_modified_list
    df["_spaces_detail"] = spaces_detail_list
    # Save cleaned description for UI, and preserve original in a separate column
    df["Manifested Description"] = cleaned_desc_list  # Cleaned description saved!
    df["_orig_manifested"] = orig_manifest_list
    df["_changed"] = changed_list
    
    # Calculate valid vs invalid based on actual MID & HS presence
    invalid_count = int(((~df["_has_mid"]) | (~df["_has_hs"])).sum())
    valid_count = total_country_rows - invalid_count

    # perfect_count = already clean rows (no change needed) — no MID mod AND no spacing issues
    perfect_count = int((df["_has_mid"] & df["_has_hs"] & (df["_mid_modified"] == False) & (df["_spaces_found"] == False)).sum())
    # modified_count (MID Words) = rows where the MID code itself was structurally modified
    modified_count = int((df["_has_mid"] & df["_has_hs"] & (df["_mid_modified"] == True)).sum())

    # Generate a session ID and store dataframe
    session_id = str(uuid.uuid4())
    SESSIONS[session_id] = {
        "df": df,
        "country": target_country,
        "filename": filename,
        "timestamp": datetime.now()
    }

    # Get sample data for the initial "Valid MID" view (unchanged rows — no MID mod AND no spaces)
    cols_to_preview = [c for c in ["Tracking Number", "Manifested Description", "CE Item HSCode", "Spaces Detail"] if c in df.columns]
    valid_df = df[df["_has_mid"] & df["_has_hs"]]
    unchanged_valid = valid_df[(valid_df["_mid_modified"] == False) & (valid_df["_spaces_found"] == False)]
    sample_data = unchanged_valid[cols_to_preview].head(20).fillna("").to_dict(orient="records")
    # Fixed (MID Words) shipments — rows where MID code was structurally modified
    fixed_df = valid_df[valid_df["_mid_modified"] == True]
    fixed_sample = []
    spaces_sample = []
    if not fixed_df.empty:
        # Build a sample showing original vs cleaned manifested description
        for _, r in fixed_df.head(20).iterrows():
            # Sanitize fields to be JSON serializable (no NaN)
            tn = r.get("Tracking Number", "")
            orig = r.get("_orig_manifested", "")
            cleaned_val = r.get("Manifested Description", "")
            hs_raw = r.get("CE Item HSCode", "") if "CE Item HSCode" in r.index else ""

            # Normalize NaN to empty string
            if pd.isna(tn):
                tn = ""
            if pd.isna(orig):
                orig = ""
            if pd.isna(cleaned_val):
                cleaned_val = ""
            if pd.isna(hs_raw):
                hs_raw = ""

            fixed_sample.append({
                "Tracking Number": str(tn),
                "original_manifested": str(orig),
                "cleaned_manifested": str(cleaned_val),
                "CE Item HSCode": str(hs_raw),
            })

        # Spaces sample (rows with spacing issues)
        spaces_sample = []
        try:
            spaces_df = df[df["_spaces_found"] == True]
            if not spaces_df.empty:
                for _, r in spaces_df.head(20).iterrows():
                    tn = r.get("Tracking Number", "")
                    mani = r.get("_orig_manifested", r.get("Manifested Description", ""))
                    detail = r.get("_spaces_detail", "")
                    if pd.isna(tn): tn = ""
                    if pd.isna(mani): mani = ""
                    if pd.isna(detail): detail = ""
                    spaces_sample.append({
                        "Tracking Number": str(tn),
                        "Manifested Description": str(mani),
                        "Spaces Detail": str(detail),
                    })
        except Exception:
            spaces_sample = []
    
    return jsonify({
        "success": True,
        "preview": {
            "session_id": session_id,
            "country": target_country,
            # Report the total number of shipments after country/envelope filtering
            # so the UI can show overall file size vs valid/invalid breakdown.
            "total_rows": total_country_rows,
            "blank_removed": blank_removed,
            "spaces_count": int(df["_spaces_found"].sum()) if "_spaces_found" in df.columns else 0,
            "valid_rows": valid_count,
            "perfect_count": perfect_count,
            "modified_count": modified_count,
            "invalid_rows": invalid_count,
            "fixed_count": int(len(fixed_df)) if 'fixed_df' in locals() else 0,
            "envelope_count": envelope_count,
            "sample_data": sample_data,
            "fixed_sample": fixed_sample,
            "spaces_sample": spaces_sample,
            "modified_sample": fixed_sample
        }
    })

@app.route("/api/preview/<session_id>")
def preview_data(session_id):
    """Return paginated data based on type (all, valid, invalid)."""
    if session_id not in SESSIONS:
        return jsonify({"success": False, "error": "Session expired or not found."}), 404
    
    session = SESSIONS[session_id]
    df = session["df"]
    
    view_type = request.args.get("type", "all")
    page = int(request.args.get("page", 1))
    per_page = int(request.args.get("per_page", 20))
    process = request.args.get("process", "false").lower() == "true"
    
    # Filter envelopes
    if "Service Type" in df.columns:
        is_envelope = df["Service Type"].astype(str).str.contains("Envelope", case=False, na=False)
        df = df[~is_envelope]
        
    if view_type == "all":
        df = df[df["_has_mid"] & df["_has_hs"]]
    elif view_type == "valid":
        df = df[df["_has_mid"] & df["_has_hs"] & (df["_changed"] == False)]
    elif view_type == "invalid":
        df = df[(~df["_has_mid"]) | (~df["_has_hs"])]
    elif view_type == "spaces":
        if "_spaces_found" in df.columns:
            df = df[df["_spaces_found"] == True]
        else:
            df = df.iloc[0:0]
        df = df.reset_index(drop=True).copy()
        if process:
            df["Original Manifested Description"] = df["_orig_manifested"].astype(str)
            cleaned = df["_orig_manifested"].astype(str).str.strip()
            cleaned = cleaned.str.replace(r"\s{2,}", " ", regex=True)
            df["Cleaned Manifested Description"] = cleaned
        else:
            df["Spaces Detail"] = df.get("_spaces_detail", "")
    elif view_type == "modified":
        df = df[df["_has_mid"] & df["_has_hs"] & (df["_mid_modified"] == True)]
        if process:
            df = df.copy()
            df["Original Manifested Description"] = df["_orig_manifested"].astype(str)
            df["Cleaned Manifested Description"] = df["Manifested Description"].astype(str)
        
    total_records = len(df)
    
    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page
    
    if view_type == "spaces" and process:
        cols_to_preview = [c for c in ["Tracking Number", "Original Manifested Description", "Cleaned Manifested Description"] if c in df.columns]
        temp_df = df[cols_to_preview].iloc[start_idx:end_idx].fillna("").copy()
        page_data = temp_df.to_dict(orient="records")
        return jsonify({"success": True, "data": page_data, "total": total_records, "page": page, "per_page": per_page})
    elif view_type == "modified":
        if process:
            cols_to_preview = [c for c in ["Tracking Number", "Original Manifested Description", "Cleaned Manifested Description"] if c in df.columns]
        else:
            cols_to_preview = [c for c in ["Tracking Number", "_orig_manifested"] if c in df.columns]
        temp_df = df[cols_to_preview].iloc[start_idx:end_idx].fillna("").copy()
        if process:
            page_data = temp_df.to_dict(orient="records")
        else:
            temp_df.columns = ["Tracking Number", "Original Manifested Description"]
            page_data = temp_df.to_dict(orient="records")
        return jsonify({"success": True, "data": page_data, "total": total_records, "page": page, "per_page": per_page})
    elif view_type == "spaces":
        cols_to_preview = [c for c in ["Tracking Number", "Manifested Description", "Spaces Detail"] if c in df.columns]
    else:
        cols_to_preview = [c for c in ["Tracking Number", "Manifested Description", "CE Item HSCode"] if c in df.columns]
    
    page_data = df[cols_to_preview].iloc[start_idx:end_idx].fillna("").to_dict(orient="records")
    
    return jsonify({
        "success": True,
        "data": page_data,
        "total": total_records,
        "page": page,
        "per_page": per_page
    })

@app.route("/api/download_invalid/<session_id>")
def download_invalid(session_id):
    """Download an Excel file of ONLY the invalid rows."""
    if session_id not in SESSIONS:
        return "Session expired or not found.", 404
        
    session = SESSIONS[session_id]
    df = session["df"]
    
    invalid_df = df[(~df["_has_mid"]) | (~df["_has_hs"])].copy()
    
    cols_to_drop = [c for c in ["_has_mid", "_has_hs", "_is_blank_desc"] if c in invalid_df.columns]
    if cols_to_drop:
        invalid_df = invalid_df.drop(columns=cols_to_drop)

    # Convert Tracking Number to string to prevent scientific notation in Excel
    if "Tracking Number" in invalid_df.columns:
        invalid_df["Tracking Number"] = invalid_df["Tracking Number"].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() not in ("", "nan") else str(x)
        )
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        invalid_df.to_excel(writer, index=False, sheet_name="Invalid Data")
    output.seek(0)
    
    filename = f"Invalid_Shipments_{session['filename']}.xlsx"
    
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/download_spaces/<session_id>")
def download_spaces(session_id):
    """Download an Excel file of ONLY the rows with spacing issues."""
    if session_id not in SESSIONS:
        return "Session expired or not found.", 404
        
    session = SESSIONS[session_id]
    df = session["df"]
    
    if "_spaces_found" in df.columns:
        spaces_df = df[df["_spaces_found"] == True].copy()
    else:
        spaces_df = df.iloc[0:0].copy()
    
    # Clean up Manifested Description (trim spaces)
    if "Manifested Description" in spaces_df.columns:
        spaces_df["Manifested Description"] = spaces_df["Manifested Description"].astype(str).str.strip()
        spaces_df["Manifested Description"] = spaces_df["Manifested Description"].str.replace(r"\s{2,}", " ", regex=True)
    
    # Drop temp columns
    cols_to_drop = [c for c in ["_has_mid", "_has_hs", "_is_blank_desc", "_spaces_found", "_spaces_detail", "_orig_manifested", "_changed"] if c in spaces_df.columns]
    if cols_to_drop:
        spaces_df = spaces_df.drop(columns=cols_to_drop)

    if "Tracking Number" in spaces_df.columns:
        spaces_df["Tracking Number"] = spaces_df["Tracking Number"].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() not in ("", "nan") else str(x)
        )
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        spaces_df.to_excel(writer, index=False, sheet_name="Spaces Fixed Data")
    output.seek(0)
    
    filename = f"Spaces_Fixed_{session['filename']}.xlsx"
    
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/download_modified/<session_id>")
def download_modified(session_id):
    """Download an Excel file of ONLY the rows modified/cleaned by the tool."""
    if session_id not in SESSIONS:
        return "Session expired or not found.", 404
        
    session = SESSIONS[session_id]
    df = session["df"]
    
    modified_df = df[df["_has_mid"] & df["_has_hs"] & (df["_changed"] == True)].copy()
    
    cols_to_drop = [c for c in ["_has_mid", "_has_hs", "_is_blank_desc", "_spaces_found", "_spaces_detail", "_orig_manifested", "_changed"] if c in modified_df.columns]
    if cols_to_drop:
        modified_df = modified_df.drop(columns=cols_to_drop)

    if "Tracking Number" in modified_df.columns:
        modified_df["Tracking Number"] = modified_df["Tracking Number"].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() not in ("", "nan") else str(x)
        )
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        modified_df.to_excel(writer, index=False, sheet_name="Modified Data")
    output.seek(0)
    
    filename = f"Modified_Shipments_{session['filename']}.xlsx"
    
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/download_by_type/<session_id>")
def download_by_type(session_id):
    """Download data based on active preview tab type — one unified endpoint."""
    if session_id not in SESSIONS:
        return "Session expired or not found.", 404

    session = SESSIONS[session_id]
    df = session["df"]
    download_type = request.args.get("type", "total")

    # ── Filter by type ─────────────────────────────────────────────────────
    if download_type == "total":
        filtered = df[df["_has_mid"] & df["_has_hs"]].copy()
    elif download_type == "valid":
        filtered = df[df["_has_mid"] & df["_has_hs"] & (df["_changed"] == False)].copy()
    elif download_type == "invalid":
        filtered = df[(~df["_has_mid"]) | (~df["_has_hs"])].copy()
    elif download_type == "envelope":
        if "Service Type" in df.columns:
            filtered = df[df["Service Type"].astype(str).str.contains("Envelope", case=False, na=False)].copy()
        else:
            filtered = df.iloc[0:0].copy()
    elif download_type == "spaces":
        if "_spaces_found" in df.columns:
            filtered = df[df["_spaces_found"] == True].copy()
        else:
            filtered = df.iloc[0:0].copy()
        # Process: create Original & Cleaned columns by trimming spaces
        if "Manifested Description" in filtered.columns:
            filtered["Original Manifested Description"] = filtered["_orig_manifested"].astype(str)
            # Actually trim spaces from original
            cleaned_spaces = filtered["_orig_manifested"].astype(str).str.strip()
            cleaned_spaces = cleaned_spaces.str.replace(r"\s{2,}", " ", regex=True)
            filtered["Cleaned Manifested Description"] = cleaned_spaces
    elif download_type == "modified":
        filtered = df[df["_has_mid"] & df["_has_hs"] & (df["_mid_modified"] == True)].copy()
        if "_orig_manifested" in filtered.columns and "Manifested Description" in filtered.columns:
            filtered["Original Manifested Description"] = filtered["_orig_manifested"].astype(str)
            filtered["Cleaned Manifested Description"] = filtered["Manifested Description"].astype(str)
    else:
        return "Invalid download type.", 400

    # ── Drop temp columns ──────────────────────────────────────────────────
    temp_cols = [c for c in ["_has_mid", "_has_hs", "_is_blank_desc", "_spaces_found", "_spaces_detail", "_orig_manifested", "_changed", "_mid_modified"] if c in filtered.columns]
    if temp_cols:
        filtered = filtered.drop(columns=temp_cols)

    # Convert Tracking Number to string (prevent scientific notation)
    if "Tracking Number" in filtered.columns:
        filtered["Tracking Number"] = filtered["Tracking Number"].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() not in ("", "nan") else str(x)
        )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        filtered.to_excel(writer, index=False, sheet_name="Data")
    output.seek(0)

    type_labels = {
        "total": "All_Shipments", "valid": "Valid_Shipments", "invalid": "Invalid_Shipments",
        "envelope": "Envelope_Shipments", "spaces": "Spaces_Fixed", "modified": "Modified_Shipments",
    }
    label = type_labels.get(download_type, "Data")
    out_filename = f"{label}_{session['filename']}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=out_filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/process", methods=["POST"])
def process():
    """Process the file using the session data."""
    session_id = request.form.get("session_id")
    # Removal of blank CE Commodity Description is enforced by default (no user option).
    remove_invalid = True
    
    if not session_id or session_id not in SESSIONS:
        return jsonify({"success": False, "error": "Session expired or invalid. Please upload the file again."}), 400

    session = SESSIONS[session_id]
    df = session["df"]
    target_country = session["country"]

    errors = validate_dataframe(df)
    if errors:
        return jsonify({"success": False, "error": " | ".join(errors)}), 422

    try:
        cleaned_df, stats = process_dataframe(df, target_country, remove_invalid)
    except ValueError as ve:
        return jsonify({"success": False, "error": str(ve)}), 422
    except Exception as e:
        logger.exception("Processing failed.")
        return jsonify({"success": False, "error": f"Processing error: {str(e)}"}), 500

    # Rename original manifested for before/after comparison, then drop temp columns
    excel_df = cleaned_df.copy()
    if "_orig_manifested" in excel_df.columns:
        excel_df = excel_df.rename(columns={"_orig_manifested": "Original Manifested Description"})
    cols_to_drop = [c for c in ["_has_mid", "_has_hs", "_is_blank_desc"] if c in excel_df.columns]
    excel_df = excel_df.drop(columns=cols_to_drop) if cols_to_drop else excel_df

    # Convert Tracking Number to string to prevent scientific notation in Excel
    if "Tracking Number" in excel_df.columns:
        excel_df = excel_df.copy()
        excel_df["Tracking Number"] = excel_df["Tracking Number"].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() not in ("", "nan") else str(x)
        )

    # Export to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        excel_df.to_excel(writer, index=False, sheet_name="Cleaned Data")
    output.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Store in uploads folder temporarily for download
    out_filename = f"cleaned_{timestamp}.xlsx"
    out_path = os.path.join(app.config["UPLOAD_FOLDER"], out_filename)
    with open(out_path, "wb") as f:
        f.write(output.getvalue())
        
    # Clean up session
    if session_id in SESSIONS:
        del SESSIONS[session_id]

    return jsonify({
        "success": True,
        "stats": stats,
        "download_token": out_filename,
        "no_mid_rows": stats.get("no_mid_rows", []),
        "no_mid_count": stats.get("no_mid_count", 0),
        "message": (
            f"✅ Processing complete! "
            f"Filtered out {stats.get('filtered_country_rows', 0)} non-{target_country} rows. "
            f"{stats.get('envelopes_removed', 0)} envelopes removed. "
            f"{stats['removed_rows']} invalid rows removed. "
            f"{stats['processed_rows']} rows processed. "
            f"{stats['mid_formatted']} MID codes cleaned."
        ),
    })


@app.route("/download/<token>")
def download(token):
    """Serve the processed Excel file for download."""
    # Security: only allow filenames matching our pattern
    if not re.fullmatch(r"cleaned_\d{8}_\d{6}\.xlsx", token):
        return "Invalid download token.", 400

    file_path = os.path.join(app.config["UPLOAD_FOLDER"], token)
    if not os.path.exists(file_path):
        return "File not found or already downloaded.", 404

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    download_name = f"Customs_Cleaned_{timestamp}.xlsx"

    response = send_file(
        file_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Delete file after serving (cleanup)
    @response.call_on_close
    def cleanup():
        try:
            os.remove(file_path)
        except OSError:
            pass

    return response


@app.route("/health")
def health():
    return jsonify({"status": "ok", "version": "1.0.0"})


# ── Entry Point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(host="0.0.0.0", port=port, debug=debug)
