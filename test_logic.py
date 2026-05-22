import sys
sys.path.insert(0, '.')
from app import clean_mid_code, parse_manifested_description, extract_first_hs_code, strip_known_prefixes

tests = [
    # ── Original tests (must still pass) ─────────────────────────────────────
    ('PKUMANAESIA51310 / 6204120010 / MEN SUIT M/O 100',  '1|6204120010', 'PKUMANAESIA/6204120010/MEN SUIT M/O 100'),
    ('PKEXTVAG95KAR/ 8906.29.0000/ UNISEX BRASS METAL CH', '1|8906290000', 'PKEXTVAGKAR/8906290000/UNISEX BRASS METAL CH'),
    ('PKSENINDSIA/6110303059/UNISEX FOOTBALL JERSEYS M/O', '',             'PKSENINDSIA/6110303059/UNISEX FOOTBALL JERSEYS M/O'),
    ('PKPLASHASIA51310 / 4203104060 / MEN FASHION JACKET', '1|4203104060', 'PKPLASHASIA/4203104060/MEN FASHION JACKET'),
    ('PKLENSUR5131SIA/8203.20 0000/ 2 PCS HAIR EXTENSION', '1|8203200000', 'PKLENSURSIA/8203200000/2 PCS HAIR EXTENSION'),

    # ── NEW: MID: prefix removal ──────────────────────────────────────────────
    ('MID:PKKTSTSIA/6566322100/STOLEM/O100 POLYSTER FOR', '', 'PKKTSTSIA/6566322100/STOLEM/O100 POLYSTER FOR'),
    ('MID: PKSENINDSIA/6110303059/MEN JACKET',            '', 'PKSENINDSIA/6110303059/MEN JACKET'),

    # ── NEW: PSWSHIPMENT: prefix removal ─────────────────────────────────────
    ('PSWSHIPMENT:PKBLUICESIA/6107299000/UNISEX SHIRTS M/O 100 POLYESTER', '', 'PKBLUICESIA/6107299000/UNISEX SHIRTS M/O 100 POLYESTER'),
    ('PSW SHIPMENT:PKBLUICESIA/6107299000/SPORTS BAGS',                   '', 'PKBLUICESIA/6107299000/SPORTS BAGS'),

    # ── NEW: SKT → SIA replacement ────────────────────────────────────────────
    ('PKAXESPOWEASKT/6505006090/MEN CAP M/O 100 POLYESTER',  '', 'PKAXESPOWEASIA/6505006090/MEN CAP M/O 100 POLYESTER'),
    ('PKAXESPOWEASKT/6110303059/MEN HOODIES M/O 100 POLYESTER', '', 'PKAXESPOWEASIA/6110303059/MEN HOODIES M/O 100 POLYESTER'),

    # ── KAR (Karachi) must NOT be changed ─────────────────────────────────────
    ('PKEXTVAG95KAR/8906290000/BRASS METAL',              '', 'PKEXTVAGKAR/8906290000/BRASS METAL'),
    
    # ── NEW: Dot, Bracket, Space, Colon parsing ───────────────────────────────
    ('PKMANASIA.6204120010.MEN SUIT.COM',                 '', 'PKMANASIA/6204120010/MEN SUIT.COM'),
    ('PKARDCRASIA(4205000500) COWHIDE LEATHER',           '', 'PKARDCRASIA/4205000500/COWHIDE LEATHER'),
    ('PKSPOENT513SIA 6104630000 JERSEY',                  '', 'PKSPOENTSIA/6104630000/JERSEY'),
    ('PKMANASIA:6204120010:MEN SUIT',                     '', 'PKMANASIA/6204120010/MEN SUIT'),
    
    # ── NEW: MID keyword removal & PK prepend ─────────────────────────────────
    ('PKMIDMANASIA/1234567890/DESC',                      '', 'PKMANASIA/1234567890/DESC'),
    ('MID:MANASIA/1234567890/DESC',                       '', 'PKMANASIA/1234567890/DESC'),
    ('MIDUMANAESIA/1234567890/DESC',                      '', 'PKUMANAESIA/1234567890/DESC'),
    
    # ── NEW: Missing MIDs & NIC ───────────────────────────────────────────────
    ('1234567890/DESC',                                   '', '1234567890/DESC'),
    ('NIC 12345-123',                                     '', 'NIC 12345-123'),

    # ── BUG FIX: Product description in second part (no HS code in MD) ─────────
    # jab raw_hs_field digit se start nahi hota, to wo product description hai
    ('PKUNIEXSIA/UNISEX SHORTS M/O KNITED',               '1|6204120010', 'PKUNIEXSIA/6204120010/UNISEX SHORTS M/O KNITED'),
    ('PKTESTABC/MEN COTTON T-SHIRT 100',                  '1|6109100010', 'PKTESTABC/6109100010/MEN COTTON T-SHIRT 100'),
    ('PKSIMSIA/WOMEN DRESS M/O POLYESTER',                '',             'PKSIMSIA/WOMEN DRESS M/O POLYESTER'),
]

print('=== Processing Logic Tests ===')
all_pass = True
for md, hs_raw, expected in tests:
    hs_fallback = extract_first_hs_code(hs_raw)
    result = parse_manifested_description(md, hs_fallback)
    got = result['cleaned']
    status = 'PASS' if got == expected else 'FAIL'
    if status == 'FAIL':
        all_pass = False
    print('[' + status + '] Input : ' + md[:60])
    print('       Output: ' + got)
    if status == 'FAIL':
        print('       Expect: ' + expected)
    print()

# ── prefix stripping unit tests ───────────────────────────────────────────────
print('=== Prefix Strip Tests ===')
prefix_tests = [
    ('PSWSHIPMENT:PKTEST/123/DESC',       'PKTEST/123/DESC'),
    ('PSW SHIPMENT:PKTEST/123/DESC',      'PKTEST/123/DESC'),
    ('PSW-SHIPMENT:PKTEST/123/DESC',      'PKTEST/123/DESC'),
    ('PKTEST/123/DESC',                   'PKTEST/123/DESC'),  # no prefix = unchanged
]
for raw, expected in prefix_tests:
    got = strip_known_prefixes(raw)
    status = 'PASS' if got == expected else 'FAIL'
    if status == 'FAIL':
        all_pass = False
    print('[' + status + '] ' + raw[:55] + '  ->  ' + got)

print()
if all_pass:
    print('All tests passed!')
else:
    print('Some tests FAILED.')
