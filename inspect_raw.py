#!/usr/bin/env python3
import requests, io, openpyxl

DROPBOX_TOKEN = "sl.u.AGgF5XKR5BwOCmo95YjLwdRE-buSMDIWiOPYXcWcE7SUTTaHr9AIpXz-9MVh42IhLVZsHJCrKnkxOwzu8i5hgLrIuJC8XlRGaCprnDHwr1Czw7hxc92BXwjNMEAMvR0iXS88gLeyV_v4Wbj4KCTCMX7qlmtztJWZSS1oN0mfnHX2dj2SP2A_4Ari0pVrlMqbUbnWlRGYRNgHrMA7K6bbUCfeBq-3IO8VaTTkZEnyYNMKPHUHu1CHyxypSrC44D5jlBEyPLguvyKppj_qwNQkfJSHHhLv-VvqohTPLRMewH0inx-UpIcq7ku28sGrd04g2-R2YJUOcEbAWOgu5Uwm2VCXXwTxTlbBPV2NQe--n_ZT_uzACDQHLV0E9c0FIA9Z8_v56DGJch7mMmR7TbavjM9v2yFDzf8Jdw4ip3LCYS1yhkHwV-_CTZ8fFdTcJJ5gm4CbjCGhI-k9eGvmp4xFeId6Wra8yWLA_DZd-Q11gHbb-FY7vCD7hWvNi0B_3N86FG4S1ha_0loOmJkav9dk6L81RhIcAzLiMsN64rO0hTS4gugUGT8MuXm9UbAwNL_7untY5spa_8ecfefeC894o-Ae3TVeAji3dvcrUYsgClYgK1fkylqt1SaeOj1K0PZFcjmxin-K9BUvfGTjUclowh-HxEDl7n6-QofS1DqczB0loetZV27LHZFoBGl4FDRpSDhydxE9XkwyoIXg6LSfUUNCwi9fFlBnISrO5eZaGrCjpqpftgEug_jPjRLSILFra6uakJOVT8gS-aZoWBZv9YacSnrJ6iuRf0mRo6X7nV9SOm2cmdtyUZvdxBxtjTfaQfFWD8VkHHmTHLR-Fmt50on5DDaocHhfzkJt8wyafU4YH7NSbz2mhfFmxcb1s5WEQrgslgW5dKmmi6VGBIcPDMeIPxoRZe43aodjcWYzZmOT7czMjPpI5J6CNalK8HZMI43oRVwKAml3TXMEg0CLhtlMRdsafEo3rIdoB81ZK3RbnnGTCXkXuXKDdzd1qHRyE3Z2DCbcHFYvwzYcT6gr4VnIc4WmsTt7f4s4bGoaHoJxD8K1zOn85J_VLHJBOXdknyDSyHxhNPSKebKo3_pAFXJe4Q4HoaCrGyxHH09xGS_1TteDVPnHqRUFK5fIa1EezABLZfDKnvKOBxLeaTe9AEO89ddAiYvpJaGkpCKKG0GIOtivpsNRwi4tczW06bmZZAUPLQFjkWhTcxvh_7QHYG_Jm63xC1B7st_7Zkez4HnAzUsTCEucTpIpTWvwAcEuu-5cgRUdbe2XHjeSeeFUUmPB"
DROPBOX_PATH = "/Previfuego/MATRIZ LOCALES/MATRIZ EXTINTORES GRUPO KFC.xlsx"

resp = requests.post(
    "https://content.dropboxapi.com/2/files/download",
    headers={
        "Authorization": f"Bearer {DROPBOX_TOKEN}",
        "Dropbox-API-Arg": f'{{"path": "{DROPBOX_PATH}"}}',
    }, timeout=60,
)
resp.raise_for_status()

# Load WITHOUT data_only to see actual values
wb = openpyxl.load_workbook(io.BytesIO(resp.content), data_only=True)
ws = wb["ENERO"]

print("=== ENERO first 30 rows (all cols 0-11) ===")
for i, row in enumerate(ws.iter_rows(values_only=True, max_row=30)):
    print(f"Row {i}: {row[:12]}")

print("\n=== Checking merged cells in ENERO ===")
print(ws.merged_cells)

# Check CAPACIDAD column values (col index 3)
print("\n=== Unique CAPACIDAD values in ENERO ===")
caps = set()
for row in ws.iter_rows(values_only=True, min_row=2):
    if row and len(row) > 3 and row[3] is not None:
        caps.add(repr(row[3]))
print(sorted(caps))

# Check TIPO values
print("\n=== Unique TIPO values in ENERO ===")
tipos = set()
for row in ws.iter_rows(values_only=True, min_row=2):
    if row and len(row) > 2 and row[2] is not None:
        tipos.add(repr(row[2]))
print(sorted(tipos))

# Check CC values
print("\n=== Unique CC values in ENERO ===")
ccs = set()
for row in ws.iter_rows(values_only=True, min_row=2):
    if row and len(row) > 0 and row[0] is not None:
        ccs.add(repr(row[0]))
print(sorted(ccs)[:30])
