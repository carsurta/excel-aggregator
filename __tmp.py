from pathlib import Path
from header_multirow import load_sheet_merge_aware, detect_header_band_and_build
base = Path('data')
df = load_sheet_merge_aware(str(base/'시험데이터_3.xlsx'), '집계표')
start, end, _ = detect_header_band_and_build(df)
print('detected', start+1, end+1)
