import yaml
import pandas as pd

def flatten_quality_model(yaml_file, csv_file):
    with open(yaml_file, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    rows = []
    for qc in data['品質特性']:
        qc_name = qc.get('名称', '')
        qc_desc = qc.get('説明', '')
        sub_list = qc.get('副特性', [])
        # 品質特性直下の測定項目（副特性がない場合）
        for m in qc.get('測定項目', []):
            if not m.get('項目', '').strip():
                continue
            row = {
                '品質特性': qc_name,
                '品質特性の説明': qc_desc,
                '品質副特性': '',
                '品質副特性の説明': '',
                '測定項目': m.get('項目', ''),
                '例': '\n'.join(m.get('例', [])) if m.get('例', []) else '',
                '違反例': '\n'.join(m.get('違反例', [])) if m.get('違反例', []) else '',
            }
            rows.append(row)
        for sub in sub_list:
            sub_name = sub.get('名称', '')
            sub_desc = sub.get('説明', '')
            for m in sub.get('測定項目', []):
                if not m.get('項目', '').strip():
                    continue
                row = {
                    '品質特性': qc_name,
                    '品質特性の説明': qc_desc,
                    '品質副特性': sub_name,
                    '品質副特性の説明': sub_desc,
                    '測定項目': m.get('項目', ''),
                    '例': '\n'.join(m.get('例', [])) if m.get('例', []) else '',
                    '違反例': '\n'.join(m.get('違反例', [])) if m.get('違反例', []) else '',
                }
                rows.append(row)
    # マージセル風：同じ値が連続する場合は最初の1行だけ値を出力し、以降は空欄
    prev = {col: None for col in ['品質特性', '品質特性の説明', '品質副特性', '品質副特性の説明']}
    for row in rows:
        for col in prev.keys():
            if row[col] == prev[col]:
                row[col] = ''
            else:
                prev[col] = row[col]
    df = pd.DataFrame(rows, columns=[
        '品質特性', '品質特性の説明', '品質副特性', '品質副特性の説明', '測定項目', '例', '違反例'
    ])
    df.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"CSVファイル '{csv_file}' を出力しました。")

if __name__ == '__main__':
    flatten_quality_model('QualityModel_V2.YAML', 'QualityModel_V2_fromYAML.csv') 