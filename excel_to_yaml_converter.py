import pandas as pd
import yaml
import re

def clean_text(text):
    """テキストをクリーニングする"""
    if pd.isna(text):
        return ""
    text = str(text).strip()
    # 改行文字を統一
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    return text

def split_examples(text):
    """例や違反例を改行で分割してリストにする"""
    if not text:
        return []
    # 改行で分割し、空の要素を除去
    examples = [ex.strip() for ex in text.split('\n') if ex.strip()]
    return examples

def convert_excel_to_yaml(excel_file, sheet_name, output_file):
    """エクセルファイルをYAMLに変換する"""
    
    # エクセルファイルを読み込み（3行目をヘッダーとして使用）
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=2)
    
    # カラム名を設定
    df.columns = ['品質特性', '品質特性の説明', '品質副特性', '品質副特性の説明', 
                  '測定項目', '例', '違反例']
    
    # 結果を格納する辞書
    result = {'品質特性': []}
    
    current_quality_characteristic = None
    current_sub_characteristic = None
    
    for index, row in df.iterrows():
        # 各行のデータをクリーニング
        quality_char = clean_text(row['品質特性'])
        quality_char_desc = clean_text(row['品質特性の説明'])
        sub_char = clean_text(row['品質副特性'])
        sub_char_desc = clean_text(row['品質副特性の説明'])
        measurement_item = clean_text(row['測定項目'])
        examples = clean_text(row['例'])
        violation_examples = clean_text(row['違反例'])
        
        # 品質特性が新しい場合
        if quality_char and quality_char != current_quality_characteristic:
            current_quality_characteristic = quality_char
            current_sub_characteristic = None
            
            # 新しい品質特性を追加
            quality_char_dict = {
                '名称': quality_char,
                '説明': quality_char_desc,
                '副特性': []
            }
            result['品質特性'].append(quality_char_dict)
        
        # 品質副特性が新しい場合
        if sub_char and sub_char != current_sub_characteristic:
            current_sub_characteristic = sub_char
            
            # 新しい品質副特性を追加
            sub_char_dict = {
                '名称': sub_char,
                '説明': sub_char_desc,
                '測定項目': []
            }
            # 最新の品質特性に副特性を追加
            if result['品質特性']:
                result['品質特性'][-1]['副特性'].append(sub_char_dict)
        
        # 測定項目がある場合
        if measurement_item:
            measurement_dict = {
                '項目': measurement_item,
                '例': split_examples(examples),
                '違反例': split_examples(violation_examples)
            }
            
            # 副特性がある場合は副特性に、ない場合は品質特性に直接追加
            if result['品質特性'] and result['品質特性'][-1]['副特性']:
                result['品質特性'][-1]['副特性'][-1]['測定項目'].append(measurement_dict)
            elif result['品質特性']:
                # 副特性がない場合は品質特性に直接測定項目を追加
                if '測定項目' not in result['品質特性'][-1]:
                    result['品質特性'][-1]['測定項目'] = []
                result['品質特性'][-1]['測定項目'].append(measurement_dict)
    
    # YAMLファイルに出力
    with open(output_file, 'w', encoding='utf-8') as f:
        yaml.dump(result, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
    
    print(f"YAMLファイル '{output_file}' が作成されました。")
    return result

if __name__ == "__main__":
    # 変換実行
    excel_file = "ASDoQ_SystemDocumentationQualityModel_v2.0a.xlsx"
    sheet_name = "品質特性・副特性・測定項目（例・違反例を含む）"
    output_file = "quality_model.yaml"
    
    result = convert_excel_to_yaml(excel_file, sheet_name, output_file)
    
    # 変換結果の概要を表示
    print(f"\n変換結果概要:")
    print(f"品質特性数: {len(result['品質特性'])}")
    for i, qc in enumerate(result['品質特性'], 1):
        sub_count = len(qc.get('副特性', []))
        direct_measurement_count = len(qc.get('測定項目', []))
        print(f"  {i}. {qc['名称']} - 副特性: {sub_count}個, 直接測定項目: {direct_measurement_count}個") 