"""
ASDoQ 文書品質モデル 変換ユーティリティ

このスクリプトは以下の変換を自動化します。
- Excel(品質モデル/用語集) -> YAML
- YAML(品質モデル) -> CSV

特にCSV生成時は、元Excelの「セル結合(表示上のマージ)」を再現するため、
同じグループ内の2行目以降のラベル列を空欄にします。
"""

import pandas as pd
import yaml
import re

def clean_text(text):
	"""文字列をトリミングし、改行コードを統一する。

	引数:
		text: 任意の型。NaNやNone、数値を含む可能性がある値
	戻り値:
		str: 先頭末尾空白を除去し、\r\n/\r を \n に統一した文字列。NaNは空文字。
	"""
	if pd.isna(text):
		return ""
	text = str(text).strip()
	# 改行コードを統一
	text = text.replace('\r\n', '\n').replace('\r', '\n')
	return text

def split_examples(text):
	"""例/違反例セルの文字列を改行で分割し、空行を除去したリストにする。

	引数:
		text: str | None
	戻り値:
		list[str]: 空でない行のみからなるリスト
	"""
	if not text:
		return []
	# 改行で分割し、空の要素を除去
	examples = [ex.strip() for ex in text.split('\n') if ex.strip()]
	return examples

def convert_excel_to_yaml(excel_file, sheet_name, output_file):
	"""品質モデルのExcelをYAMLへ変換する。

	前提:
	- 指定シートの1行目がヘッダ
	- カラム構成を固定(品質特性/説明/副特性/説明/測定項目/例/違反例)

	引数:
		excel_file: 変換元のExcelファイルパス
		sheet_name: 対象シート名
		output_file: 出力YAMLファイルパス
	戻り値:
		dict: YAMLへ書き出した品質モデルデータ(辞書構造)
	"""
	# Excel読み込み
	df = pd.read_excel(excel_file, sheet_name=sheet_name, header=0)
	# 以降の処理で参照しやすいようにカラム名を固定
	df.columns = ['品質特性', '品質特性の説明', '品質副特性', '品質副特性の説明', 
				  '測定項目', '例', '違反例']
	# 出力のベース構造
	result = {'品質特性': []}
	# 現在の走査位置における「直近の品質特性/副特性」を保持
	current_quality_characteristic = None
	current_sub_characteristic = None
	# 行順で走査し、階層構造を構築
	for index, row in df.iterrows():
		# 各セルを正規化
		quality_char = clean_text(row['品質特性'])
		quality_char_desc = clean_text(row['品質特性の説明'])
		sub_char = clean_text(row['品質副特性'])
		sub_char_desc = clean_text(row['品質副特性の説明'])
		measurement_item = clean_text(row['測定項目'])
		examples = clean_text(row['例'])
		violation_examples = clean_text(row['違反例'])
		# 品質特性の切替検知
		if quality_char and quality_char != current_quality_characteristic:
			current_quality_characteristic = quality_char
			current_sub_characteristic = None
			# 新しい品質特性ノードを追加
			quality_char_dict = {
				'名称': quality_char,
				'説明': quality_char_desc,
				'副特性': []
			}
			result['品質特性'].append(quality_char_dict)
		# 品質副特性の切替検知
		if sub_char and sub_char != current_sub_characteristic:
			current_sub_characteristic = sub_char
			# 新しい副特性ノードを追加
			sub_char_dict = {
				'名称': sub_char,
				'説明': sub_char_desc,
				'測定項目': []
			}
			# 直近の品質特性にぶら下げる
			if result['品質特性']:
				result['品質特性'][-1]['副特性'].append(sub_char_dict)
		# 測定項目がある場合にのみ追加
		if measurement_item:
			measurement_dict = {
				'項目': measurement_item,
				'例': split_examples(examples),
				'違反例': split_examples(violation_examples)
			}
			# 副特性が存在する場合はその末尾へ、存在しないなら品質特性直下へ
			if result['品質特性'] and result['品質特性'][-1]['副特性']:
				result['品質特性'][-1]['副特性'][-1]['測定項目'].append(measurement_dict)
			elif result['品質特性']:
				if '測定項目' not in result['品質特性'][-1]:
					result['品質特性'][-1]['測定項目'] = []
				result['品質特性'][-1]['測定項目'].append(measurement_dict)
	# YAMLへ書き出し
	with open(output_file, 'w', encoding='utf-8') as f:
		yaml.dump(result, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
	print(f"YAMLファイル '{output_file}' が作成されました。")
	return result

def convert_glossary_to_yaml(excel_file, sheet_name, output_file):
	"""用語集シートをYAMLへ変換する。

	引数:
		excel_file: 変換元Excel
		sheet_name: 用語集シート名
		output_file: 出力YAML
	戻り値:
		list[dict]: 用語集エントリのリスト
	"""
	# 必要列のみ抽出
	df = pd.read_excel(excel_file, sheet_name=sheet_name, header=0)
	df = df[['用語', '該当する品質特性 - 副特性', '用語の説明', '補足']]
	glossary = []
	for _, row in df.iterrows():
		entry = {
			'用語': clean_text(row['用語']),
			'該当する品質特性-副特性': clean_text(row['該当する品質特性 - 副特性']),
			'用語の説明': clean_text(row['用語の説明']),
			'補足': clean_text(row['補足'])
		}
		glossary.append(entry)
	with open(output_file, 'w', encoding='utf-8') as f:
		yaml.dump({'用語集': glossary}, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
	print(f"YAMLファイル '{output_file}' が作成されました。")
	return glossary

def convert_yaml_to_csv(yaml_file, csv_file):
	"""品質モデルのYAMLをCSVへ変換する。

	- 階層(品質特性/副特性/測定項目)をフラット化
	- 元Excelの「セル結合」表現を再現するため、
	  同一グループの2行目以降のラベル列(品質特性/説明、副特性/説明)を空欄にする

	引数:
		yaml_file: 入力YAML
		csv_file: 出力CSV
	戻り値:
		pd.DataFrame: 書き出し済みのデータフレーム
	"""
	# YAMLを読み込み
	with open(yaml_file, 'r', encoding='utf-8') as f:
		yaml_data = yaml.safe_load(f)
	# フラット化のための一時リスト
	csv_data = []
	# 品質特性配列を走査
	for qc in yaml_data.get('品質特性', []):
		qc_name = qc.get('名称', '')
		qc_desc = qc.get('説明', '')
		# 副特性がある場合
		if qc.get('副特性'):
			for sub_qc in qc.get('副特性', []):
				sub_qc_name = sub_qc.get('名称', '')
				sub_qc_desc = sub_qc.get('説明', '')
				# 測定項目を列挙
				for measurement in sub_qc.get('測定項目', []):
					measurement_item = measurement.get('項目', '')
					examples = '\n'.join(measurement.get('例', []))
					violation_examples = '\n'.join(measurement.get('違反例', []))
					csv_data.append({
						'品質特性': qc_name,
						'品質特性の説明': qc_desc,
						'品質副特性': sub_qc_name,
						'品質副特性の説明': sub_qc_desc,
						'測定項目': measurement_item,
						'例': examples,
						'違反例': violation_examples
					})
		else:
			# 副特性が無い場合は品質特性直下の測定項目
			for measurement in qc.get('測定項目', []):
				measurement_item = measurement.get('項目', '')
				examples = '\n'.join(measurement.get('例', []))
				violation_examples = '\n'.join(measurement.get('違反例', []))
				csv_data.append({
					'品質特性': qc_name,
					'品質特性の説明': qc_desc,
					'品質副特性': '',
					'品質副特性の説明': '',
					'測定項目': measurement_item,
					'例': examples,
					'違反例': violation_examples
				})
	# DataFrameへ
	df = pd.DataFrame(csv_data)
	# ===== マージ再現: グループ内2行目以降のラベル列を空欄化 =====
	df_cleaned = df.copy()
	# グループキーを保持(後で破棄)
	df_cleaned['_qc'] = df_cleaned['品質特性']
	df_cleaned['_sub'] = df_cleaned['品質副特性']
	# 品質特性単位: 2行目以降は「品質特性」「品質特性の説明」を空欄
	for qc, idx in df_cleaned.groupby('_qc', sort=False).groups.items():
		if pd.isna(qc) or qc == '':
			continue
		idx_list = list(idx)
		if len(idx_list) > 1:
			df_cleaned.loc[idx_list[1:], ['品質特性', '品質特性の説明']] = ''
	# (品質特性, 品質副特性)単位: 2行目以降は「品質副特性」「品質副特性の説明」を空欄
	for (qc, sub), idx in df_cleaned.groupby(['_qc', '_sub'], sort=False).groups.items():
		if pd.isna(sub) or sub == '':
			continue
		idx_list = list(idx)
		if len(idx_list) > 1:
			df_cleaned.loc[idx_list[1:], ['品質副特性', '品質副特性の説明']] = ''
	# 補助列の削除
	df_cleaned = df_cleaned.drop(columns=['_qc', '_sub'])
	# CSVへ保存
	df_cleaned.to_csv(csv_file, index=False, encoding='utf-8-sig')
	print(f"CSVファイル '{csv_file}' が作成されました。")
	print(f"行数: {len(df_cleaned)}")
	return df_cleaned

if __name__ == "__main__":
	# 変換元Excel/出力ファイル名の設定
	excel_file = "ASDoQ_SystemDocumentationQualityModel_v2.0a-3.xlsx"
	sheet_name_model = "品質特性・副特性・測定項目（例・違反例を含む）"
	output_file_model = "QualityModel_v2.0a-3.YAML"
	# Excel -> YAML
	result = convert_excel_to_yaml(excel_file, sheet_name_model, output_file_model)
	# YAML -> CSV (セル結合再現)
	csv_file = "QualityModel_v2.0a-3.csv"
	csv_result = convert_yaml_to_csv(output_file_model, csv_file)
	# 用語集 -> YAML
	sheet_name_glossary = "用語集"
	output_file_glossary = "Glossary_QualityModel_v2.0a-2.YAML"
	glossary = convert_glossary_to_yaml(excel_file, sheet_name_glossary, output_file_glossary)
	# 変換概要の標準出力
	print(f"\n変換結果概要:")
	print(f"品質特性数: {len(result['品質特性'])}")
	for i, qc in enumerate(result['品質特性'], 1):
		sub_count = len(qc.get('副特性', []))
		direct_measurement_count = len(qc.get('測定項目', []))
		print(f"  {i}. {qc['名称']} - 副特性: {sub_count}個, 直接測定項目: {direct_measurement_count}個")
	print(f"CSV行数: {len(csv_result)}") 