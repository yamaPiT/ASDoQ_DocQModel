"""
Markdown化した文書品質モデルをExcelへ変換スクリプトです。
Excelファイルの構造は、次のようにします。
1行目について
- 1列目 「品質副特性」と決め打ち
- 2列目「説明」と決め打ち
- 3列目「測定項目」と決め打ち
- 4列目「例」と決め打ち
- 5列目「違反例」と決め打ち
2行目以降について
- 1列目 見出し1の本文
- 2列目 見出し2（説明）の本文
- 3列目 見出し3（測定項目）のタイトルと本文（タイトルと本文の間は改行）
- 4列目 見出し4（例）の本文
- 5列目 見出し4（違反例）の本文
- 見出し3は、複数個あれば行を改める。行を改めた場合、1列目と2列目は上の内容と同じものを入力する。
- 4列目及び5列目に、2層の箇条書きがあれば2層目は"--"で表す。3層目以下も同様に"---"などと増やしていく。
- 4列目及び5列目に、見出し5の内容を含む場合、見出し5のタイトル及び本文を追加する。
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment

def markdown_to_quality_excel(md_path, xlsx_path):
    # Markdown全文を読み込む
    with open(md_path, encoding='utf-8') as f:
        lines = f.readlines()

    # frontmatterスキップ用フラグ
    in_yaml = False

    # 「レベル1」「レベル2」の本文を貯めておくリスト
    l1_body_lines = []
    l2_body_lines = []

    # 最新のレベル1/2見出しテキスト（見出しタイトル自体はIDとかなので使わない）
    # もしタイトルもセルに出したいなら別途変数を用意してください
    current_state = None  # 'l1', 'l2', 'body', 'example', 'violation'

    # レベル3アイテムをまとめる変数
    current_item = None
    items = []

    INDENT_WIDTH = 4

    # 正規表現
    header_re = re.compile(r'^(#{1,6})\s*(.*)$')
    list_re   = re.compile(r'^(\s*)([-*])\s+(.*)$')

    for raw in lines:
        line = raw.rstrip('\n')

        # --- で囲まれたYAML frontmatterはスキップ
        if line.strip() == '---':
            in_yaml = not in_yaml
            continue
        if in_yaml:
            continue

        # 空行はスキップ
        if not line.strip():
            continue

        # 見出し行か判定
        m = header_re.match(line)
        if m:
            level = len(m.group(1))
            title = m.group(2).strip()

            # レベル1見出し => 本文収集中の前アイテムを確定、stateを'l1'に
            if level == 1:
                if current_item:
                    items.append(current_item)
                    current_item = None

                l1_body_lines = []
                current_state = 'l1'
                continue

            # レベル2見出し => 本文収集中の前アイテムを確定、stateを'l2'に
            if level == 2:
                # current_item の確定はレベル3のタイミングのみに任せる
                # （もしレベル2で current_item をappendしたいなら別途ロジックを追加）

                if title == '説明':
                    # 「説明」セクションの本文をこれから貯め始める
                    l2_body_lines = []
                    current_state = 'l2'
                else:
                    # 「測定項目」などは本文を集めない
                    current_state = None
                continue


            # レベル3見出し => 新しい測定項目スタート
            if level == 3:
                # 前に作っていたcurrent_itemがあれば確定
                if current_item:
                    items.append(current_item)

                # 今までの L1/L2 本文を確定して、各行にセットする
                c1 = "\n".join(l1_body_lines).strip()
                c2 = "\n".join(l2_body_lines).strip()

                # 新しいアイテムを生成
                current_item = {
                    'col1': c1,              # レベル1 の本文
                    'col2': c2,              # レベル2 の本文
                    'title3': title,         # レベル3 のタイトル
                    'body3': [],             # レベル3 の本文
                    'examples': [],          # レベル4 (例)
                    'violations': []         # レベル4 (違反例)
                }
                current_state = 'body'
                continue

            # レベル4見出し => "例" or "違反例" の切り替え
            if level == 4 and current_item:
                if title.startswith('例'):
                    current_state = 'example'
                elif title.startswith('違反例'):
                    current_state = 'violation'
                else:
                    current_state = None
                continue

            # レベル5以降は例／違反例セルにプレフィックス付きで入れる
            if level >= 5 and current_item and current_state in ('example','violation'):
                prefix = ''
                if current_state == 'example':
                    current_item['examples'].append(prefix + title)
                else:
                    current_item['violations'].append(prefix + title)
                continue

            # それ以外の見出しは無視
            continue

        # 箇条書き行か判定
        lm = list_re.match(line)
        if lm and current_item:
            indent = len(lm.group(1)) // INDENT_WIDTH
            text = lm.group(3).strip()
            prefix = '-' * (indent + 1)+ ' '
            if current_state == 'body':
                current_item['body3'].append(prefix + text)
            elif current_state == 'example':
                current_item['examples'].append(prefix + text)
            elif current_state == 'violation':
                current_item['violations'].append(prefix + text)
            continue

        # 通常テキスト行
        if current_state == 'l1':
            l1_body_lines.append(line.strip())
        elif current_state == 'l2':
            l2_body_lines.append(line.strip())
        elif current_item and current_state == 'body':
            current_item['body3'].append(line.strip())
        elif current_item and current_state == 'example':
            current_item['examples'].append(line.strip())
        elif current_item and current_state == 'violation':
            text = line.strip()
            if text == "***":
                continue
            current_item['violations'].append(text)
        # else: それ以外は無視

    # ループ終わりで最後のcurrent_itemを確定
    if current_item:
        items.append(current_item)

    # Excelファイル書き出し
    wb = Workbook()
    ws = wb.active

    # ヘッダー行（1行目）
    ws.append(['品質副特性', '説明', '測定項目', '例', '違反例'])

    # データ行
    for it in items:
        # Col1: レベル1本文
        c1 = it['col1']
        # Col2: レベル2本文
        c2 = it['col2']
        # Col3: レベル3「タイトル\n本文」
        c3 = "\n".join([it['title3']] + it['body3'])
        # Col4: 「例」
        c4 = "\n".join(it['examples'])
        # Col5: 「違反例」
        c5 = "\n".join(it['violations'])

        ws.append([c1, c2, c3, c4, c5])

    # セル内折り返しを有効化
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)

    wb.save(xlsx_path)
    print(f"保存完了: {xlsx_path}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(f"Usage: python {sys.argv[0]} input.md [output.xlsx]")
        sys.exit(1)
    md_file   = sys.argv[1]
    xlsx_file = sys.argv[2] if len(sys.argv) > 2 else 'output.xlsx'
    markdown_to_quality_excel(md_file, xlsx_file)