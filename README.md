# docx-researchmap
グループメンバーの業績(論文、招待講演、書籍、その他)をresearchmapv2のAPIを通じて集め、docxファイルにしてダウンロードするコードです。

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/dbkk/docx-researchmap/blob/rev2026/researchmapv2_to_docx.ipynb)

[メンバー情報例](https://docs.google.com/spreadsheets/d/1T5QtMv4M_peHHM-Zj4oFmS1jHG4voDBJipbEEY0xFQs/edit?usp=sharing)

## 手順

1. 領域メンバーにresearchmapを更新してもらう(業績を"公開"でお願い) ([入力方法参考](https://sites.google.com/view/researchmap3))
2. 上の"Open with Colab"をクリック
3. 最初のセル(Parameters)のフォームでパラメータを入力する
4. ランタイム → すべてのセルを実行
5. docxがダウンロードされるのを待つ
6. ダウンロードされず`files.download(file_name)`でエラーが出た場合は、最後のセルの左側の▶を押す(セルを実行)

## rev2026での変更点

- Colabフォーム対応: パラメータセルがフォームUIで表示され、他のセルは折りたたまれる
- 雑誌名をISO 4略称で表示 (例: Science Advances → Sci. Adv.)、[abbreviso](https://abbreviso.toolforge.org/)経由
- bioRxiv自動検出: DOIプレフィックス(`10.1101/`, `10.64898/`)から判定
- PermissionError対策(ローカル用): 出力ファイルが開かれている場合、タイムスタンプ付きファイル名で保存