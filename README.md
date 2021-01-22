# docx-researchmap
グループメンバーの業績(論文、招待講演、書籍、その他)をresearchmapv2のAPIを通じて集め、docxファイルにしてダウンロードするコードです(2021-01-22更新)。

[新学術領域](https://infophys-bio.jp/)の業績報告用に作りましたが、メンバー情報のスプレッドシートを変更すれば他領域や個人でも使えるはずです。

グループの場合：[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/dbkk/docx-researchmap/blob/testing/researchmapv2_to_docx.ipynb)

個人の場合：[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/dbkk/docx-researchmap/blob/testing/researchmapv2_to_docx_single.ipynb)

[メンバー情報例](https://docs.google.com/spreadsheets/d/1wce1XHSFGSBttupnSIqe_5abtijBb_hBYM2bfaV9Jn4/edit)


## 手順

1. 領域メンバーにresearchmapを更新してもらう(業績を"公開"でお願い)
2. "Open with Colab"をクリック
3. 最初のセルでパラメータをいじる
4. ランタイム/すべてのセルを実行
5. docxがダウンロードされるのを待つ
6. ダウンロードされず'''files.download(file_name)'''でエラーが出た場合は、最後のセルの左側の▶を押す(セルを実行)

(大人数だとresearchmapのデータダウンロードに時間がかかるので、ローカルで実行するのがおすすめ)

## 修正すべき事項

1. 共同研究の数等を別紙の書式通りに出力
2. Arxiv対応 (MISCの中の著者ルールなど)
3. 論文表記指定対応、名前表記指定対応、日付表記指定対応
4. 公募班メンバーの業績を加入後に限定 (spreadsheetから日付を読み込んで個別対応?)

## 中間報告書式指定(2020参考)

"研究項目ごとに計画研究・公募研究の順で、本研究領域により得られた研究成果の発表の状況（主な雑誌論文、 学会発表、書籍、産業財産権、ホームページ、主催シンポジウム、一般向けのアウトリーチ活動等の状況。令和２年６月末までに掲載等が確定しているものに限る。）について、具体的かつ簡潔に５頁以内で記述すること。なお、 雑誌論文の記述に当たっては、新しいものから順に発表年次をさかのぼり、研究代表者（発表当時、以下同様。） には二重下線、研究分担者には一重下線、corresponding author には左に＊印を付すこと。"


## 科研費実施状況報告書用(個人)

科研費実施状況報告書のためにアップロードできる業績リストcsvファイルを作るコードです。

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/dbkk/docx-researchmap/blob/testing/researchmapv2_to_csv.ipynb)

* c.f. https://www-shinsei.jsps.go.jp/kaken/docs/2_csv_torikomi.pdf
* 国際共著, オープンアクセス, 国際学会かどうか, 総ページ数は拾えず