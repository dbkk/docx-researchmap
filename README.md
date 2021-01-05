# docx-researchmap
グループメンバーの業績(論文、招待講演、書籍、その他)をresearchmapv2のAPIを通じて集め、docxファイルにしてダウンロードするコードです。

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

1. 班をまたいで著者名がいる場合のマークづけ
2. 共同研究の数等を別紙の書式通りに出力
3. Arxiv対応 (MISCの中の著者ルールなど)
4. 論文表記指定対応（略称リスト作る？）、名前表記指定対応、日付表記指定対応

## 中間報告書式指定(2019参考)

* 本研究課題（公募研究を含む）により得られた研究成果の公表の状況（主な論文、書籍、ホームページ、主催シンポジウム等の状況）について具体的に記述してください。記述に当たっては、本研究課題により得られたものに厳に限ることとします。
* 論文の場合、新しいものから順に発表年次をさかのぼり、研究項目ごとに計画研究・公募研究の順に記載し、研究代表者には二重下線、研究分担者には一重下線、連携研究者には点線の下線を付し、corresponding author には左に＊印を付してください。
* 別添の「（２）発表論文」の融合研究論文として整理した論文については、冒頭に◎を付してください。
* 補助条件に定められたとおり、本研究課題に係り交付を受けて行った研究の成果であることを表示したもの（論文等の場合は謝辞に課題番号を含め記載したもの）について記載したものについては、冒頭に▲を付してください（前項と重複する場合は、「◎▲・・・」と記載してください。）。
* 一般向けのアウトリーチ活動を行った場合はその内容についても記述してください。


## 科研費実施状況報告書用(個人)

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/dbkk/docx-researchmap/blob/testing/researchmapv2_to_csv.ipynb)

* c.f. https://www-shinsei.jsps.go.jp/kaken/docs/2_csv_torikomi.pdf
* 国際共著, オープンアクセス, 国際学会かどうか, 総ページ数は拾えず