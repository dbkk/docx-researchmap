# docx-researchmap
グループメンバーの業績をresearchmapv2のAPIを通じて集め、docxファイルにしてダウンロードするコードです。

- [Open with Colab](https://colab.research.google.com/github/dbkk/docx-researchmap/blob/master/researchmapv2_to_docx.ipynb)

- [メンバー情報](https://docs.google.com/spreadsheets/d/1wce1XHSFGSBttupnSIqe_5abtijBb_hBYM2bfaV9Jn4/edit)

手順:
0. 領域メンバー各自にresearchmapを更新してもらう(業績を"公開"でお願い)
1. "Open with Colab"をクリックし、"ランタイム"/"すべてのセルを実行"(コードはいじれるが保存はされない)
2. 3番目くらいのセルでgoogle spreadsheetへのアクセス認証を求められるので実行("メンバー情報"を参照するために必要)
3. docxがダウンロードされるのを待つ
4. 最後のセル(files.download(file_name)とある)でエラーが出た場合は、最後のセルの左側の▶を押す(セルを実行)

修正すべき事項:
1. 著者名の例外処理... 登録されているメンバーの名前を検出してunderlineしたりSurname, Firstnameの順を決めてたりしているが、登録名と少しでも違うと見つけられない。a別表記もspreadsheetに登録しておくしかないか。
2. 班をまたいで著者名がいる場合のマークづけ...
3. 共同研究の数等の数字を別紙の書式通りに出力...

