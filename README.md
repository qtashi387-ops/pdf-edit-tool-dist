# PDF編集ツール 配布用

社内向けPDF編集ツールの自動更新チェック用に、以下のファイルを公開しています。

- `version.txt` / `index.html` — アプリ本体の最新版
- `manual_version.txt` / `PDF編集ツール_利用マニュアル.docx` — 利用マニュアルの最新版
- `launcher_version.txt` / `open_in_pdf_editor.ps1` — 送る起動用ランチャーの最新版
- `bridge_server_version.txt` / `bridge_server.ps1` — 橋渡しサーバーの最新版

`open_in_pdf_editor.ps1`(「送る」起動用スクリプト、2026-09-10に旧
`open_in_pdf_editor.vbs`から移行)が起動のたびに上2つをチェックし、新しい版が
あれば自動でダウンロード・差し替えます。下2つ(`open_in_pdf_editor.ps1`
自身と`bridge_server.ps1`)は、起動スクリプト自身が自分を更新するのではなく、
別途スタートアップに登録された`update_checker.ps1`が独立してログオン時に
チェック・差し替えを行います(旧`open_in_pdf_editor.vbs`時代は自己更新
ロジックの追加が一貫してブロックされ、この2ファイルだけ自動更新できな
かった -- 詳細は開発リポジトリ側のセッションメモリ参照)。

このリポジトリはツール本体の開発用ソースではなく、配布専用のミラーです。
