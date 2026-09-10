# PDF編集ツール - 「送る」起動用ランチャー(PowerShell版)
#
# open_in_pdf_editor.vbs(VBScript版)の後継。VBScriptはMicrosoftにより非推奨と
# 発表されており将来的にWindowsから削除される見込みのため、この機能を
# PowerShellへ移植した。ロジックはvbs版と同じ: ローカルの橋渡しサーバー
# (bridge_server.ps1)が起動していなければ起動し、選択されたPDFのパスを
# 短命な一時トークンファイル経由で渡し(日本語パスをURLエンコードせずに
# 済ませるため)、既定のブラウザでツールを開く。
#
# このファイル自身は自己更新ロジックを持たない -- それは別ファイルの
# update_checker.ps1の役割(タスクスケジューラ等から独立して定期実行され、
# このファイルとbridge_server.ps1をまとめて更新する想定)。理由は
# open_in_pdf_editor.vbs時代からの経緯を参照(セッションメモリ)。
#
# Windows 11の「プログラムを選択して開く」ダイアログはコマンド対象が
# powershell.exe/wscript.exe等の汎用スクリプトホストだと一覧から除外して
# しまうため、ダブルクリック/OpenWith経由ではこのファイルを直接登録せず、
# 従来通りネイティブのトランポリンexe(PdfEditorLauncher.exe)経由で起動する。
# 「送る」メニューはこのフィルタの対象外なので、こちらは直接
# powershell.exeをターゲットにしたショートカットで問題ない。

param(
  [Parameter(Position = 0)]
  [string]$PdfPath
)

$ErrorActionPreference = "Stop"

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ServerScript = Join-Path $ToolDir "bridge_server.ps1"
$Port = 8743

function Show-Message([string]$Message) {
  $shell = New-Object -ComObject WScript.Shell
  # 第2引数0 = タイムアウトなし(ユーザーが閉じるまで待つ)、48 = vbExclamation
  $shell.Popup($Message, 0, "PDF Editor Tool", 48) | Out-Null
}

if (-not $PdfPath) {
  Show-Message "ファイル(PDF)を右クリックし、送るから実行してください."
  exit
}

if ([System.IO.Path]::GetExtension($PdfPath).ToLower() -ne ".pdf") {
  Show-Message "PDFファイルを選択してから実行してください."
  exit
}

# 常にサーバーの起動を試みる。既に前回の「送る」から起動済みなら、この
# 試みはポートのバインドに失敗してすぐ終了するだけ -- 無害であり、起動
# 済みかどうかをHTTPで先に確認するより単純。
try {
  # 注意: -ArgumentListに配列を渡すと、Windows PowerShell 5.1は各要素を
  # 自動ではクォートしない -- このツールフォルダの実際のパス(例:
  # "...OneDrive - 朝日工業株式会社\...")のようにスペースを含む場合、
  # 配列渡しだと$ServerScriptがスペースの位置で複数の引数に分割されて
  # 起動に失敗する(実機のパスで確認済み)。手動でクォートした単一の
  # 文字列として渡すことでこれを回避する。
  Start-Process -FilePath "powershell.exe" `
    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ServerScript`" -Port $Port" `
    -WindowStyle Hidden
} catch {
  # 起動コマンド自体が失敗した場合も、後続のWaitForServerReadyが
  # タイムアウトして適切なエラーメッセージを出すので、ここでは何もしない。
}

# サーバーがhttp://127.0.0.1:<port>/index.htmlに実際に応答するまで短い
# HEADリクエストでポーリングする(応答があればどんなHTTPステータスでも
# 良い -- 「何か動いているか」だけを見る単純な確認)。初回起動時、
# ダウンロード直後で未署名のスクリプト実行がWindows Defender等の
# スキャンで遅くなる場合があるため、固定の待機時間ではなくポーリングで
# 確認する(vbs版から引き継いだ設計)。
function Wait-ForServerReady([int]$Port, [int]$MaxWaitMs) {
  $deadline = [DateTime]::UtcNow.AddMilliseconds($MaxWaitMs)
  while ([DateTime]::UtcNow -lt $deadline) {
    try {
      $req = [System.Net.HttpWebRequest]::Create("http://127.0.0.1:$Port/index.html")
      $req.Method = "HEAD"
      $req.Timeout = 500
      $resp = $req.GetResponse()
      $resp.Close()
      return $true
    } catch [System.Net.WebException] {
      if ($_.Exception.Response) {
        # 実際にHTTP応答が返ってきた(エラーステータスでも)ならサーバーは
        # 生きている。
        $_.Exception.Response.Close()
        return $true
      }
      # それ以外(接続不可等)はまだ起動していない -- ポーリングを続ける。
    } catch {
      # 想定外のエラーも「まだ起動していない」として扱い、ポーリングを続ける。
    }
    Start-Sleep -Milliseconds 300
  }
  return $false
}

if (-not (Wait-ForServerReady $Port 20000)) {
  Show-Message "サーバーが応答しません。少し時間をおいて、もう一度「送る」から開き直してください。"
  exit
}

$token = Get-Random -Minimum 100000 -Maximum 999999
$tokenFile = Join-Path $env:TEMP "pdf_editor_$token.path"
[System.IO.File]::WriteAllText($tokenFile, $PdfPath, (New-Object System.Text.UTF8Encoding($false)))

Start-Process "http://127.0.0.1:$Port/index.html?open=$token"
