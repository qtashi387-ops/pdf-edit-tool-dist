# PDF編集ツール - 「送る」起動用ランチャー(PowerShell版)
#
# open_in_pdf_editor.vbs(VBScript版)の後継。VBScriptはMicrosoftにより非推奨と
# 発表されており将来的にWindowsから削除される見込みのため、この機能を
# PowerShellへ移植した。ロジックはvbs版と同じ: ローカルの橋渡しサーバー
# (bridge_server.ps1)が起動していなければ起動し、選択されたPDFのパスを
# 短命な一時トークンファイル経由で渡し(日本語パスをURLエンコードせずに
# 済ませるため)、既定のブラウザでツールを開く。
#
# このファイル自身が自分自身(open_in_pdf_editor.ps1)とbridge_server.ps1を
# 更新するロジックは持たない -- それは別ファイルのupdate_checker.ps1の役割
# (スタートアップフォルダから独立して定期実行され、この2ファイルをまとめて
# 更新する想定)。理由はopen_in_pdf_editor.vbs時代からの経緯を参照
# (セッションメモリ)。
#
# 一方、index.html(アプリ本体)と利用マニュアルの自動更新チェックは、
# open_in_pdf_editor.vbs時代からずっとこの起動スクリプト自身が「送る」の
# たびに行ってきた実績のある仕組みで、この2ファイルは(update_checker.ps1と
# 違って)このファイル自身を書き換えるわけではないため自己更新ブロックの
# 対象外 -- そのままこのファイルに移植し、以前と同じタイミング(サーバー
# 起動より前、PDFが選択されているかのチェックより前)で毎回実行する。
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
$DistHost = "https://pdf-edit-tool-dist.haruno.workers.dev"

function Show-Message([string]$Message) {
  $shell = New-Object -ComObject WScript.Shell
  # 第2引数0 = タイムアウトなし(ユーザーが閉じるまで待つ)、48 = vbExclamation
  $shell.Popup($Message, 0, "PDF Editor Tool", 48) | Out-Null
}

function Show-InfoToast([string]$Message) {
  $shell = New-Object -ComObject WScript.Shell
  # 第2引数4 = 4秒で自動的に閉じる(無人/バックグラウンド実行でも
  # クリック待ちで固まらないように)、64 = vbInformation
  $shell.Popup($Message, 4, "PDF Editor Tool", 64) | Out-Null
}

function Test-VersionString([string]$Value) {
  if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
  $parts = $Value -split '\.'
  if ($parts.Count -ne 3) { return $false }
  foreach ($p in $parts) {
    if ($p -notmatch '^\d+$' -or $p.Length -gt 9) { return $false }
  }
  return $true
}

# 1ならa>b、-1ならa<b、0なら等しい。両方Test-VersionStringを通過済みの前提。
function Compare-Versions([string]$A, [string]$B) {
  $pa = $A -split '\.'
  $pb = $B -split '\.'
  for ($i = 0; $i -lt 3; $i++) {
    $na = [int64]$pa[$i]
    $nb = [int64]$pb[$i]
    if ($na -gt $nb) { return 1 }
    if ($na -lt $nb) { return -1 }
  }
  return 0
}

# index.html/利用マニュアルの自動更新チェック本体(vbs版のTryAutoUpdate/
# TryUpdateManualに相当)。何が失敗しても(オフライン、配信元が落ちている、
# 壊れたデータ等)静かに諦めてローカルの既存コピーで起動を続行する --
# 「送る」の実行がこのチェックのせいで止まったり見えるエラーを出したり
# することは絶対にない。
function Update-DistFile {
  param(
    [string]$VersionUrl,
    [string]$FileUrl,
    [string]$VersionFileName,
    [string]$TargetFileName,
    [int]$MinBytes,
    [string]$AppliedMessagePrefix
  )
  try {
    $versionFile = Join-Path $ToolDir $VersionFileName
    $targetPath  = Join-Path $ToolDir $TargetFileName
    $newPath     = "$targetPath.new"
    $bakPath     = "$targetPath.bak"

    $localVersion = "0.0.0"
    if ([System.IO.File]::Exists($versionFile)) {
      $v = [System.IO.File]::ReadAllText($versionFile, [System.Text.Encoding]::UTF8).Trim()
      if (Test-VersionString $v) { $localVersion = $v }
    }

    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    $wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")

    $remoteVersion = $wc.DownloadString($VersionUrl).Trim()
    if (-not (Test-VersionString $remoteVersion)) { return }
    if ((Compare-Versions $remoteVersion $localVersion) -le 0) { return }

    $bytes = $wc.DownloadData($FileUrl)
    if ($bytes.Length -lt $MinBytes) { return }

    [System.IO.File]::WriteAllBytes($newPath, $bytes)

    if ([System.IO.File]::Exists($bakPath)) { [System.IO.File]::Delete($bakPath) }
    if ([System.IO.File]::Exists($targetPath)) { [System.IO.File]::Move($targetPath, $bakPath) }
    try {
      [System.IO.File]::Move($newPath, $targetPath)
    } catch {
      if ([System.IO.File]::Exists($bakPath) -and -not [System.IO.File]::Exists($targetPath)) {
        [System.IO.File]::Move($bakPath, $targetPath)
      }
    }

    if (-not [System.IO.File]::Exists($targetPath)) { return }

    [System.IO.File]::WriteAllText($versionFile, $remoteVersion, (New-Object System.Text.UTF8Encoding($false)))

    Show-InfoToast "$AppliedMessagePrefix(v$remoteVersion)"
  } catch {
    # オフライン/配信元不達/破損データ等、何が原因でも静かに諦める。
  }
}

Update-DistFile `
  -VersionUrl "$DistHost/version.txt" `
  -FileUrl "$DistHost/index.html" `
  -VersionFileName "version.txt" `
  -TargetFileName "index.html" `
  -MinBytes 1000000 `
  -AppliedMessagePrefix "PDF編集ツールを更新しました"

Update-DistFile `
  -VersionUrl "$DistHost/manual_version.txt" `
  -FileUrl "$DistHost/PDF%E7%B7%A8%E9%9B%86%E3%83%84%E3%83%BC%E3%83%AB_%E5%88%A9%E7%94%A8%E3%83%9E%E3%83%8B%E3%83%A5%E3%82%A2%E3%83%AB.docx" `
  -VersionFileName "manual_version.txt" `
  -TargetFileName "PDF編集ツール_利用マニュアル.docx" `
  -MinBytes 10000 `
  -AppliedMessagePrefix "利用マニュアルを更新しました"

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
