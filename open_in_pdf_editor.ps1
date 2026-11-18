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
# 一方、index.html(アプリ本体)・利用マニュアル・ocr-data.bin(2026-09-11
# 追加)の自動更新チェックは、open_in_pdf_editor.vbs時代からずっとこの
# 起動スクリプト自身が「送る」のたびに行ってきた実績のある仕組みで、この
# 3ファイルは(update_checker.ps1が対象とするファイル群と違って)このファイル
# 自身を書き換えるわけではないため自己更新ブロックの対象外 -- そのまま
# このファイルに移植し、以前と同じタイミング(サーバー起動より前、PDFが
# 選択されているかのチェックより前)で実行する。
#
# launcher v1.0.4(2026-11-18、起動を速くした):
#  (1) 更新確認は、直近10分以内に成功していれば3つとも省略する。行うときは、3つの
#      版ファイルを同時に取りに行き(以前は1つずつ順番)、最長8秒で打ち切る。
#  (2) ブラウザで開くURLを「index.html#open=番号」にした(画面側がv1.43.2から対応)。
#      「?open=番号」は番号が毎回変わり、ブラウザのキャッシュが効かなかった。
#      ブリッジ(v1.0.9)が目印(ETag)を付けるので、2回目以降は約15MBの画面本体を
#      取り直さずに済む。
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
# 変更する場合はbridge_server.ps1のparam既定値も必ず同じ値に変更すること
# (このスクリプトが-Portで明示的に渡すので通常は既定値は使われないが、
# 2箇所の値がずれると起動待ちが必ずタイムアウトする)。
$Port = 8743
$DistHost = "https://pdf-edit-tool-dist.haruno.workers.dev"

# --- 起動を速くするための設定(launcher v1.0.4) ---
# 更新確認(index.html・利用マニュアル・OCR用データの3つ)の頻度制限: 直近にこの分数以内に
# 確認に成功していれば、3つともまるごと省略する(確認は毎回約0.3〜0.8秒かかる)。
# 代償: 新しい版が出てから手元に届くまで、最長でこの分数だけ遅れうる。
$UpdateCheckIntervalMinutes = 10
$UpdateCheckStampFile = Join-Path $ToolDir "update_check_stamp.txt"
# 3つの版ファイル(数十バイト)の確認を、同時に待つ最大ミリ秒。これを超えたら、今回は更新せず
# 手元の版のまま起動する(次回また確認する)。以前は時間の制限が無く、通信が詰まると、
# WebClientの既定(100秒)まで1つずつ待ち続けた。
$VersionCheckTimeoutMs = 8000
# 「index.html#open=番号」でブラウザを開いてよい、アプリ(index.html)の最低バージョン。
# #open=を読めるのは画面側がv1.43.2から。手元のindex.htmlがこれより古い(更新できず
# 古いまま起動する)ときは、従来の「?open=番号」で開く。
$HashOpenMinAppVersion = "1.43.2"

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

# --- 更新確認の頻度制限(launcher v1.0.4) ---
# 直近に更新確認に成功した時刻を、ツールフォルダの小さなファイルに残す。
# 返り値: 今回、更新確認を行うべきならtrue。ファイルが無い・読めない・壊れている・
# 間隔(分)を過ぎている・時計が戻って未来の時刻になっている、のどれでもtrue
# (=疑わしいときは、確認する側に倒す)。
function Test-UpdateCheckDue([string]$StampFile, [int]$IntervalMinutes) {
  try {
    if (-not [System.IO.File]::Exists($StampFile)) { return $true }
    $text = [System.IO.File]::ReadAllText($StampFile).Trim()
    $stamp = [DateTime]::ParseExact($text, "o", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
    $ageMinutes = ([DateTime]::UtcNow - $stamp.ToUniversalTime()).TotalMinutes
    if ($ageMinutes -lt 0) { return $true }
    return ($ageMinutes -ge $IntervalMinutes)
  } catch {
    return $true
  }
}

function Write-UpdateCheckStamp([string]$StampFile) {
  try {
    [System.IO.File]::WriteAllText($StampFile, [DateTime]::UtcNow.ToString("o"), (New-Object System.Text.UTF8Encoding($false)))
  } catch {
    # 書けなくても、次回また確認するだけ(起動は止めない)。
  }
}

# --- 版ファイルの確認を同時に行う(launcher v1.0.4) ---
# 以前は、3つの版ファイル(version.txt・manual_version.txt・ocr_data_version.txt)を
# 1つずつ順番に取りに行っていた。同時に取りに行き、3回分の往復を1回分にまとめる。
# WebClientは1つで同時に1通信しかできないので、1通信につき1つ作る。
function Start-VersionFetch([string]$Url) {
  try {
    # 同じ相手への同時接続数の既定は2本まで(3つ目が待たされ、同時にならない)。引き上げる
    # (このプロセスだけの設定。接続先ごとの上限は最初の通信のときに決まるので、
    # 最初の通信より前に行う)。
    if ([System.Net.ServicePointManager]::DefaultConnectionLimit -lt 8) { [System.Net.ServicePointManager]::DefaultConnectionLimit = 8 }
    $wc = New-Object System.Net.WebClient
    # UTF-8の指定とUser-Agentの理由は、Update-DistFileの同じ箇所のコメントを参照。
    $wc.Encoding = [System.Text.Encoding]::UTF8
    $wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
    return @{ Client = $wc; Task = $wc.DownloadStringTaskAsync($Url) }
  } catch {
    return $null
  }
}

# 全部の取得が終わるまで(または、全体で$TimeoutMsを過ぎるまで)待つ。
# 失敗した取得があっても例外にはしない(結果はGet-FetchedVersionで見る)。
function Wait-VersionFetches([object[]]$Fetches, [int]$TimeoutMs) {
  $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
  foreach ($f in $Fetches) {
    if ($null -eq $f) { continue }
    $remainingMs = [int][Math]::Max(0, ($deadline - [DateTime]::UtcNow).TotalMilliseconds)
    try { $null = $f.Task.Wait($remainingMs) } catch { }
  }
}

# 取得できた版の文字列(前後の空白を除いたもの)。取得が失敗・時間切れなら$null
# (時間切れの通信は取り消す)。
function Get-FetchedVersion($Fetch) {
  if ($null -eq $Fetch) { return $null }
  try {
    if ($Fetch.Task.Status -ne [System.Threading.Tasks.TaskStatus]::RanToCompletion) {
      try { $Fetch.Client.CancelAsync() } catch { }
      return $null
    }
    return ([string]$Fetch.Task.Result).Trim()
  } catch {
    return $null
  }
}

# ブラウザで開くURL。画面側(index.html)が「#open=番号」を読めるバージョン以上なら、
# 「#open=番号」(URLの#以降は、ブラウザのキャッシュの照合に使われない=毎回同じURLとして
# 約15MBの画面本体をキャッシュから使える)。古いままなら従来の「?open=番号」
# (この場合は番号が毎回変わるのでキャッシュは効かないが、古い画面でも確実にPDFが開く)。
function Get-BrowserUrl([int]$Port, $Token, [string]$LocalAppVersion, [string]$HashOpenMinVersion) {
  $separator = "?"
  if ((Test-VersionString $LocalAppVersion) -and (Compare-Versions $LocalAppVersion $HashOpenMinVersion) -ge 0) { $separator = "#" }
  return "http://127.0.0.1:$Port/index.html${separator}open=$Token"
}

# index.html/利用マニュアルの自動更新チェック本体(vbs版のTryAutoUpdate/
# TryUpdateManualに相当)。何が失敗しても(オフライン、配信元が落ちている、
# 壊れたデータ等)静かに諦めてローカルの既存コピーで起動を続行する --
# 「送る」の実行がこのチェックのせいで止まったり見えるエラーを出したり
# することは絶対にない。
#
# 配信元の版($RemoteVersionText)は、呼び出し側が先に同時に取得したものを受け取る
# (Start-VersionFetch/Get-FetchedVersion)。取得できなかった場合は$nullを渡す。
# 返り値(文字列): "applied"=新しい版を適用した / "current"=確認できて、すでに最新 /
# "failed"=確認または適用に失敗した(今回は更新できていない)。
function Update-DistFile {
  param(
    [string]$RemoteVersionText,
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

    $remoteVersion = "$RemoteVersionText".Trim()
    if (-not (Test-VersionString $remoteVersion)) { return "failed" }
    if ((Compare-Versions $remoteVersion $localVersion) -le 0) { return "current" }

    $wc = New-Object System.Net.WebClient
    # WebClient.Encodingの既定はOSのANSIコードページ(この日本語ロケール機では
    # Shift-JIS)であってUTF-8ではない -- 明示しないと、将来配信元の内容が
    # ASCII以外を含んだ場合に静かに文字化けする(このセッションで実際に
    # 踏んだBOM/エンコーディングの罠と同じ種類の問題)。
    $wc.Encoding = [System.Text.Encoding]::UTF8
    $wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")

    $bytes = $wc.DownloadData($FileUrl)
    if ($bytes.Length -lt $MinBytes) { return "failed" }

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

    if (-not [System.IO.File]::Exists($targetPath)) { return "failed" }

    [System.IO.File]::WriteAllText($versionFile, $remoteVersion, (New-Object System.Text.UTF8Encoding($false)))

    Show-InfoToast "$AppliedMessagePrefix(v$remoteVersion)"
    return "applied"
  } catch {
    # オフライン/配信元不達/破損データ等、何が原因でも静かに諦める。
    return "failed"
  }
}

# 今回、更新確認を行うか(直近$UpdateCheckIntervalMinutes分以内に成功していれば省略)。
# 3つの確認(index.html・利用マニュアル・OCR用データ)は、まとめて行う・まとめて省略する。
$updateDue = Test-UpdateCheckDue $UpdateCheckStampFile $UpdateCheckIntervalMinutes
$updateResults = @()
$versionFetches = @{}
if ($updateDue) {
  $versionFetches = @{
    App    = Start-VersionFetch "$DistHost/version.txt"
    Manual = Start-VersionFetch "$DistHost/manual_version.txt"
    Ocr    = Start-VersionFetch "$DistHost/ocr_data_version.txt"
  }
  Wait-VersionFetches @($versionFetches.App, $versionFetches.Manual, $versionFetches.Ocr) $VersionCheckTimeoutMs

  # 戻り値は必ず変数で受ける(受けないと、パイプライン/stdoutに文字列がそのまま出る --
  # update_checker.ps1で実際に踏んだ罠)。
  $updateResults += Update-DistFile `
    -RemoteVersionText (Get-FetchedVersion $versionFetches.App) `
    -FileUrl "$DistHost/index.html" `
    -VersionFileName "version.txt" `
    -TargetFileName "index.html" `
    -MinBytes 1000000 `
    -AppliedMessagePrefix "PDF編集ツールを更新しました"

  $updateResults += Update-DistFile `
    -RemoteVersionText (Get-FetchedVersion $versionFetches.Manual) `
    -FileUrl "$DistHost/PDF%E7%B7%A8%E9%9B%86%E3%83%84%E3%83%BC%E3%83%AB_%E5%88%A9%E7%94%A8%E3%83%9E%E3%83%8B%E3%83%A5%E3%82%A2%E3%83%AB.docx" `
    -VersionFileName "manual_version.txt" `
    -TargetFileName "PDF編集ツール_利用マニュアル.docx" `
    -MinBytes 10000 `
    -AppliedMessagePrefix "利用マニュアルを更新しました"
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

# ocr-data.bin(index.htmlが実行時に読み込む、index.html本体とは別出しの
# ~9MB OCR用データファイル)もindex.htmlと同じ仕組みで自動更新する --
# こちらもこのファイル自身を書き換えるわけではないため自己更新ブロックの
# 対象外。OCRはクリックされたときだけ必要なファイルなので、index.html/
# 利用マニュアルのチェックと違って「送る」の応答性を優先し、あえて
# Start-Process(非同期)の直後・Wait-ForServerReadyのポーリングループの
# 直前に置く -- このチェックのネットワーク待ち時間がサーバー自身の起動
# 待ち時間と重なり、素朴に直列実行するより「送る」全体の体感速度が
# 落ちない。
if ($updateDue) {
  $updateResults += Update-DistFile `
    -RemoteVersionText (Get-FetchedVersion $versionFetches.Ocr) `
    -FileUrl "$DistHost/ocr-data.bin" `
    -VersionFileName "ocr_data_version.txt" `
    -TargetFileName "ocr-data.bin" `
    -MinBytes 5000000 `
    -AppliedMessagePrefix "OCR用データを更新しました"

  # 3つとも、確認できた(最新だった・適用できた)ときだけ、成功した時刻を残す。
  # 1つでも失敗(オフライン・時間切れ・ダウンロード失敗等)していたら残さない
  # =次回の起動でまた確認する。
  if ($updateResults.Count -eq 3 -and -not ($updateResults -contains "failed")) {
    Write-UpdateCheckStamp $UpdateCheckStampFile
  }
}

if (-not (Wait-ForServerReady $Port 20000)) {
  Show-Message "サーバーが応答しません。少し時間をおいて、もう一度「送る」から開き直してください。"
  exit
}

$token = Get-Random -Minimum 100000 -Maximum 999999
$tokenFile = Join-Path $env:TEMP "pdf_editor_$token.path"
[System.IO.File]::WriteAllText($tokenFile, $PdfPath, (New-Object System.Text.UTF8Encoding($false)))

# 手元のindex.htmlの版(上の更新で変わった後の値)を見て、URLの形を決める。
$localAppVersion = ""
try { $localAppVersion = [System.IO.File]::ReadAllText((Join-Path $ToolDir "version.txt"), [System.Text.Encoding]::UTF8).Trim() } catch { }
Start-Process (Get-BrowserUrl $Port $token $localAppVersion $HashOpenMinAppVersion)
