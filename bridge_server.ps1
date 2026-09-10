# PDF編集ツール - ローカル橋渡しサーバー(PowerShell版)
#
# 旧server.py(旧PdfEditorServer.exeのビルド元、2026-09-10に削除)と全く同じ
# 役割を、コンパイル不要なPowerShellスクリプトとして再実装したもの。目的は、
# この橋渡しサーバー自体をindex.html/マニュアルと同じ「ダウンロードして
# ローカルファイルを上書きする」自動更新の仕組みに乗せられるようにすること
# (コンパイル済みEXEはこの方式に乗せられず、修正のたびに手動での再配布が
# 必要だった)。
#
# 2026-09-10、実機での動作確認・server.pyとのバイト単位の突き合わせ検証を
# 終え、open_in_pdf_editor.ps1(旧open_in_pdf_editor.vbsの後継)から起動する
# 本番の橋渡しサーバーとして既に切り替え済み。以後の更新は
# update_checker.ps1が自動配布する。
#
# 127.0.0.1のみでリッスン(ネットワークからは到達不可)。index.htmlを通常の
# 静的ファイルとして配信するほか、3つの追加エンドポイントを持つ:
#
#   GET  /__load  ローカルで選択されたPDFのバイト列をJSONでページへ渡す
#                (file://ページはブラウザの制約で任意のローカルファイルを
#                読めないため)
#   POST /__save  編集後のPDFバイト列を同じパスへ書き戻す。あるパスへの
#                最初の保存時のみ、編集前のオリジナルの一度きりの
#                "<name>.bak"バックアップを作る(既に存在する場合は作らない)
#   GET  /__font  Windowsライセンスのシステムフォント(MSゴシック/明朝、
#                游ゴシック等)を1書体分抽出してJSONで返す。許可された
#                キーのみ受け付ける(下記FontDefs参照)。フォントファイル
#                自体はこのツールに同梱・再配布されず、その場でこのPCから
#                読み取るのみ。
#
# 実際のファイルパスは/__loadのURLには一切含めない。open_in_pdf_editor.vbsが
# 短命な一時ファイルにパスを書き込み、URLにはランダムな数値トークンのみを
# 渡す。/__saveはJSONボディ(URLではなく)でパスを受け取る -- 既にページが
# /__load経由でパスを学習済みのため。このサーバー自身が過去に配信した
# パスにしか書き込めないよう制限している(下記参照)。

param(
  [int]$Port = 8743
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# このサーバーが/__load経由で実際に渡したパス(元のパスと、ダウンロード
# ボタンが同じフォルダへの保存に使う"_編集済み"派生パスの両方)。/__save は
# これに含まれるパスにしか書き込まない -- ループバックポートを直接叩く
# 他のローカルプロセスに対する多層防御。
$Script:ServedPaths = New-Object 'System.Collections.Generic.HashSet[string]'

# フォントキー -> 既に抽出済みのバイト列。同じシステムフォントで再選択/
# 再エクスポートするたびに(数MBの)元ファイルを読み直さないためのキャッシュ。
$Script:FontCache = @{}

# key -> フォントファイル内で探す書体名、候補ファイル名(見つかった順に採用)
$FontDefs = @{
  "msgothic"          = @{ Family = "MS Gothic";  Candidates = @("msgothic.ttc") }
  "mspgothic"         = @{ Family = "MS PGothic"; Candidates = @("msgothic.ttc") }
  "msmincho"          = @{ Family = "MS Mincho";  Candidates = @("msmincho.ttc") }
  "mspmincho"         = @{ Family = "MS PMincho"; Candidates = @("msmincho.ttc") }
  "yugothic-regular"  = @{ Family = "Yu Gothic";  Candidates = @("YuGothR.ttc", "YuGoth.ttc") }
  "yugothic-bold"     = @{ Family = "Yu Gothic";  Candidates = @("YuGothB.ttc") }
}

$WindowsFontsDir = Join-Path $env:WINDIR "Fonts"
$FontsRegKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"

# ----------------------------------------------------------------------
# ビッグエンディアン読み取りヘルパー -- SFNT/TTCフォーマットは全てビッグ
# エンディアンだが、[BitConverter]は既定でリトルエンディアンのため、
# server.pyのstruct.unpack(">...")と同じ結果になるよう手書きする。
# ----------------------------------------------------------------------
function Read-UInt16BE([byte[]]$Buf, [int]$Offset) {
  return ([int]$Buf[$Offset] -shl 8) -bor [int]$Buf[$Offset + 1]
}
function Read-UInt32BE([byte[]]$Buf, [int]$Offset) {
  return ([uint32]$Buf[$Offset] -shl 24) -bor ([uint32]$Buf[$Offset + 1] -shl 16) -bor ([uint32]$Buf[$Offset + 2] -shl 8) -bor [uint32]$Buf[$Offset + 3]
}
function Read-Tag([byte[]]$Buf, [int]$Offset) {
  return [System.Text.Encoding]::ASCII.GetString($Buf, $Offset, 4)
}
function Write-UInt32BE([byte[]]$Buf, [int]$Offset, [uint32]$Value) {
  $Buf[$Offset]     = [byte](($Value -shr 24) -band 0xFF)
  $Buf[$Offset + 1] = [byte](($Value -shr 16) -band 0xFF)
  $Buf[$Offset + 2] = [byte](($Value -shr 8) -band 0xFF)
  $Buf[$Offset + 3] = [byte]($Value -band 0xFF)
}

function Read-TableDirectory([byte[]]$Buf, [int]$Offset) {
  $numTables = Read-UInt16BE $Buf ($Offset + 4)
  $entries = @()
  $dirStart = $Offset + 12
  for ($i = 0; $i -lt $numTables; $i++) {
    $entryOff = $dirStart + $i * 16
    $entries += [PSCustomObject]@{
      Tag      = Read-Tag $Buf $entryOff
      Checksum = Read-UInt32BE $Buf ($entryOff + 4)
      Offset   = Read-UInt32BE $Buf ($entryOff + 8)
      Length   = Read-UInt32BE $Buf ($entryOff + 12)
    }
  }
  return $entries
}

function Decode-NameTable([byte[]]$NameTableBytes) {
  $count = Read-UInt16BE $NameTableBytes 2
  $stringOffset = Read-UInt16BE $NameTableBytes 4
  $best = $null
  $bestPriority = [int]::MaxValue
  for ($i = 0; $i -lt $count; $i++) {
    $recOff = 6 + $i * 12
    $platformId = Read-UInt16BE $NameTableBytes $recOff
    $nameId     = Read-UInt16BE $NameTableBytes ($recOff + 6)
    $length     = Read-UInt16BE $NameTableBytes ($recOff + 8)
    $offset     = Read-UInt16BE $NameTableBytes ($recOff + 10)
    if ($nameId -ne 1 -and $nameId -ne 4 -and $nameId -ne 16) { continue }
    $rawStart = $stringOffset + $offset
    if ($rawStart + $length -gt $NameTableBytes.Length) { continue }
    $text = $null
    if ($platformId -eq 3) {
      try { $text = [System.Text.Encoding]::BigEndianUnicode.GetString($NameTableBytes, $rawStart, $length) } catch { continue }
    } elseif ($platformId -eq 1) {
      try { $text = [System.Text.Encoding]::ASCII.GetString($NameTableBytes, $rawStart, $length) } catch { continue }
    } else {
      continue
    }
    $priority = switch ($nameId) { 1 { 0 }; 16 { 1 }; 4 { 2 } }
    if ($priority -lt $bestPriority) { $bestPriority = $priority; $best = $text }
  }
  return $best
}

function Get-FaceFamilyName([byte[]]$Buf, [int]$FaceSfntOffset) {
  $entries = Read-TableDirectory $Buf $FaceSfntOffset
  $nameEntry = $entries | Where-Object { $_.Tag -eq "name" } | Select-Object -First 1
  if (-not $nameEntry) { return $null }
  $nameBytes = New-Object byte[] $nameEntry.Length
  [Array]::Copy($Buf, $nameEntry.Offset, $nameBytes, 0, $nameEntry.Length)
  return Decode-NameTable $nameBytes
}

function Find-FaceIndex([byte[]]$Buf, [string]$TargetFamily) {
  $magic = Read-Tag $Buf 0
  if ($magic -ne "ttcf") { return 0 } # コレクションでない単体フォントは面が1つだけ
  $numFonts = Read-UInt32BE $Buf 8
  for ($i = 0; $i -lt $numFonts; $i++) {
    $off = Read-UInt32BE $Buf (12 + $i * 4)
    if ((Get-FaceFamilyName $Buf $off) -eq $TargetFamily) { return $i }
  }
  return -1
}

function Get-SfntChecksum([byte[]]$Data) {
  $pad = (4 - ($Data.Length % 4)) % 4
  if ($pad -gt 0) {
    $padded = New-Object byte[] ($Data.Length + $pad)
    [Array]::Copy($Data, $padded, $Data.Length)
    $Data = $padded
  }
  # PowerShellは0xFFFFFFFFのような0x7FFFFFFFを超える16進リテラルを符号付き
  # Int32(つまり-1)として解釈するため、"-band 0xFFFFFFFF"による32bitマスクは
  # 実質no-opになり、合計がuint32の範囲を超えた際にキャスト例外を起こす。
  # [uint64]で加算し、10進の4294967296で剰余を取ることでこの罠を回避する。
  [uint64]$total = 0
  for ($i = 0; $i -lt $Data.Length; $i += 4) {
    $word = Read-UInt32BE $Data $i
    $total = ($total + [uint64]$word) % 4294967296
  }
  return [uint32]$total
}

function Extract-Face([byte[]]$Buf, [int]$FaceIndex) {
  $magic = Read-Tag $Buf 0
  if ($magic -eq "ttcf") {
    $sfntOffset = Read-UInt32BE $Buf (12 + $FaceIndex * 4)
  } else {
    $sfntOffset = 0
  }

  $entries = Read-TableDirectory $Buf $sfntOffset
  $sfntVersionBytes = New-Object byte[] 4
  [Array]::Copy($Buf, $sfntOffset, $sfntVersionBytes, 0, 4)
  $numTables = $entries.Count

  $entrySelector = 0
  if ($numTables -gt 0) { $entrySelector = [int][Math]::Floor([Math]::Log($numTables, 2)) }
  $searchRange = [Math]::Pow(2, $entrySelector) * 16
  $rangeShift = $numTables * 16 - $searchRange

  $headerStream = New-Object System.IO.MemoryStream
  $headerStream.Write($sfntVersionBytes, 0, 4)
  $tmp = New-Object byte[] 2
  $writeU16 = {
    param($val)
    $b = New-Object byte[] 2
    $b[0] = [byte](($val -shr 8) -band 0xFF); $b[1] = [byte]($val -band 0xFF)
    $headerStream.Write($b, 0, 2)
  }
  & $writeU16 $numTables
  & $writeU16 ([int]$searchRange)
  & $writeU16 $entrySelector
  & $writeU16 ([int]$rangeShift)

  $bodyStart = 12 + $numTables * 16
  $newEntries = @()
  $bodyChunks = New-Object System.Collections.Generic.List[byte[]]
  $cursor = $bodyStart
  foreach ($e in $entries) {
    $data = New-Object byte[] $e.Length
    [Array]::Copy($Buf, $e.Offset, $data, 0, $e.Length)
    $pad = (4 - ($data.Length % 4)) % 4
    if ($pad -gt 0) {
      $padded = New-Object byte[] ($data.Length + $pad)
      [Array]::Copy($data, $padded, $data.Length)
      $data = $padded
    }
    $newEntries += [PSCustomObject]@{ Tag = $e.Tag; Checksum = $e.Checksum; Offset = $cursor; Length = $e.Length }
    $bodyChunks.Add($data)
    $cursor += $data.Length
  }

  $dirStream = New-Object System.IO.MemoryStream
  foreach ($ne in $newEntries) {
    $tagBytes = [System.Text.Encoding]::ASCII.GetBytes($ne.Tag)
    $dirStream.Write($tagBytes, 0, 4)
    $buf4 = New-Object byte[] 4
    Write-UInt32BE $buf4 0 ([uint32]$ne.Checksum); $dirStream.Write($buf4, 0, 4)
    Write-UInt32BE $buf4 0 ([uint32]$ne.Offset); $dirStream.Write($buf4, 0, 4)
    Write-UInt32BE $buf4 0 ([uint32]$ne.Length); $dirStream.Write($buf4, 0, 4)
  }

  $outStream = New-Object System.IO.MemoryStream
  $headerBytes = $headerStream.ToArray()
  $outStream.Write($headerBytes, 0, $headerBytes.Length)
  $dirBytes = $dirStream.ToArray()
  $outStream.Write($dirBytes, 0, $dirBytes.Length)
  foreach ($chunk in $bodyChunks) { $outStream.Write($chunk, 0, $chunk.Length) }

  $fontBytes = $outStream.ToArray()

  $headIdx = -1
  for ($i = 0; $i -lt $newEntries.Count; $i++) { if ($newEntries[$i].Tag -eq "head") { $headIdx = $i; break } }
  if ($headIdx -ge 0) {
    $headOff = $newEntries[$headIdx].Offset
    $fontBytes[$headOff + 8] = 0; $fontBytes[$headOff + 9] = 0; $fontBytes[$headOff + 10] = 0; $fontBytes[$headOff + 11] = 0
    $checksum = Get-SfntChecksum $fontBytes
    # 0xB1B0AFBA(10進2981146554)はSFNT "head"テーブルのcheckSumAdjustment用
    # マジックナンバー。16進リテラルのまま書くとPowerShellがInt32の負数
    # (-1313820742)として解釈し、減算結果がuint32の範囲を超えてキャスト例外に
    # なるため、10進で書いて[uint64]で計算し4294967296の剰余で32bitに収める。
    $adjustment = ((2981146554 - [uint64]$checksum) + 4294967296) % 4294967296
    Write-UInt32BE $fontBytes ($headOff + 8) ([uint32]$adjustment)
  }

  return $fontBytes
}

function Resolve-SystemFont([string]$Key) {
  if (-not $FontDefs.ContainsKey($Key)) { return $null, "unknown font key" }
  $def = $FontDefs[$Key]

  $regFiles = @{}
  try {
    $props = Get-ItemProperty -Path $FontsRegKey -ErrorAction Stop
    foreach ($p in $props.PSObject.Properties) {
      if ($p.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$') { $regFiles[$p.Value.ToString().ToLower()] = $true }
    }
  } catch { }

  $path = $null
  foreach ($cand in $def.Candidates) {
    if ($regFiles.ContainsKey($cand.ToLower())) { $path = Join-Path $WindowsFontsDir $cand; break }
  }
  if (-not $path) {
    foreach ($cand in $def.Candidates) {
      $p = Join-Path $WindowsFontsDir $cand
      if (Test-Path $p -PathType Leaf) { $path = $p; break }
    }
  }
  if (-not $path) { return $null, "font file not found on this PC" }

  try {
    $buf = [System.IO.File]::ReadAllBytes($path)
  } catch {
    return $null, $_.Exception.Message
  }

  $idx = Find-FaceIndex $buf $def.Family
  if ($idx -lt 0) { return $null, "expected font face not found inside $path" }

  try {
    $fontBytes = Extract-Face $buf $idx
    return $fontBytes, $null
  } catch {
    return $null, "failed to read font: $($_.Exception.Message)"
  }
}

# ----------------------------------------------------------------------
# セキュリティ: 自分自身のページ以外からのリクエストを拒否する。
# ----------------------------------------------------------------------
function Test-FromThisToolsOwnPage([System.Net.HttpListenerRequest]$Request) {
  $expected = "http://127.0.0.1:$Port"
  $origin = $Request.Headers["Origin"]
  if ($origin) { return $origin -eq $expected }
  $referer = $Request.Headers["Referer"]
  if ($referer) { return ($referer -eq $expected) -or $referer.StartsWith("$expected/") -or $referer.StartsWith("$expected?") }
  return $true
}

function Send-JsonResponse([System.Net.HttpListenerResponse]$Response, [int]$StatusCode, [hashtable]$Obj) {
  $json = $Obj | ConvertTo-Json -Compress -Depth 10
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
  $Response.StatusCode = $StatusCode
  $Response.ContentType = "application/json; charset=utf-8"
  $Response.Headers.Add("Cache-Control", "no-store")
  $Response.ContentLength64 = $bytes.Length
  $Response.OutputStream.Write($bytes, 0, $bytes.Length)
  $Response.OutputStream.Close()
}

function Send-Error([System.Net.HttpListenerResponse]$Response, [int]$StatusCode, [string]$Message, [System.Net.HttpListenerRequest]$Request) {
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Message)
  $Response.StatusCode = $StatusCode
  $Response.ContentType = "text/plain; charset=utf-8"
  $Response.ContentLength64 = $bytes.Length
  # HEADは本文を書かない(Handle-StaticFileと同じ理由 -- 上のコメント参照)。
  if (-not $Request -or $Request.HttpMethod -ne "HEAD") {
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
  }
  $Response.OutputStream.Close()
}

function Handle-Load([System.Net.HttpListenerRequest]$Request, [System.Net.HttpListenerResponse]$Response) {
  $qs = [System.Web.HttpUtility]::ParseQueryString($Request.Url.Query)
  $token = $qs["open"]
  if (-not $token -or $token -notmatch '^\d+$') { Send-Error $Response 400 "invalid token" $Request; return }

  $tokenFile = Join-Path ([System.IO.Path]::GetTempPath()) "pdf_editor_$token.path"
  if (-not (Test-Path $tokenFile -PathType Leaf)) { Send-Error $Response 404 "token expired or not found" $Request; return }
  try {
    # utf-8-sig: ADODB.Streamがcharset "utf-8"で保存すると常にBOM付きになる
    $pdfPath = [System.IO.File]::ReadAllText($tokenFile, [System.Text.Encoding]::UTF8)
  } finally {
    Remove-Item $tokenFile -Force -ErrorAction SilentlyContinue
  }

  if (-not $pdfPath.ToLower().EndsWith(".pdf") -or -not (Test-Path $pdfPath -PathType Leaf)) {
    Send-Error $Response 400 "not a pdf file" $Request; return
  }

  try {
    $data = [System.IO.File]::ReadAllBytes($pdfPath)
  } catch {
    Send-Error $Response 500 $_.Exception.Message $Request; return
  }

  [void]$Script:ServedPaths.Add($pdfPath)
  # ダウンロードボタンが元ファイルと同じフォルダへ保存する際に使う
  # "_編集済み"派生パスも同時に事前登録しておく(index.htmlのexportPdf()/
  # sameFolderSavePath()参照) -- /__save自身の許可リストはこのままで、
  # ここで登録するパスが増えるだけ。
  $ext = [System.IO.Path]::GetExtension($pdfPath)
  $base = $pdfPath.Substring(0, $pdfPath.Length - $ext.Length)
  [void]$Script:ServedPaths.Add($base + "_編集済み" + $ext)

  Send-JsonResponse $Response 200 @{
    filename = [System.IO.Path]::GetFileName($pdfPath)
    path     = $pdfPath
    data     = [Convert]::ToBase64String($data)
  }
}

function Handle-Font([System.Net.HttpListenerRequest]$Request, [System.Net.HttpListenerResponse]$Response) {
  $qs = [System.Web.HttpUtility]::ParseQueryString($Request.Url.Query)
  # nameパラメータ自体が無いと$qs["name"]は$nullを返し、Hashtable.ContainsKey($null)
  # は(存在しないキーとしてFalseを返すのではなく)例外を投げる -- 先に
  # $nullチェックしてから渡す。
  $key = $qs["name"]
  if (-not $key -or -not $FontDefs.ContainsKey($key)) { Send-Error $Response 400 "unknown font key" $Request; return }

  if ($Script:FontCache.ContainsKey($key)) {
    $data = $Script:FontCache[$key]
  } else {
    $data, $err = Resolve-SystemFont $key
    if ($err) { Send-Error $Response 404 $err $Request; return }
    $Script:FontCache[$key] = $data
  }

  Send-JsonResponse $Response 200 @{ name = $key; data = [Convert]::ToBase64String($data) }
}

function Handle-Save([System.Net.HttpListenerRequest]$Request, [System.Net.HttpListenerResponse]$Response) {
  if ($Request.ContentLength64 -le 0 -or $Request.ContentLength64 -gt 200MB) { Send-Error $Response 400 "invalid content length" $Request; return }

  $reader = New-Object System.IO.StreamReader($Request.InputStream, [System.Text.Encoding]::UTF8)
  $bodyText = $reader.ReadToEnd()
  $reader.Close()

  try {
    $body = $bodyText | ConvertFrom-Json
    $pdfPath = $body.path
    $data = [Convert]::FromBase64String($body.data)
  } catch {
    Send-Error $Response 400 "malformed request" $Request; return
  }

  # do_POSTの直後ではなく、ボディを読み切った後にここで確認する
  # (server.py側のdo_POST自身のコメントを参照 -- ボディを読まずに閉じると
  # OS側がTCP RSTを送り得て、既に書いたはずの403応答が消えてしまう場合が
  # ある。ここでは順序が違ってもPowerShellのHttpListenerが同様の問題を
  # 起こすかは未検証だが、念のため同じ順序を踏襲している)。
  if (-not (Test-FromThisToolsOwnPage $Request)) { Send-Error $Response 403 "forbidden" $Request; return }

  if (-not $Script:ServedPaths.Contains($pdfPath)) { Send-Error $Response 403 "unknown path" $Request; return }
  if (-not $pdfPath.ToLower().EndsWith(".pdf")) { Send-Error $Response 400 "not a pdf path" $Request; return }

  $tmpPath = "$pdfPath.tmp-$PID-$([guid]::NewGuid().ToString('N'))"
  try {
    $bakPath = "$pdfPath.bak"
    if ((Test-Path $pdfPath -PathType Leaf) -and -not (Test-Path $bakPath)) {
      Copy-Item -Path $pdfPath -Destination $bakPath -Force
    }
    [System.IO.File]::WriteAllBytes($tmpPath, $data)
    # 一時ファイル書き込み+アトミックな置き換え -- 書き込み中のクラッシュや
    # ディスクフルでPDFが壊れた状態のまま残らないようにするため。参照実装の
    # server.pyはos.replace()(Windows上ではReplaceFile/MoveFileExによる
    # 真にアトミックな置換)を使っているが、.NET Framework(Windows
    # PowerShell 5.1)にはos.replace()に相当する単一APIがない --
    # [System.IO.File]::Move()は既存の宛先があると例外を投げるだけで
    # 置き換えてくれず、Delete()してからMove()すると、その間に例外
    # (AVによる一時ファイルのロック、権限エラー等)が起きた場合に
    # オリジナルもコピーもどちらも存在しない状態になり得る(削除だけ
    # 成功してMoveが失敗するケース)。File.Replace()は宛先が既存の場合に
    # 限りReplaceFile Win32 APIで真にアトミックに置き換えるので、宛先の
    # 有無で使い分けて常にアトミックな置換になるようにする。
    if ([System.IO.File]::Exists($pdfPath)) {
      # 注意: ここで$nullをそのまま渡すと"パスの形式が無効です"で例外になる --
      # PowerShellがこの[string]引数へのバインド時に$nullを空文字列へ変換
      # してしまい、File.Replace()は(真のnullなら「バックアップ不要」と
      # 解釈するのに)空文字列だと不正なパスとして拒否するため。実機で
      # 再現・特定済み。[NullString]::Valueを使うとPowerShellの$null→
      # 空文字列変換を回避して真のnullを渡せる。
      [System.IO.File]::Replace($tmpPath, $pdfPath, [NullString]::Value)
    } else {
      # ダウンロードボタンが同じフォルダへ保存する"_編集済み"派生パスは、
      # 初回保存時点ではまだ存在しない -- その場合はFile.Replace()が
      # 使えない(宛先の存在を要求する)ので、単純なMove()にフォール
      # バックする(宛先が存在しない場合、Move()自体が同一ボリューム内
      # では単一のリネーム操作としてアトミック)。
      [System.IO.File]::Move($tmpPath, $pdfPath)
    }
  } catch {
    Remove-Item $tmpPath -Force -ErrorAction SilentlyContinue
    Send-Error $Response 500 $_.Exception.Message $Request
    return
  }

  Send-JsonResponse $Response 200 @{ ok = $true }
}

function Handle-StaticFile([System.Net.HttpListenerRequest]$Request, [System.Net.HttpListenerResponse]$Response) {
  $relPath = [Uri]::UnescapeDataString($Request.Url.AbsolutePath.TrimStart('/'))
  if ([string]::IsNullOrEmpty($relPath)) { $relPath = "index.html" }
  $fullPath = Join-Path $ToolDir $relPath
  # ディレクトリトラバーサル対策 -- 解決後のパスが必ずToolDir配下にあることを確認
  $resolvedFull = [System.IO.Path]::GetFullPath($fullPath)
  $resolvedRoot = [System.IO.Path]::GetFullPath($ToolDir)
  if (-not $resolvedFull.StartsWith($resolvedRoot, [StringComparison]::OrdinalIgnoreCase)) {
    Send-Error $Response 403 "forbidden" $Request; return
  }
  if (-not (Test-Path $resolvedFull -PathType Leaf)) { Send-Error $Response 404 "not found" $Request; return }

  $ext = [System.IO.Path]::GetExtension($resolvedFull).ToLower()
  $contentType = switch ($ext) {
    ".html" { "text/html; charset=utf-8" }
    ".js"   { "text/javascript; charset=utf-8" }
    ".css"  { "text/css; charset=utf-8" }
    ".bin"  { "application/octet-stream" }
    ".ico"  { "image/x-icon" }
    default { "application/octet-stream" }
  }
  $bytes = [System.IO.File]::ReadAllBytes($resolvedFull)
  $Response.StatusCode = 200
  $Response.ContentType = $contentType
  $Response.ContentLength64 = $bytes.Length
  # HEADはヘッダーのみを返す(本文を書かない) -- HttpListenerResponseは
  # リクエストメソッドを見て自動で本文を抑制したりはしないため、ここで
  # 明示的に分岐する必要がある。抑制しないと、HEADで問い合わせたクライアント
  # (open_in_pdf_editor.ps1の起動待ちポーリング等)が応答をハングしたように
  # 誤認する。
  if ($Request.HttpMethod -ne "HEAD") {
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
  }
  $Response.OutputStream.Close()
}

# ----------------------------------------------------------------------
# メインループ
# ----------------------------------------------------------------------
Add-Type -AssemblyName System.Web

$listener = New-Object System.Net.HttpListener
$prefix = "http://127.0.0.1:$Port/"
$listener.Prefixes.Add($prefix)
try {
  $listener.Start()
} catch [System.Net.HttpListenerException] {
  # ErrorCode 183 (ERROR_ALREADY_EXISTS) = 既に別プロセスが同じURLプレフィックス
  # を登録済み -- 以前の「送る」起動が生きているということなので、Python版の
  # ReusableTCPServer(WinError 10048チェック)と同じ考え方で、この「想定内の
  # 1パターンだけ」を静かに終了させる。実機で`$_.Exception.ErrorCode`を
  # 確認して183であることを検証済み(Python版のWSAEADDRINUSE=10048とは
  # 別の値 -- HttpListenerはソケットではなくHTTP.SYS経由のため)。
  if ($_.Exception.ErrorCode -eq 183) {
    exit 0
  }
  # それ以外は本当に想定外の失敗(URL ACL不足、無効なプレフィックス等) --
  # server.pyが「この1パターン以外は再送出して見えるダイアログを出す」の
  # と同じ理由で、静かに諦めず利用者に伝える。
  $shell = New-Object -ComObject WScript.Shell
  $shell.Popup("橋渡しサーバーの起動に失敗しました。`n`n" + $_.Exception.Message, 0, "PDF Editor Tool", 16) | Out-Null
  exit 1
} catch {
  $shell = New-Object -ComObject WScript.Shell
  $shell.Popup("橋渡しサーバーの起動に失敗しました。`n`n" + $_.Exception.Message, 0, "PDF Editor Tool", 16) | Out-Null
  exit 1
}

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response
    try {
      $path = $request.Url.AbsolutePath
      if ($request.HttpMethod -eq "GET" -and $path -eq "/__load") {
        if (-not (Test-FromThisToolsOwnPage $request)) { Send-Error $response 403 "forbidden" $request } else { Handle-Load $request $response }
      } elseif ($request.HttpMethod -eq "GET" -and $path -eq "/__font") {
        if (-not (Test-FromThisToolsOwnPage $request)) { Send-Error $response 403 "forbidden" $request } else { Handle-Font $request $response }
      } elseif ($request.HttpMethod -eq "POST" -and $path -eq "/__save") {
        Handle-Save $request $response
      } elseif ($request.HttpMethod -eq "GET" -or $request.HttpMethod -eq "HEAD") {
        # HEADは、open_in_pdf_editor.ps1(旧vbs)がサーバーの起動完了を確認する
        # ためだけに使う -- 本文を丸ごと転送せずに済ませる意図なので、
        # Handle-StaticFile側でも本文の書き込みを省略する必要がある。
        Handle-StaticFile $request $response
      } else {
        Send-Error $response 404 "not found" $request
      }
    } catch {
      try { Send-Error $response 500 $_.Exception.Message $request } catch { }
    }
  }
} finally {
  $listener.Stop()
  $listener.Close()
}
