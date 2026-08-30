# -----------------------------------------------------------------------------
#  AI 繪本電子書書架 — 本機管理小工具
#  新增：複製電子書檔案到系列資料夾、抽出第一頁當封面、寫進 books.json
#  刪除：從 books.json 移除，可一併刪掉封面與電子書檔案
#  請用同資料夾的「管理書架.bat」啟動（也可以把電子書檔案拖到 .bat 上面）
# -----------------------------------------------------------------------------
param([string[]]$Files)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$Root     = Split-Path -Parent $MyInvocation.MyCommand.Path
$JsonPath = Join-Path $Root 'books.json'
$IndexPath= Join-Path $Root 'index.html'
$CoverDir = Join-Path $Root 'covers'

function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }
function Line() { Say ('-' * 62) 'DarkGray' }

# ---------------------------------------------------------------- JSON 讀寫 --
function Read-Books {
    if (-not (Test-Path $JsonPath)) { return @() }
    $raw = [System.IO.File]::ReadAllText($JsonPath, [System.Text.Encoding]::UTF8)
    $obj = $raw | ConvertFrom-Json
    foreach ($b in $obj.books) {
        [pscustomobject]@{
            id = [string]$b.id; title = [string]$b.title; group = [string]$b.group
            lang = [string]$b.lang; collection = [string]$b.collection
            pages = [int]$b.pages; path = [string]$b.path; cover = [string]$b.cover
            hidden = [bool]$b.hidden
        }
    }
}

function Read-Cols {
    if (-not (Test-Path $JsonPath)) { return }
    $raw = [System.IO.File]::ReadAllText($JsonPath, [System.Text.Encoding]::UTF8)
    $obj = $raw | ConvertFrom-Json
    foreach ($c in $obj.collections) {
        if (-not $c.name) { continue }
        [pscustomobject]@{ name = [string]$c.name; folder = [string]$c.folder }
    }
}

function Sync-Cols($cols, $books) {
    $out = @()
    foreach ($c in $cols) {
        if (-not $c.name) { continue }
        if ($out | Where-Object { $_.name -eq $c.name }) { continue }
        $folder = $c.folder
        if (-not $folder) { $folder = $c.name }
        $out += [pscustomobject]@{ name = [string]$c.name; folder = [string]$folder }
    }
    foreach ($b in $books) {
        if (-not $b.collection) { continue }
        if ($out | Where-Object { $_.name -eq $b.collection }) { continue }
        $f = Split-Path $b.path -Parent
        if (-not $f) { $f = $b.collection }
        $out += [pscustomobject]@{ name = [string]$b.collection; folder = ($f -replace [regex]::Escape([string][char]92), '/') }
    }
    foreach ($o in $out) { $o }
}

function Esc([string]$s) {
    if ($null -eq $s) { $s = '' }
    $s = $s.Replace('\', '\\').Replace('"', '\"')
    $s = $s.Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t')
    return '"' + $s + '"'
}

function Build-Json($books, $cols) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('{')
    [void]$sb.AppendLine('  "updated": ' + (Esc (Get-Date -Format 'yyyy-MM-dd')) + ',')
    [void]$sb.AppendLine('  "collections": [')
    for ($i = 0; $i -lt $cols.Count; $i++) {
        [void]$sb.AppendLine('    {')
        [void]$sb.AppendLine('      "name": ' + (Esc $cols[$i].name) + ',')
        [void]$sb.AppendLine('      "folder": ' + (Esc $cols[$i].folder))
        if ($i -lt $cols.Count - 1) { [void]$sb.AppendLine('    },') } else { [void]$sb.AppendLine('    }') }
    }
    [void]$sb.AppendLine('  ],')
    [void]$sb.AppendLine('  "books": [')
    for ($i = 0; $i -lt $books.Count; $i++) {
        $b = $books[$i]
        $rows = @(
            '      "id": '         + (Esc $b.id)
            '      "title": '      + (Esc $b.title)
            '      "group": '      + (Esc $b.group)
            '      "lang": '       + (Esc $b.lang)
            '      "collection": ' + (Esc $b.collection)
            '      "pages": '      + [int]$b.pages
            '      "path": '       + (Esc $b.path)
            '      "cover": '      + (Esc $b.cover)
            '      "hidden": '     + $(if ($b.hidden) { 'true' } else { 'false' })
        )
        [void]$sb.AppendLine('    {')
        [void]$sb.AppendLine(($rows -join ",`r`n"))
        if ($i -lt $books.Count - 1) { [void]$sb.AppendLine('    },') } else { [void]$sb.AppendLine('    }') }
    }
    [void]$sb.AppendLine('  ]')
    [void]$sb.Append('}')
    return $sb.ToString()
}

function Save-Books($books, $cols) {
    $json = Build-Json $books $cols
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($JsonPath, $json, $utf8)

    # 同步 index.html 裡的備援書單（給 file:// 直接開啟時用）
    if (Test-Path $IndexPath) {
        $html  = [System.IO.File]::ReadAllText($IndexPath, [System.Text.Encoding]::UTF8)
        $open  = '<script type="application/json" id="fallback">'
        $s = $html.IndexOf($open)
        if ($s -ge 0) {
            $s2 = $s + $open.Length
            $e  = $html.IndexOf('</script>', $s2)
            if ($e -gt $s2) {
                $html = $html.Substring(0, $s2) + $json.Replace('</', '<\/') + $html.Substring($e)
                [System.IO.File]::WriteAllText($IndexPath, $html, $utf8)
            }
        }
    }
}

# ------------------------------------------------------------ 讀電子書內容 --
function Read-Ebook([string]$path) {
    $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $re   = [regex]'data:image/(webp|png|jpeg|jpg);base64,([A-Za-z0-9+/=]+)'

    $m = $null
    $i = $text.IndexOf('const THUMBS')
    if ($i -ge 0) { $m = $re.Match($text, $i) }
    if (-not $m -or -not $m.Success) {
        $i = $text.IndexOf('const PAGES')
        if ($i -ge 0) { $m = $re.Match($text, $i) }
    }
    if (-not $m -or -not $m.Success) { $m = $re.Match($text) }
    if (-not $m.Success) { throw '這個檔案裡找不到頁面圖片，看起來不是翻頁書電子書。' }

    $ext = $m.Groups[1].Value
    if ($ext -eq 'jpeg') { $ext = 'jpg' }

    $pages = 0
    $pm = [regex]::Match($text, 'const PAGES\s*=\s*\[')
    if ($pm.Success) {
        $s = $pm.Index + $pm.Length
        $e = $text.IndexOf('];', $s)
        if ($e -gt $s) { $pages = ([regex]::Matches($text.Substring($s, $e - $s), 'data:image/')).Count }
    }

    $innerTitle = ''
    $tm = [regex]::Match($text, 'const TITLE\s*=\s*"([^"]*)"')
    if ($tm.Success -and $tm.Groups[1].Value -notlike '*我的翻頁書*') { $innerTitle = $tm.Groups[1].Value }

    return [pscustomobject]@{
        Bytes = [Convert]::FromBase64String($m.Groups[2].Value)
        Ext   = $ext
        Pages = $pages
        Title = $innerTitle
    }
}

function Guess-Meta([string]$fileName, [string]$innerTitle) {
    $name  = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $parts = $name -split '[_＿]'
    $group = ''
    if ($parts[0] -match '^\s*(第.{1,2}組|GROUP\s*\d+)\s*$') { $group = $parts[0].Trim() }
    $lang = ''
    if ($parts.Count -gt 2) { $lang = $parts[$parts.Count - 1].Trim() }

    $title = $name
    if ($innerTitle) { $title = $innerTitle }
    elseif ($group) {
        $mid = if ($parts.Count -gt 2) { $parts[1..($parts.Count - 2)] } else { $parts[1..($parts.Count - 1)] }
        if ($mid) { $title = ($mid -join ' ').Trim() }
    }
    return [pscustomobject]@{ Title = $title; Group = $group; Lang = $lang }
}

function Ask([string]$label, [string]$default) {
    if ($default) { $a = Read-Host "$label [$default]" } else { $a = Read-Host $label }
    if ([string]::IsNullOrWhiteSpace($a)) { return $default }
    return $a.Trim()
}

function New-Id([array]$books) {
    do {
        $id = 'bk-' + (Get-Date -Format 'yyMMddHHmmss') + '-' + (Get-Random -Minimum 100 -Maximum 999)
    } while ($books | Where-Object { $_.id -eq $id })
    return $id
}

function Pick-Files {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = '選擇要加入書架的電子書 HTML 檔（可複選）'
    $dlg.Filter = '電子書 (*.html;*.htm)|*.html;*.htm'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileNames }
    return @()
}

# ---------------------------------------------------------------- 新增流程 --
function Add-Book([string]$src, [ref]$booksRef, [ref]$colsRef) {
    $books = $booksRef.Value
    Line
    Say "來源檔案：$src" 'White'

    $info = Read-Ebook $src
    $meta = Guess-Meta ([System.IO.Path]::GetFileName($src)) $info.Title
    Say ("已讀出第 1 頁封面，全書 {0} 頁" -f $info.Pages) 'Green'
    Say ''

    # 類別（系列）
    $cols = $colsRef.Value
    Say '要放進哪個類別？' 'White'
    for ($i = 0; $i -lt $cols.Count; $i++) {
        $n = @($books | Where-Object { $_.collection -eq $cols[$i].name }).Count
        Say ("  [{0}] {1}   （{2} 本）  →  {3}" -f ($i + 1), $cols[$i].name, $n, $cols[$i].folder)
    }
    Say ("  [N] 新增一個類別（會建立新資料夾）")
    $sel = Read-Host '請輸入編號'

    if ($sel -match '^[Nn]$') {
        $colName = Ask '新類別名稱（顯示在書架上）' ''
        if (-not $colName) { Say '沒有輸入名稱，取消這一本。' 'Yellow'; return }
        if ($cols | Where-Object { $_.name -eq $colName }) { Say '已經有同名的類別。' 'Yellow'; return }
        $folder = Ask '資料夾名稱' ("AI繪本-$colName(電子書)")
        $colsRef.Value = @($cols) + [pscustomobject]@{ name = $colName; folder = $folder }
    } else {
        $k = 0
        if (-not [int]::TryParse($sel, [ref]$k) -or $k -lt 1 -or $k -gt $cols.Count) {
            Say '編號不正確，取消這一本。' 'Yellow'; return
        }
        $colName = $cols[$k - 1].name
        $folder  = $cols[$k - 1].folder
    }

    # 書名等欄位
    Say ''
    $title = Ask '書名'        $meta.Title
    $group = Ask '組別（可留空）' $meta.Group
    $lang  = Ask '語言（可留空）' $meta.Lang

    # 複製檔案
    $destDir = Join-Path $Root $folder
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null; Say "已建立資料夾 $folder" 'Green' }

    $fileName = [System.IO.Path]::GetFileName($src)
    $dest     = Join-Path $destDir $fileName
    $srcFull  = (Resolve-Path $src).Path

    if ((Test-Path $dest) -and ((Resolve-Path $dest).Path -eq $srcFull)) {
        Say '檔案已經在這個資料夾裡，不需要複製。' 'DarkGray'
    } elseif (Test-Path $dest) {
        Say "注意：$folder/$fileName 已經存在。" 'Yellow'
        if ((Read-Host '要覆蓋嗎？(y/N)') -match '^[Yy]') {
            Copy-Item $srcFull $dest -Force; Say '已覆蓋。' 'Green'
        } else { Say '保留原檔，不複製。' 'DarkGray' }
    } else {
        Copy-Item $srcFull $dest
        Say ("已複製電子書 → {0}/{1}  ({2:N1} MB)" -f $folder, $fileName, ((Get-Item $dest).Length / 1MB)) 'Green'
    }

    # 封面
    if (-not (Test-Path $CoverDir)) { New-Item -ItemType Directory -Path $CoverDir | Out-Null }
    $relPath = "$folder/$fileName"
    $exist   = $books | Where-Object { $_.path -eq $relPath } | Select-Object -First 1
    $id      = if ($exist) { $exist.id } else { New-Id $books }
    $coverRel= "covers/$id.$($info.Ext)"
    [System.IO.File]::WriteAllBytes((Join-Path $Root $coverRel), $info.Bytes)
    Say "已存下封面 → $coverRel" 'Green'

    $entry = [pscustomobject]@{
        id = $id; title = $title; group = $group; lang = $lang
        collection = $colName; pages = $info.Pages; path = $relPath; cover = $coverRel
        hidden = $(if ($exist) { [bool]$exist.hidden } else { $false })
    }

    if ($exist) {
        $arr = @(); foreach ($b in $books) { if ($b.path -eq $relPath) { $arr += $entry } else { $arr += $b } }
        $booksRef.Value = $arr
        Say "書架上已有同路徑的書，改成更新這一筆。" 'Yellow'
    } else {
        $booksRef.Value = @($books) + $entry
    }
    Say "『$title』已加入書架。" 'Cyan'
}

# ------------------------------------------------------------ 顯示／隱藏 --
function Show-List($books) {
    Say '書架上的書：' 'White'
    for ($i = 0; $i -lt $books.Count; $i++) {
        $b = $books[$i]
        $mark = if ($b.hidden) { '隱藏' } else { '顯示' }
        $color = if ($b.hidden) { 'DarkGray' } else { 'Gray' }
        Say ("  [{0,2}] {1}  {2,-22} {3,-8} {4,-6} {5}" -f ($i + 1), $mark, $b.title, $b.collection, $b.group, $b.lang) $color
    }
}

function Toggle-Book([ref]$booksRef) {
    $books = $booksRef.Value
    if ($books.Count -eq 0) { Say '書架上沒有書。' 'Yellow'; return }
    Line
    Show-List $books
    Say '（「隱藏」的書不會出現在前台書架，但檔案仍在 GitHub 上，知道網址還是打得開）' 'DarkGray'
    $sel = Read-Host '要切換第幾本的顯示狀態？（直接按 Enter 取消）'
    $k = 0
    if (-not [int]::TryParse($sel, [ref]$k) -or $k -lt 1 -or $k -gt $books.Count) { Say '取消。' 'DarkGray'; return }
    $b = $books[$k - 1]
    $b.hidden = -not $b.hidden
    if ($b.hidden) { Say ("『{0}』已設為隱藏，前台看不到。" -f $b.title) 'Yellow' }
    else           { Say ("『{0}』已恢復顯示。" -f $b.title) 'Cyan' }
    $booksRef.Value = $books
}

# ---------------------------------------------------------------- 類別管理 --
function Manage-Cols([ref]$booksRef, [ref]$colsRef) {
    while ($true) {
        $books = $booksRef.Value
        $cols  = $colsRef.Value
        Line
        Say '目前的類別：' 'White'
        for ($i = 0; $i -lt $cols.Count; $i++) {
            $n = @($books | Where-Object { $_.collection -eq $cols[$i].name }).Count
            Say ("  [{0}] {1,-14} {2,3} 本   資料夾：{3}" -f ($i + 1), $cols[$i].name, $n, $cols[$i].folder)
        }
        Say '  [A] 新增類別    [E] 修改名稱／資料夾    [D] 刪除沒有書的類別    [Enter] 返回'
        $c = Read-Host '請選擇'

        if ($c -match '^[Aa]$') {
            $name = Ask '類別名稱（顯示在書架上）' ''
            if (-not $name) { Say '取消。' 'DarkGray'; continue }
            if ($cols | Where-Object { $_.name -eq $name }) { Say '已經有同名的類別。' 'Yellow'; continue }
            $folder = Ask '資料夾名稱' ("AI繪本-$name(電子書)")
            $colsRef.Value = @($cols) + [pscustomobject]@{ name = $name; folder = $folder }
            $dir = Join-Path $Root $folder
            if (-not (Test-Path $dir)) {
                if ((Read-Host "要現在建立資料夾 $folder 嗎？(Y/n)") -notmatch '^[Nn]') {
                    New-Item -ItemType Directory -Path $dir | Out-Null
                    Say "已建立資料夾 $folder" 'Green'
                }
            }
            Say "已新增類別『$name』。" 'Cyan'
        }
        elseif ($c -match '^[Ee]$') {
            $sel = Read-Host '要修改第幾個類別？'
            $k = 0
            if (-not [int]::TryParse($sel, [ref]$k) -or $k -lt 1 -or $k -gt $cols.Count) { Say '取消。' 'DarkGray'; continue }
            $target = $cols[$k - 1]
            $oldName = $target.name
            $oldFolder = $target.folder
            $newName = Ask '類別名稱' $oldName
            $newFolder = Ask '資料夾名稱' $oldFolder
            if ($newName -ne $oldName) {
                if ($cols | Where-Object { $_.name -eq $newName }) { Say '已經有同名的類別。' 'Yellow'; continue }
            }
            $n = @($books | Where-Object { $_.collection -eq $oldName }).Count
            if ($newFolder -ne $oldFolder -and $n -gt 0) {
                Say ("這個類別底下有 {0} 本書，路徑會一起改到 {1}/" -f $n, $newFolder) 'Yellow'
                $dir = Join-Path $Root $newFolder
                if (Test-Path $dir) {
                    Say '新資料夾已存在。' 'DarkGray'
                    if ((Read-Host '要把電子書檔案搬過去嗎？(Y/n)') -notmatch '^[Nn]') { $doMove = $true } else { $doMove = $false }
                } else {
                    if ((Read-Host ("要建立 {0} 並把電子書搬過去嗎？(Y/n)" -f $newFolder)) -notmatch '^[Nn]') {
                        New-Item -ItemType Directory -Path $dir | Out-Null
                        $doMove = $true
                    } else { $doMove = $false }
                }
            } else { $doMove = $false }

            foreach ($b in $books) {
                if ($b.collection -ne $oldName) { continue }
                $b.collection = $newName
                $base = Split-Path $b.path -Leaf
                $newPath = "$newFolder/$base"
                if ($doMove) {
                    $from = Join-Path $Root ($b.path -replace '/', [string][char]92)
                    $to   = Join-Path $Root ($newPath -replace '/', [string][char]92)
                    if ((Test-Path $from) -and -not (Test-Path $to)) {
                        Move-Item $from $to
                        Say ("已搬移 {0}" -f $base) 'DarkGray'
                    }
                }
                $b.path = $newPath
            }
            $target.name = $newName
            $target.folder = $newFolder
            $colsRef.Value = $cols
            $booksRef.Value = $books
            Say ("已更新類別『{0}』。" -f $newName) 'Cyan'
        }
        elseif ($c -match '^[Dd]$') {
            $sel = Read-Host '要刪除第幾個類別？'
            $k = 0
            if (-not [int]::TryParse($sel, [ref]$k) -or $k -lt 1 -or $k -gt $cols.Count) { Say '取消。' 'DarkGray'; continue }
            $target = $cols[$k - 1]
            $n = @($books | Where-Object { $_.collection -eq $target.name }).Count
            if ($n -gt 0) {
                Say ("『{0}』底下還有 {1} 本書，請先移除或改到別的類別。" -f $target.name, $n) 'Yellow'
                continue
            }
            if ((Read-Host ("確定刪除類別『{0}』？(y/N)" -f $target.name)) -notmatch '^[Yy]') { Say '取消。' 'DarkGray'; continue }
            $colsRef.Value = @($cols | Where-Object { $_.name -ne $target.name })
            Say ("已刪除類別『{0}』（資料夾沒有動）。" -f $target.name) 'Cyan'
        }
        else { break }
    }
}

# ---------------------------------------------------------------- 刪除流程 --
function Remove-Book([ref]$booksRef) {
    $books = $booksRef.Value
    if ($books.Count -eq 0) { Say '書架上沒有書。' 'Yellow'; return }

    Line
    Show-List $books
    $sel = Read-Host '要移除第幾本？（直接按 Enter 取消）'
    $k = 0
    if (-not [int]::TryParse($sel, [ref]$k) -or $k -lt 1 -or $k -gt $books.Count) { Say '取消。' 'DarkGray'; return }

    $b = $books[$k - 1]
    Line
    Say ("書名：{0}" -f $b.title) 'White'
    Say ("電子書：{0}" -f $b.path)
    Say ("封面：  {0}" -f $b.cover)
    if ((Read-Host '確定要從書架移除嗎？(y/N)') -notmatch '^[Yy]') { Say '取消。' 'DarkGray'; return }

    $coverFull = Join-Path $Root ($b.cover -replace '/', '\')
    if ($b.cover -and (Test-Path $coverFull)) {
        if ((Read-Host '順便刪掉封面檔？(Y/n)') -notmatch '^[Nn]') { Remove-Item $coverFull -Force; Say '已刪除封面檔。' 'DarkGray' }
    }
    $bookFull = Join-Path $Root ($b.path -replace '/', '\')
    if (Test-Path $bookFull) {
        Say ("電子書檔案本身：{0}" -f $bookFull) 'Yellow'
        if ((Read-Host '要一併刪除這個電子書檔案嗎？(y/N)') -match '^[Yy]') { Remove-Item $bookFull -Force; Say '已刪除電子書檔案。' 'Yellow' }
        else { Say '保留電子書檔案。' 'DarkGray' }
    }

    $booksRef.Value = @($books | Where-Object { $_.id -ne $b.id })
    Say ("『{0}』已從書架移除。" -f $b.title) 'Cyan'
}

# -------------------------------------------------------------------- 主流程 --
if (-not (Test-Path $JsonPath)) {
    Say "找不到 books.json，請確認這個工具和 index.html 放在同一個資料夾。" 'Red'
    Read-Host '按 Enter 結束' | Out-Null
    exit 1
}

$books = @(Read-Books)
$cols  = @(Sync-Cols (Read-Cols) $books)
$before = Build-Json $books $cols

Say ''
Say '  📖 AI 繪本電子書書架 — 管理工具' 'Cyan'
Say ("  資料夾：{0}" -f $Root) 'DarkGray'
$hiddenCount = @($books | Where-Object { $_.hidden }).Count
Say ("  目前書架上有 {0} 本書{1}，共 {2} 個類別" -f $books.Count, $(if ($hiddenCount) { "（其中 $hiddenCount 本在前台隱藏）" } else { '' }), $cols.Count) 'DarkGray'

# 拖到 .bat 上面的檔案：直接進新增流程
$dropped = @()
foreach ($f in $Files) { if ($f -and (Test-Path $f)) { $dropped += $f } }
if ($Files -and $dropped.Count -lt $Files.Count) {
    Say '（有拖進來的檔案路徑無法辨識，改用檔案選取視窗）' 'Yellow'
}
foreach ($f in $dropped) {
    try { Add-Book $f ([ref]$books) ([ref]$cols) } catch { Say ("處理失敗：{0}" -f $_.Exception.Message) 'Red' }
}

if ($dropped.Count -eq 0) {
    while ($true) {
        Line
        Say '  [1] 新增電子書（選檔案 → 自動複製 + 抓封面）' 'White'
        Say '  [2] 前台顯示／隱藏某一本' 'White'
        Say '  [3] 移除電子書' 'White'
        Say '  [4] 類別管理（新增／刪除空類別）' 'White'
        Say '  [5] 結束' 'White'
        $c = Read-Host '請選擇'
        if ($c -eq '1') {
            $picked = Pick-Files
            if ($picked.Count -eq 0) { Say '沒有選擇檔案。' 'DarkGray'; continue }
            foreach ($f in $picked) {
                try { Add-Book $f ([ref]$books) ([ref]$cols) } catch { Say ("處理失敗：{0}" -f $_.Exception.Message) 'Red' }
            }
        } elseif ($c -eq '2') {
            try { Toggle-Book ([ref]$books) } catch { Say ("處理失敗：{0}" -f $_.Exception.Message) 'Red' }
        } elseif ($c -eq '3') {
            try { Remove-Book ([ref]$books) } catch { Say ("處理失敗：{0}" -f $_.Exception.Message) 'Red' }
        } elseif ($c -eq '4') {
            try { Manage-Cols ([ref]$books) ([ref]$cols) } catch { Say ("處理失敗：{0}" -f $_.Exception.Message) 'Red' }
        } elseif ($c -eq '5' -or $c -eq '') {
            break
        }
    }
}

# 有變動才寫檔
$after = Build-Json $books $cols
if ($after -ne $before) {
    Save-Books $books $cols
    Line
    Say '已更新 books.json（index.html 內的備援書單也同步了）。' 'Green'
    Say ("書架上現在有 {0} 本書。" -f $books.Count) 'Green'
    Say ''
    Say '最後一步：把這個資料夾的變更上傳到 GitHub' 'Yellow'
    Say '  books.json、index.html、covers/、以及新增的電子書檔案' 'DarkGray'
} else {
    Line
    Say '沒有任何變更。' 'DarkGray'
}

Say ''
Read-Host '按 Enter 結束' | Out-Null
