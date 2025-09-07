#===========================
#InspectingAndMovingInputData.ps1
#外接メール送信_メール送信準備
#背景:
#処理内容:
# - ユーザー領域のCSVを読み込み
# - 形式検証＆件数集計
# - 送信内容の確認ダイアログ
# - 問題なければ ProcessingArea へ安全に移送
# ポイント:
# - INIはスクリプトの相対パスで読取
# - CSVは UserArea（絶対パス）から拾う
#起動契機:
#┗ー「UserData\InspectingAndMovingInputData.ps1 - ショートカット」をユーザーが起動
#===========================

param([string]$IniPath = "\\goa-mailsv\001_InspectingAndMovingInputData\mailsettings_本番用.ini")

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

### A-0-1: 時間制限チェック（8-18時のみ実行許可）
$currentHour = (Get-Date).Hour
if ($currentHour -lt 8 -or $currentHour -ge 18) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("今の時間は起動を許可されていません。`r`n`r`n実行可能時間: 8:00-18:00", "時間制限エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    exit 1
}

### A-2: INIファイルを読み込んでハッシュテーブル化する関数
function Get-Ini {
    param([string]$Path)
    if (-not (Test-Path $Path)) { throw "INIが見つかりません: $Path" }
    $iniHashTable = @{}; $currentSection='General'; $iniHashTable[$currentSection]=@{}
    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $trimmedLine=$line.Trim(); if (-not $trimmedLine -or $trimmedLine -match '^#') { continue }
        if ($trimmedLine -match '^\[(.+?)\]') { $currentSection=$matches[1]; if (-not $iniHashTable[$currentSection]){$iniHashTable[$currentSection]=@{}}; continue }
        if ($trimmedLine -match '^(.*?)\s*=\s*(.*)$') { $iniHashTable[$currentSection][$matches[1].Trim()]=$matches[2].Trim() }
    }
    return $iniHashTable
}

### A-0: 初期チェック（多重起動防止エラー通知確認）
$iniData = Get-Ini -Path $IniPath
$pathSettings = $iniData['Paths']
$UserData = $pathSettings['UserData']
$InspectingAndMovingInputDataLog = $pathSettings['001_InspectingAndMovingInputDataLog']

# エラー通知フォルダの存在確認
$errorFolder = Join-Path $UserData "送信処理に問題が発生しています。事務部に連絡してください"
if (Test-Path $errorFolder) {
    $errorMsg = "エラー通知フォルダが存在するため処理を中止"
    Write-Host $errorMsg
    Write-Log $errorMsg 'ERROR'
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("送信処理に問題が発生しています。事務部に連絡してください。", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

# 多重起動チェック用のロックファイル存在確認
$InspectingAndMovingInputData = $pathSettings['001_InspectingAndMovingInputData']
$lockFile = Join-Path $InspectingAndMovingInputData "起動中.txt"
if (Test-Path $lockFile) {
    $warningMsg = "メール送信スクリプトが起動しています。処理中止"
    Write-Host $warningMsg
    Write-Log $warningMsg 'WARN'
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("メール送信スクリプトが起動しています。", "警告", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    exit 1
}

# ログファイル、ロックファイルはスクリプト配置フォルダに作成
Set-Content -Path $lockFile -Value "処理中" -Encoding UTF8
if (-not (Test-Path $InspectingAndMovingInputDataLog)) { New-Item -ItemType Directory -Path $InspectingAndMovingInputDataLog | Out-Null }

# ログ出力用関数（INFO/WARN/ERRORレベルでログを出力）
function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')][string]$Level='INFO')
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    $logLine = "$timestamp [$Level] $Message"
    Write-Host $logLine
    try {
        if ($script:LogFile) { Add-Content -LiteralPath $script:LogFile -Value $logLine -Encoding UTF8 }
    } catch {}
}

### A-1: 肥大化ログファイル削除判定
function Remove-OldLogFiles {
    param([string]$TargetPath)
  
    # 保存期間　日単位で指定
    $Days = 731
  
    # 指定日以前に更新されたログファイル＝削除対象
    $target = Get-ChildItem $TargetPath -Filter "*.log" -ErrorAction SilentlyContinue | Where-Object {$_.CreationTime -lt (Get-Date).AddDays(-1 * $Days) }
  
    If ($null -ne $Error[0]){
        Write-Log "ログファイル削除処理でエラー: $($Error[0])" 'ERROR'
        return $false
    }
  
    $target | ForEach-Object {
        Write-Log "削除対象: $($_.FullName)" 'INFO'
        # 削除
        Remove-Item $_.FullName -Force
    
        If ($? -eq $false){
            Write-Log "ログファイル削除失敗: $($_.FullName)" 'ERROR'
            return $false
        } else {
            Write-Log "ログファイル削除完了: $($_.FullName)" 'INFO'
        }
    }
    return $true
}

### A-3: CSVファイルのヘッダーからカラム名を取得し、動的に変数を設定する関数
function Get-CsvColumns {
    param([string]$FilePath, [hashtable]$Ini, [string]$Section)
  
    # UTF-8 BOM付き判定を厳密に実施（BOM付きUTF-8のみ許可）
    try {
        $bom = Get-Content -LiteralPath $FilePath -Encoding Byte -TotalCount 3 -ErrorAction Stop
        if ($bom.Length -lt 3 -or $bom[0] -ne 0xEF -or $bom[1] -ne 0xBB -or $bom[2] -ne 0xBF) {
            throw "CSVファイルがUTF-8 BOM付きではありません！（UTF-8 BOM付きのみ使用可能）"
        }
    
        # UTF-8エンコーディングの検証（バイト単位でのチェック）
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        $utf8Decoder = [System.Text.UTF8Encoding]::new($true, $true)  # BOMあり、例外発生あり
        try {
            $null = $utf8Decoder.GetString($bytes)  # エンコーディング検証のみ
        } catch {
            throw "CSVファイルがUTF-8 BOM付きではありません！"
        }
    } catch {
        $errorMsg = "CSVファイルのエンコーディング判定エラー: $($_.Exception.Message)"
        Write-Host $errorMsg
        Write-Log $errorMsg 'ERROR'
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("CSVファイルがUTF-8 BOM付きではありません！`r`n`r`nファイル: $FilePath", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        throw "CSVファイルがUTF-8 BOM付きではありません！"
    }
  
    # CSVファイルのヘッダー行を取得
    $headerLine = Get-Content -LiteralPath $FilePath -Encoding UTF8 -TotalCount 1
    $headers = $headerLine -split ','
  
    # 動的にカラム名を取得（固定カラム名を前提とした自動検出）
    $columnMappings = @{}
    $invalidColumns = @()  # 書式が間違っているカラムを記録
  
    foreach ($header in $headers) {
        $trimmedHeader = $header.Trim('"').Trim()
    
        # メール送信日カラムの判定（完全一致）
        if ($trimmedHeader -eq 'メール送信日') {
            $columnMappings['SentDate'] = $trimmedHeader
            continue
        }
    
        # メールアドレスカラムの判定（完全一致）
        if ($trimmedHeader -eq 'メールアドレス') {
            $columnMappings['Email'] = $trimmedHeader
            continue
        }
    
        # メール件名カラムの判定（新旧両形式対応）
        if ($trimmedHeader -eq 'メール件名') {
            # 番号なしの場合は1として扱う（現在の要件）
            $columnMappings["Title1"] = $trimmedHeader
            continue
        } elseif ($trimmedHeader -like 'メール件名*') {
            # 番号ありの場合（将来の要件対応）
            $columnMappings[$trimmedHeader] = $trimmedHeader
            continue
        }
    
        # メール本文カラムの判定（新旧両形式対応）
        if ($trimmedHeader -eq 'メール本文') {
            # 番号なしの場合は1として扱う（現在の要件）
            $columnMappings["Body1"] = $trimmedHeader
            continue
        } elseif ($trimmedHeader -like 'メール本文*') {
            # 番号ありの場合（将来の要件対応）
            $columnMappings[$trimmedHeader] = $trimmedHeader
            continue
        }
    
        # 想定外の「件名」「本文」関連カラムをチェック（表記の揺れを検出）
        if ($trimmedHeader -match '件名|本文') {
            # 正しい形式ではないカラム名を記録
            $invalidColumns += $trimmedHeader
        }
    }
  
    # 表記の揺れがあった場合はエラー
    if ($invalidColumns.Count -gt 0) {
        $errorMessage = "CSVファイルの書式が間違っています。`r`n`r`n正しくないカラム名: " + ($invalidColumns -join ", ") + "`r`n`r`n正しい形式: 'メール件名', 'メール件名2', 'メール本文', 'メール本文2' など（1番目は番号なし、2番目以降は半角数字）"
        Write-Host "CSVカラム書式エラー: $errorMessage"
        Write-Log "CSVカラム書式エラー: $errorMessage" 'ERROR'
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "CSVファイル書式エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        throw "CSVファイルの書式が間違っています"
    }
  
    # 必須カラムの確認
    # このスクリプトでは以下の2つのカラムのみが必須です：
    # メール送信日|メールアドレス
    if (-not $columnMappings.ContainsKey('SentDate')) {
        $errorMessage = "必須カラム 'メール送信日' がありません"
        Write-Host $errorMessage
        Write-Log $errorMessage 'ERROR'
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "CSVファイル書式エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        throw $errorMessage
    }
  
    if (-not $columnMappings.ContainsKey('Email')) {
        $errorMessage = "必須カラム 'メールアドレス' がありません"
        Write-Host $errorMessage
        Write-Log $errorMessage 'ERROR'
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "CSVファイル書式エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        throw $errorMessage
    }

    return @{
        Headers = $headers
        Mappings = $columnMappings
    }
}

### A-3-1: ヘッダー付きでCSVを読み込む関数（カラム名は動的に取得）
function Read-Csv-WithDynamicHeader {
    param([string]$FilePath, [array]$Headers)
    # ヘッダーを自動認識してCSVを読み込み（1行目はヘッダーとして扱う）
    Import-Csv -LiteralPath $FilePath -Encoding UTF8
}

### A-4: 初期化（パスログ）
### 各種パス変数の設定とログファイル準備
$BulkSendingExternalEmails = $pathSettings['002_BulkSendingExternalEmails']
$script:LogFile = Join-Path $InspectingAndMovingInputDataLog ("A_" + (Get-Date -Format 'yyyy_MM_dd') + ".log")

Write-Log "Start ScriptA. UserData=$UserData, InspectingAndMovingInputData=$InspectingAndMovingInputData"

### A-5: ユーザー領域のCSV列挙
### ユーザー領域(UserData)からCSVファイルを列挙
$csvFiles = Get-ChildItem -LiteralPath $UserData -Filter '*.csv' -File -ErrorAction SilentlyContinue | Sort-Object Name

### A-5-1: CSVファイル名の形式チェック
### CsvKeyword_yyyy_mm_dd.csv形式かどうかをチェック
$validKeywords = @()
foreach ($sectionName in $iniData.Keys) {
    if ($sectionName -in 'General','SMTP','Paths') { continue }
    if ($iniData[$sectionName].ContainsKey('CsvKeyword')) {
        $validKeywords += $iniData[$sectionName]['CsvKeyword']
    }
}


# ファイル名形式チェックと有効ファイル抽出を同時実行
$validCsvFiles = @()
$invalidFileNames = @()
foreach ($csvFile in $csvFiles) {
    $fileName = $csvFile.Name
    $isValidFormat = $false
    
    foreach ($keyword in $validKeywords) {
        # CsvKeyword_yyyy_mm_dd hh_mm_ss.csv の形式をチェック
        $pattern = "^" + [regex]::Escape($keyword) + "_\d{4}_\d{2}_\d{2} \d{2}_\d{2}_\d{2}\.csv$"
        if ($fileName -match $pattern) {
            $isValidFormat = $true
            break
        }
    }
    
    if ($isValidFormat) {
        $validCsvFiles += $csvFile
    } else {
        $invalidFileNames += $fileName
    }
}

# ファイル名形式エラーがある場合は処理を中止（エラー通知フォルダは作成しない）
if ($invalidFileNames.Count -gt 0) {
    $errorMessage = "ファイル名の形式が間違っています。`r`n`r`n対象ファイル: " + ($invalidFileNames -join ", ") + "`r`n`r`n正しい形式: [CsvKeyword]_yyyy_mm_dd hh_mm_ss.csv"
    Write-Log "ファイル名形式エラー: $errorMessage" 'ERROR'
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($errorMessage, 'ファイル名形式エラー', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    
    # ロックファイルを削除して終了
    if (Test-Path $lockFile) { Remove-Item $lockFile -Force }
    exit 1
}

# 有効なCSVファイルが存在しない場合は処理終了
if (-not $validCsvFiles) { 
    Write-Log "有効な形式のCSVファイルなし。処理終了" 
    if (Test-Path $lockFile) { Remove-Item $lockFile -Force }
    exit 0 
}

# 以降の処理では$validCsvFilesを使用
$csvFiles = $validCsvFiles

### A-6: 形式検証＆件数集計（Toが空／送信済の検出）
### 各CSVの行ごとに「宛先(To)が空」「未送信」をチェックし、警告集計
$totalRowCount=0; $readyToSendCount=0; $warningMessages=@(); $hasValidationError = $false; $missingColumnsError = $false
foreach($csvFile in $csvFiles){
    # CSVファイルのカラム情報を取得
    try {
        $columnInfo = Get-CsvColumns -FilePath $csvFile.FullName -Ini $iniData -Section $null
    } catch {
        # 必須カラムエラーかどうかを判定
        if ($_.Exception.Message -match "必須カラム") {
            $missingColumnsError = $true
            $warningMessages += "$($csvFile.Name): $($_.Exception.Message)"
        } else {
            # その他のエラー（エンコーディングエラーなど）
            Write-Log "CSVファイル処理エラー: $($_.Exception.Message)" 'ERROR'
            $warningMessages += "$($csvFile.Name): $($_.Exception.Message)"
            $hasValidationError = $true
        }
        continue
    }
  
    $csvRows = @(Read-Csv-WithDynamicHeader -FilePath $csvFile.FullName -Headers $columnInfo.Headers)
    $totalRowCount += $csvRows.Count
    $rowIndex=0
    foreach($csvRow in $csvRows){ 
        $rowIndex++
        $emailColumn = $columnInfo.Mappings['Email']
        $sentDateColumn = $columnInfo.Mappings['SentDate']
        if (-not $csvRow.$emailColumn){ $warningMessages += "$($csvFile.Name) ${rowIndex}行目: 宛先無し" }
        if (-not $csvRow.$sentDateColumn){ $readyToSendCount++ }
    }
}

# 必須カラムエラーがあった場合は処理を中止（エラー通知フォルダは作成しない）
if ($missingColumnsError) {
    $errorMessage = "CSVファイルに必須カラムがありません。`r`n`r`n詳細: " + (($warningMessages | Where-Object { $_ -match "必須カラム" }) -join "`r`n")
    Write-Log "必須カラムエラー: $errorMessage" 'ERROR'
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($errorMessage, 'CSVファイル必須カラムエラー', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    # ロックファイルを削除して終了
    if (Test-Path $lockFile) { Remove-Item $lockFile -Force }
    exit 1
}

# CSVファイル検証でエラーがあった場合は処理を中止（エラー通知フォルダは作成しない）
if ($hasValidationError) {
    Write-Log "CSVファイルの検証でエラーが発生しました。処理を中止します。" 'ERROR'
    Add-Type -AssemblyName System.Windows.Forms
    $errorMessage = "CSVファイルの書式にエラーがあります。`r`n`r`n詳細はログファイルを確認してください。`r`n`r`n警告: $($warningMessages.Count) 件"
    if ($warningMessages.Count -gt 0){ $errorMessage += "`r`n----エラー詳細----`r`n" + (($warningMessages | Select-Object -First 5) -join "`r`n") }
    [System.Windows.Forms.MessageBox]::Show($errorMessage, 'CSVファイル検証エラー', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
  
    # ロックファイルを削除して終了
    if (Test-Path $lockFile) { Remove-Item $lockFile -Force }
    exit 1
}

### A-7: 送信前の確認ポップアップ（誤送信防止）
### 送信前に確認ダイアログを表示し、ユーザーに最終確認
$needConfirm = ($iniData['General']['Confirm'] -eq 'true')
if ($needConfirm) {
    Add-Type -AssemblyName System.Windows.Forms
    $csvFileCount = if ($csvFiles) { @($csvFiles).Count } else { 0 }
    $summaryMessage = "OKを押下したら処理が行われます。`r`n`r`n対象ファイル: $csvFileCount 件`r`n未送信行(送信候補): $readyToSendCount`r`n警告: $($warningMessages.Count) 件"
    if ($warningMessages.Count -gt 0){ $summaryMessage += "`r`n----警告抜粋----`r`n" + (($warningMessages | Select-Object -First 5) -join "`r`n") }
    $userResponse = [System.Windows.Forms.MessageBox]::Show($summaryMessage, '送信準備 確認', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($userResponse -ne 'OK'){
        Write-Log "ユーザー取消"
        if (Test-Path $lockFile) { Remove-Item $lockFile -Force }
        exit 0
    }
}

### A-8: バックアップCSVファイル作成とInspectingAndMovingInputDataへの移動
### CSVファイルのバックアップを作成してから、準備領域へ移送
function Create-BackupAndMove-CsvFiles {
    param([array]$CsvFiles, [hashtable]$IniData)
    
    $CSV_FileBackup = $IniData['Paths']['CSV_FileBackup']
    $InspectingAndMovingInputData = $IniData['Paths']['001_InspectingAndMovingInputData']
    
    # バックアップディレクトリが存在しない場合は作成
    if (-not (Test-Path $CSV_FileBackup)) {
        New-Item -ItemType Directory -Path $CSV_FileBackup -Force | Out-Null
        Write-Log "バックアップディレクトリを作成: $CSV_FileBackup"
    }
    
    foreach($csvFile in $CsvFiles){
        # 既存のバックアップファイルを削除（最新一世代のみ保持）
        $backupFileName = $csvFile.Name
        $backupPath = Join-Path $CSV_FileBackup $backupFileName
        if (Test-Path $backupPath) {
            Remove-Item -LiteralPath $backupPath -Force
            Write-Log "既存バックアップファイル削除: $backupPath"
        }
        
        # 新しいバックアップを作成（UTF-8 BOM付きを維持）
        $content = Get-Content -LiteralPath $csvFile.FullName -Encoding UTF8 -Raw
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($backupPath, $content, $utf8WithBom)
        Write-Log "バックアップファイル作成: $backupPath"
        
        # InspectingAndMovingInputDataへ移動（送信準備完了の命名規則で）
        $timeStamp = Get-Date -Format 'yyyy-MM-dd-HH-mm'
        $destinationFileName = "【送信準備完了_$timeStamp】_$($csvFile.Name)"
        $destinationPath = Join-Path $InspectingAndMovingInputData $destinationFileName
        Write-Log "Move: $($csvFile.FullName) -> $destinationPath"
        
        # ファイル内容を1行ずつ読み込んで確実に構造を保持
        $lines = Get-Content -LiteralPath $csvFile.FullName -Encoding UTF8
        $moveContent = $lines -join "`r`n"
        
        # UTF-8 BOM付きエンコーディングで出力
        [System.IO.File]::WriteAllText($destinationPath, $moveContent, $utf8WithBom)
        Remove-Item -LiteralPath $csvFile.FullName -Force
    }
}

### A-8: 送信処理領域へCSVファイルを移動
### 各CSVファイルを準備領域へ送信準備完了の命名規則で移動（UTF-8 BOM付きを維持）
foreach($csvFile in $csvFiles){
    $timeStamp = Get-Date -Format 'yyyy-MM-dd-HH-mm'
    $destinationFileName = "【送信準備完了_$timeStamp】_$($csvFile.Name)"
    $destinationPath = Join-Path $InspectingAndMovingInputData $destinationFileName
    Write-Log "Move: $($csvFile.FullName) -> $destinationPath"
    
    # ファイル内容を1行ずつ読み込んで確実に構造を保持
    $lines = Get-Content -LiteralPath $csvFile.FullName -Encoding UTF8
    $content = $lines -join "`r`n"
    
    # UTF-8 BOM付きエンコーディングで出力
    $utf8WithBom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($destinationPath, $content, $utf8WithBom)
    Remove-Item -LiteralPath $csvFile.FullName -Force
}

### A-8: バックアップ作成と準備領域への移送
Create-BackupAndMove-CsvFiles -CsvFiles $csvFiles -IniData $iniData
Write-Log "移送完了。スクリプトB（JP1）が拾います"
Write-Log "ScriptA完了"