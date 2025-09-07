#===========================
#BulkSendingExternalEmail.psm1
#外接メール送信_共通関数モジュール
#機能:
# - INI設定読み込み
# - CSV解析・メール送信機能
# - エラーハンドリング共通処理
#===========================

### 共通: INIファイル読み込み関数
function Get-IniData {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) { 
        throw "INI設定ファイルが見つかりません: $Path" 
    }
    
    $iniHashTable = @{}
    $currentSection = 'General'
    $iniHashTable[$currentSection] = @{}
    
    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $trimmedLine = $line.Trim()
        
        # コメント行や空行をスキップ
        if (-not $trimmedLine -or $trimmedLine -match '^#') { 
            continue 
        }
        
        # セクション名の処理
        if ($trimmedLine -match '^\[(.+?)\]') {
            $currentSection = $matches[1]
            if (-not $iniHashTable[$currentSection]) {
                $iniHashTable[$currentSection] = @{}
            }
            continue
        }
        
        # キー=値のペアの処理
        if ($trimmedLine -match '^(.*?)\s*=\s*(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $iniHashTable[$currentSection][$key] = $value
        }
    }
    
    return $iniHashTable
}

### 共通: CSV解析関数
function Get-CsvColumnInfo {
    param([string]$FilePath)
    
    # UTF-8 BOM付き判定
    try {
        $bom = Get-Content -LiteralPath $FilePath -Encoding Byte -TotalCount 3 -ErrorAction Stop
        if ($bom.Length -lt 3 -or $bom[0] -ne 0xEF -or $bom[1] -ne 0xBB -or $bom[2] -ne 0xBF) {
            throw "CSVファイルがUTF-8 BOM付きではありません"
        }
        
        # UTF-8エンコーディングの検証
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        $utf8Decoder = [System.Text.UTF8Encoding]::new($true, $true)
        $null = $utf8Decoder.GetString($bytes)
        
    } catch {
        throw "CSVファイルのエンコーディングエラー: $($_.Exception.Message)"
    }
    
    # ヘッダー行を取得
    $headerLine = Get-Content -LiteralPath $FilePath -Encoding UTF8 -TotalCount 1
    $headers = $headerLine -split ','
    
    # カラムマッピングを作成
    $columnMappings = @{}
    
    foreach ($header in $headers) {
        $trimmedHeader = $header.Trim('"').Trim()
        
        # 標準カラムの判定
        switch ($trimmedHeader) {
            'メール送信日' { $columnMappings['SentDate'] = $trimmedHeader }
            'メールアドレス' { $columnMappings['Email'] = $trimmedHeader }
            'メール件名' { $columnMappings['Subject'] = $trimmedHeader }
            'メール本文' { $columnMappings['Body'] = $trimmedHeader }
            default {
                # 番号付きカラムの判定
                if ($trimmedHeader -match '^メール件名(\d+)$') {
                    $columnMappings["Subject$($matches[1])"] = $trimmedHeader
                } elseif ($trimmedHeader -match '^メール本文(\d+)$') {
                    $columnMappings["Body$($matches[1])"] = $trimmedHeader
                }
            }
        }
    }
    
    # 必須カラムチェック
    if (-not $columnMappings.ContainsKey('SentDate')) {
        throw "必須カラム 'メール送信日' がありません"
    }
    if (-not $columnMappings.ContainsKey('Email')) {
        throw "必須カラム 'メールアドレス' がありません"
    }
    
    return @{
        Headers = $headers
        Mappings = $columnMappings
    }
}

### 共通: メール送信関数
function Send-SingleEmail {
    param(
        [string]$To,
        [string]$From,
        [string]$Subject,
        [string]$Body,
        [string]$SmtpServer,
        [int]$SmtpPort,
        [string]$Bcc = $null
    )
    
    try {
        # メールアドレスの基本検証
        if (-not $To -or $To -notmatch '^[^@]+@[^@]+\.[^@]+$') {
            throw "無効なメールアドレス: $To"
        }
        
        # SMTPクライアントの設定
        $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $smtpClient.EnableSsl = $false  # 内部SMTPサーバーのため
        
        # メールメッセージの作成
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $From
        $mailMessage.To.Add($To)
        
        if ($Bcc) {
            $mailMessage.Bcc.Add($Bcc)
        }
        
        $mailMessage.Subject = $Subject
        $mailMessage.Body = $Body
        $mailMessage.IsBodyHtml = $false
        $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
        
        # メール送信
        $smtpClient.Send($mailMessage)
        
        # リソースの解放
        $mailMessage.Dispose()
        $smtpClient.Dispose()
        
        return @{
            Success = $true
            ErrorMessage = $null
        }
        
    } catch {
        return @{
            Success = $false
            ErrorMessage = $_.Exception.Message
        }
    }
}

### 共通: CSVファイル一括メール送信関数
function Send-CsvEmails {
    param(
        [System.IO.FileInfo]$CsvFile,
        [hashtable]$IniData
    )
    
    try {
        $smtpSettings = $IniData['SMTP']
        $generalSettings = $IniData['General']
        
        # CSV列情報を取得
        $columnInfo = Get-CsvColumnInfo -FilePath $CsvFile.FullName
        
        # CSVデータを読み込み
        $csvData = Import-Csv -LiteralPath $CsvFile.FullName -Encoding UTF8
        
        $totalRows = $csvData.Count
        $sentCount = 0
        $skippedCount = 0
        $errorCount = 0
        $errors = @()
        
        Write-Host "CSV処理開始: $($CsvFile.Name), 総行数: $totalRows"
        
        foreach ($row in $csvData) {
            $rowIndex = $csvData.IndexOf($row) + 1
            
            try {
                # 必須フィールドの取得
                $emailAddress = $row.($columnInfo.Mappings['Email'])
                $sentDate = $row.($columnInfo.Mappings['SentDate'])
                
                # 送信済みチェック（送信日が入力されている場合はスキップ）
                if ($sentDate) {
                    $skippedCount++
                    Write-Host "行 $rowIndex : 送信済みのためスキップ ($emailAddress)"
                    continue
                }
                
                # メールアドレスの存在チェック
                if (-not $emailAddress) {
                    $skippedCount++
                    Write-Host "行 $rowIndex : メールアドレスが空のためスキップ"
                    continue
                }
                
                # 件名と本文の取得（デフォルト値設定）
                $subject = if ($columnInfo.Mappings.ContainsKey('Subject')) { 
                    $row.($columnInfo.Mappings['Subject']) 
                } else { 
                    "重要なお知らせ" 
                }
                
                $body = if ($columnInfo.Mappings.ContainsKey('Body')) { 
                    $row.($columnInfo.Mappings['Body']) 
                } else { 
                    "こちらは自動送信メールです。" 
                }
                
                # メール送信実行
                $result = Send-SingleEmail -To $emailAddress -From $generalSettings['From'] -Subject $subject -Body $body -SmtpServer $smtpSettings['Server'] -SmtpPort $smtpSettings['Port'] -Bcc $generalSettings['Bcc']
                
                if ($result.Success) {
                    $sentCount++
                    # 送信日を更新（CSVファイルを更新）
                    $row.($columnInfo.Mappings['SentDate']) = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    Write-Host "行 $rowIndex : 送信成功 ($emailAddress)"
                } else {
                    $errorCount++
                    $errorMessage = "行 $rowIndex : 送信失敗 ($emailAddress) - $($result.ErrorMessage)"
                    $errors += $errorMessage
                    Write-Host $errorMessage
                }
                
            } catch {
                $errorCount++
                $errorMessage = "行 $rowIndex : 処理エラー - $($_.Exception.Message)"
                $errors += $errorMessage
                Write-Host $errorMessage
            }
        }
        
        # 更新されたCSVを保存（送信日が更新されたもの）
        if ($sentCount -gt 0) {
            $csvData | Export-Csv -LiteralPath $CsvFile.FullName -Encoding UTF8 -NoTypeInformation -Force
            Write-Host "CSV更新完了: 送信日フィールドを更新"
        }
        
        $summary = "処理完了 - 送信: $sentCount, スキップ: $skippedCount, エラー: $errorCount, 総行数: $totalRows"
        Write-Host $summary
        
        # 結果の判定
        if ($errorCount -eq 0) {
            return @{
                Success = $true
                ErrorMessage = $null
                Summary = $summary
                SentCount = $sentCount
                SkippedCount = $skippedCount
                ErrorCount = $errorCount
            }
        } else {
            return @{
                Success = $false
                ErrorMessage = "送信エラー $errorCount 件: " + ($errors -join "; ")
                Summary = $summary
                SentCount = $sentCount
                SkippedCount = $skippedCount
                ErrorCount = $errorCount
            }
        }
        
    } catch {
        return @{
            Success = $false
            ErrorMessage = "CSV処理中に致命的エラー: $($_.Exception.Message)"
            Summary = "処理失敗"
            SentCount = 0
            SkippedCount = 0
            ErrorCount = 1
        }
    }
}

### 共通: 文字列検証関数
function Test-EmailAddress {
    param([string]$EmailAddress)
    
    if (-not $EmailAddress) {
        return $false
    }
    
    # 基本的なメールアドレス形式チェック
    return $EmailAddress -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

### 共通: ファイル操作関数
function Copy-FileWithEncoding {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [System.Text.Encoding]$Encoding = [System.Text.UTF8Encoding]::new($true)
    )
    
    try {
        $content = Get-Content -LiteralPath $SourcePath -Encoding UTF8 -Raw
        [System.IO.File]::WriteAllText($DestinationPath, $content, $Encoding)
        return $true
    } catch {
        Write-Error "ファイルコピーエラー: $($_.Exception.Message)"
        return $false
    }
}

# モジュールの関数をエクスポート
Export-ModuleMember -Function Get-IniData, Get-CsvColumnInfo, Send-SingleEmail, Send-CsvEmails, Test-EmailAddress, Copy-FileWithEncoding