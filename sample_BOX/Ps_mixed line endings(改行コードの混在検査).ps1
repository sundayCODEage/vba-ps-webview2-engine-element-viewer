# =========================================================
# 改行コード（CRLF）チェック ＆ 一括変換スクリプト
# =========================================================

# 1. 起動時の注意メッセージ（警告表示）
Write-Host "==========================================================" -ForegroundColor Red
Write-Host " 同一階層PSモジュールの、" -ForegroundColor Cyan
Write-Host "         改行コード（CRLF）チェック ＆ 一括変換スクリプト" -ForegroundColor Cyan
Write-Host " " -ForegroundColor Yellow
Write-Host "【注意】ファイルを直接更新・上書き保存する処理を含みます。" -ForegroundColor Yellow
Write-Host " 誤動作や意図しない変更を防ぐため、実行前に 必ず対象ファイルを" -ForegroundColor Yellow
Write-Host " 別フォルダ等へバックアップ（コピー）してください。" -ForegroundColor Yellow
Write-Host "==========================================================`n" -ForegroundColor Red

# スクリプトの保存先フォルダ（同一階層）を取得
$targetPath = if ($PSScriptRoot) {$PSScriptRoot } else { Get-Location }
Write-Host "検索対象フォルダ（同一階層のみ）: $targetPath`n" -ForegroundColor Cyan

# 2. 同一階層のファイルを読み込み・混在チェック
$targetFiles = Get-ChildItem -Path $targetPath -File | Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1' }
$mixedFiles = [System.Collections.Generic.List[PSObject]]::new()

foreach ($file in $targetFiles) {
    $filePath = $file.FullName
    $lines = [System.IO.File]::ReadAllText($filePath) -split "(?<=\n)"
    
    $crlfLines = [System.Collections.Generic.List[int]]::new()
    $lfLines   = [System.Collections.Generic.List[int]]::new()
    
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $lineNum = $i + 1
        $lineText = $lines[$i]
        
        if ($lineText.EndsWith("`r`n")) {
            $crlfLines.Add($lineNum)
        } elseif ($lineText.EndsWith("`n")) {
            $lfLines.Add($lineNum)
        }
    }
    
    # CRLF と LF の双方が混在している場合のみ表示・記録
    if ($crlfLines.Count -gt 0 -and$lfLines.Count -gt 0) {
        Write-Host "--------------------------------------------------" -ForegroundColor Yellow
        Write-Host "ファイル : $($file.Name)" -ForegroundColor Green
        Write-Host "パス     : $filePath" -ForegroundColor Gray
        
        if ($crlfLines.Count -ge$lfLines.Count) {
            Write-Host "状態     : 主に CRLF ($($crlfLines.Count)行) ですが、LF が $($lfLines.Count)行 混在しています" -ForegroundColor Yellow
            Write-Host "【LF 改行の行とコード】" -ForegroundColor Red
            foreach ($lineNum in $lfLines) {$code = $lines[$lineNum - 1].TrimEnd("`r", "`n")
                Write-Host ("  行 {0,4}: {1}" -f $lineNum,$code)
            }
        } else {
            Write-Host "状態     : 主に LF ($($lfLines.Count)行) ですが、CRLF が $($crlfLines.Count)行 混在しています" -ForegroundColor Yellow
            Write-Host "【CRLF 改行の行とコード】" -ForegroundColor Red
            foreach ($lineNum in $crlfLines) {$code = $lines[$lineNum - 1].TrimEnd("`r", "`n")
                Write-Host ("  行 {0,4}: {1}" -f $lineNum,$code)
            }
        }
        Write-Host ""
        
        $mixedFiles.Add($file)
    }
}

# 3. 混在ファイルの有無に応じた選択・置換処理
if ($mixedFiles.Count -eq 0) {
    Write-Host "同一階層内に改行コードが混在しているファイルはありませんでした。" -ForegroundColor Green
} else {
    Write-Host "計 $($mixedFiles.Count) 件のファイルで改行コードの混在が検出されました。`n" -ForegroundColor Yellow
    
    # Y/N の入力確認プロンプト
    $response = Read-Host "検出されたファイルをすべて CRLF に置換して上書き保存しますか？ (Y/N)"
    
    if ($response -eq 'Y' -or $response -eq 'y') {
        Write-Host "`n置換処理を実行中..." -ForegroundColor Cyan
        $convertedCount = 0
        
        # PowerShell ISEで確実に動作する UTF-8 (BOMあり) エンコーディングの定義
        $utf8Encoding = New-Object System.Text.UTF8Encoding($true)

        foreach ($file in $mixedFiles) {
            $filePath =$file.FullName
            $text = [System.IO.File]::ReadAllText($filePath)
            
            # 一度すべて LF に揃えてから CRLF に一括置換
            $normalizedText =$text.Replace("`r`n", "`n").Replace("`n", "`r`n")
            
            # UTF-8 (BOMあり) で保存
            [System.IO.File]::WriteAllText($filePath, $normalizedText,$utf8Encoding)
            
            Write-Host "変換・保存完了 : $($file.Name)" -ForegroundColor Green
            $convertedCount++
        }
        
        Write-Host "`n計 $convertedCount 件のファイルを CRLF に置換して保存しました。" -ForegroundColor Cyan
    } else {
        Write-Host "`n処理をキャンセルしました。ファイルは一切変更されていません。" -ForegroundColor Red
    }
}