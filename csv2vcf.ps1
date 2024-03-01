Add-Type -AssemblyName System.Windows.Forms

# CSVファイルの読み込み
function Read-Csv($filepath) {
    $csv = Import-Csv -Path $filepath -Delimiter ',' -Encoding Default
    $contacts = @()
    foreach ($row in $csv) {
        $contact = @{
            "N" = $row.'氏名'
            "ADR" = $row.'住所'
            "TEL;TYPE=HOME" = $row.'固定電話番号'
            "TEL;TYPE=CELL" = $row.'携帯電話番号'
        }
        if ($contact["TEL;TYPE=HOME"] -ne "" -or $contact["TEL;TYPE=CELL"] -ne "") {
            $contacts += $contact
        }
    }
    return $contacts
}

# VCFファイルの書き込み
function Write-Vcf($filepath, $contacts) {
    $vcfContent = @()
    foreach ($contact in $contacts) {
        $vcfContent += "BEGIN:VCARD"
        $vcfContent += "VERSION:3.0"
        $vcfContent += "N:$($contact['N'])"
        $vcfContent += "FN:$($contact['N'])"
        $vcfContent += "item1.ADR;TYPE=HOME;TYPE=pref:;;;`"$($contact['ADR'])`";;;;日本"
        $vcfContent += "item1.X-ABADR:jp"
        if ($contact["TEL;TYPE=HOME"]) {
            $vcfContent += "TEL;TYPE=HOME;TYPE=VOICE:$($contact['TEL;TYPE=HOME'])"
        }
        if ($contact["TEL;TYPE=CELL"]) {
            $vcfContent += "TEL;TYPE=CELL;TYPE=pref;TYPE=VOICE:$($contact['TEL;TYPE=CELL'])"
        }
        $vcfContent += "PRODID:-//Apple Inc.//iCloud Web Address Book 2406B22//EN"
        $vcfContent += "REV:$((Get-Date).ToUniversalTime().ToString('s'))Z"
        $vcfContent += "END:VCARD"
    }
    Set-Content -Path $filepath -Value $vcfContent
}

# メイン処理
$form = New-Object System.Windows.Forms.Form
$form.Text = "CSV to VCF Converter"
$form.Width = 400
$form.Height = 400

# タイトル
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "CSVファイルからvCardファイルを作成"
$titleLabel.Location = New-Object System.Drawing.Point(5, 5)
$titleLabel.Font = New-Object System.Drawing.Font("UD デジタル 教科書体 NP-R", 14) 
$titleLabel.Width = 380
$titleLabel.Height = 40
$form.Controls.Add($titleLabel)
 
# ファイル選択ボタン
$selectFileButton = New-Object System.Windows.Forms.Button
$selectFileButton.Text = "CSVファイルを選択"
$selectFileButton.Location = New-Object System.Drawing.Point(100, 50)
$selectFileButton.Font = New-Object System.Drawing.Font("UD デジタル 教科書体 NP-R", 11)
$selectFileButton.Height = 45
$selectFileButton.Width = 100
$selectFileButton.Add_Click({
    $openFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $openFileDialog.Filter = "CSV (*.csv)|*.csv"
    $openFileDialog.InitialDirectory = "."
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filepath = $openFileDialog.FileName
        $outputFileTextBox.Text = $filepath
    }
})
$form.Controls.Add($selectFileButton)

# 変換ボタン
$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Text = "VCFファイルに変換"
$convertButton.Location = New-Object System.Drawing.Point(100, 100)
$convertButton.Font = New-Object System.Drawing.Font("UD デジタル 教科書体 NP-R", 11)
$convertButton.Height = 45
$convertButton.Width = 100
$convertButton.Add_Click({
    $inputFilePath = $outputFileTextBox.Text
    $outputFilePath = $inputFilePath.Replace(".csv", ".vcf")
    $contacts = Read-Csv $inputFilePath
    Write-Vcf $outputFilePath $contacts

    # メッセージ表示
    [System.Windows.Forms.MessageBox]::Show("VCFファイルへの変換が完了しました。", "完了", [System.Windows.Forms.MessageBoxButtons]::OK)
})
$form.Controls.Add($convertButton)

# 出力ファイル名ラベル
$outputFileLabel = New-Object System.Windows.Forms.Label
$outputFileLabel.Text = "出力ファイル名:"
$outputFileLabel.Location = New-Object System.Drawing.Point(100, 150)
$outputFileLabel.Font = New-Object System.Drawing.Font("UD デジタル 教科書体 NP-R", 11)
$outputFileLabel.Width = 150
$form.Controls.Add($outputFileLabel)

# 出力ファイル名エントリ
$outputFileTextBox = New-Object System.Windows.Forms.TextBox
$outputFileTextBox.Location = New-Object System.Drawing.Point(100, 170)
$outputFileTextBox.Font = New-Object System.Drawing.Font("UD デジタル 教科書体 NP-R", 11)
$outputFileTextBox.Width = 200
$form.Controls.Add($outputFileTextBox)

# 出力ファイル名ラベル
$memoLabel = New-Object System.Windows.Forms.Label
$memoLabel.Text = "CSVファイルの書式:`n氏名,住所,固定電話番号,携帯電話番号`n`n電話番号は最低1つ必要。1つも無ければvcfには出力されない。`n"
$memoLabel.Location = New-Object System.Drawing.Point(25, 210)
$memoLabel.Font = New-Object System.Drawing.Font("UD デジタル 教科書体 NP-R", 11)
$memoLabel.Width = 350
$memoLabel.Height = 150
$memoLabel.AutoSize = $false
$form.Controls.Add($memoLabel)


# アプリの起動
$form.ShowDialog()
