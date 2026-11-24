param([Parameter(Mandatory=$true)][string]$FilePath)
$ErrorActionPreference = 'Stop'
$word = $docs = $doc = $null
$temp = $null
$failed = $false
try {
    $source = Get-Item -LiteralPath $FilePath
    if ($source.PSIsContainer -or $source.Extension -notin @('.doc', '.docx')) { throw 'Wybierz plik DOC lub DOCX.' }
    $target = [IO.Path]::ChangeExtension($source.FullName, '.pdf')
    if (Test-Path -LiteralPath $target) { throw 'Plik PDF już istnieje.' }
    $temp = Join-Path $source.DirectoryName ([guid]::NewGuid().ToString() + '.pdf')
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $word.Options.UpdateLinksAtOpen = $false
    $docs = $word.Documents
    $doc = $docs.Open($source.FullName, $false, $true, $false)
    $doc.ExportAsFixedFormat($temp, 17)
    [IO.File]::Move($temp, $target)
} catch {
    $failed = $true
    Add-Type -AssemblyName System.Windows.Forms
    [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Błąd konwersji PDF')
} finally {
    if ($doc) { try { $doc.Close(0) } catch {} ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
    if ($docs) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($docs) }
    if ($word) { try { $word.Quit(0) } catch {} ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    if ($temp -and (Test-Path -LiteralPath $temp)) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
}
if ($failed) { exit 1 }
