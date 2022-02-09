Set-StrictMode -version Latest

Add-Type -assemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = 'JSON Files|*.json|All Files|*.*'
$dialog.InitialDirectory = '.'

[void] $dialog.ShowDialog()

$dialog.FileName
