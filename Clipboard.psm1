Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Set-ClipboardWithImage {
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    $bmp = [System.Drawing.Bitmap]::new($Path)
    [System.Windows.Forms.Clipboard]::SetImage($bmp)
    $bmp.Dispose()
}