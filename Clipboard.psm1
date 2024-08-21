function Set-ClipboardWithImage {
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    $bmp = [System.Drawing.Bitmap]::new($Path)
    [System.Windows.Forms.Clipboard]::SetImage($bmp)
    $bmp.Dispose()
}