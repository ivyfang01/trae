Add-Type -AssemblyName System.Drawing

function Remove-Watermark {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    $img = [System.Drawing.Image]::FromFile($InputPath)
    $bmp = New-Object System.Drawing.Bitmap $img.Width, $img.Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.DrawImage($img, 0, 0, $img.Width, $img.Height)

    $w = $bmp.Width
    $h = $bmp.Height
    $x0 = [int]($w * 0.66)
    $y0 = [int]($h * 0.88)
    $rw = $w - $x0
    $rh = $h - $y0

    $samples = New-Object System.Collections.Generic.List[System.Drawing.Color]
    for ($y = $y0 - 35; $y -lt $y0; $y++) {
        if ($y -lt 0) { continue }
        for ($x = $x0; $x -lt $w; $x++) {
            $samples.Add($bmp.GetPixel($x, $y)) | Out-Null
        }
    }
    if ($samples.Count -eq 0) {
        for ($x = $x0 - 80; $x -lt $x0; $x++) {
            if ($x -lt 0) { continue }
            $samples.Add($bmp.GetPixel($x, $h - 20)) | Out-Null
        }
    }

    $r = 0; $gC = 0; $b = 0
    foreach ($c in $samples) {
        $r += $c.R; $gC += $c.G; $b += $c.B
    }
    $n = [Math]::Max(1, $samples.Count)
    $fill = [System.Drawing.Color]::FromArgb([int]($r / $n), [int]($gC / $n), [int]($b / $n))
    $brush = New-Object System.Drawing.SolidBrush $fill

    for ($y = $y0; $y -lt $h; $y++) {
        for ($x = $x0; $x -lt $w; $x++) {
            $p = $bmp.GetPixel($x, $y)
            $lum = 0.299 * $p.R + 0.587 * $p.G + 0.114 * $p.B
            if ($lum -gt 175) {
                $srcY = [Math]::Max(0, $y - 45)
                $bmp.SetPixel($x, $y, $bmp.GetPixel($x, $srcY))
            }
        }
    }

    $g.FillRectangle($brush, $x0, $y0, $rw, [int]($rh * 0.35))

    $bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose(); $img.Dispose()
}

$base = "C:\Users\HY\.cursor\projects\d\assets"
$files = @(
    @{
        In  = "$base\c__Users_HY_AppData_Roaming_Cursor_User_workspaceStorage_4329c1b5e94952ef11948a41effc4ef4_images_______AI___2_-cc4d4b23-a55a-4e57-866a-9f3792a03acf.png"
        Out = "$base\avatar-anime-no-watermark.png"
    },
    @{
        In  = "$base\c__Users_HY_AppData_Roaming_Cursor_User_workspaceStorage_4329c1b5e94952ef11948a41effc4ef4_images_______AI___3_-b15a97f5-b91c-422a-848a-95de85f79baa.png"
        Out = "$base\avatar-photo-no-watermark.png"
    }
)

foreach ($f in $files) {
    Remove-Watermark -InputPath $f.In -OutputPath $f.Out
    Write-Output "Saved: $($f.Out)"
}
