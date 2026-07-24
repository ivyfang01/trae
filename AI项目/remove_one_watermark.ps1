Add-Type -AssemblyName System.Drawing

$in = "C:\Users\HY\.cursor\projects\d\assets\c__Users_HY_AppData_Roaming_Cursor_User_workspaceStorage_4329c1b5e94952ef11948a41effc4ef4_images_G7______G7________2_-7fa47f01-da35-4601-b386-db6bf55720a4.png"
$out = "C:\Users\HY\.cursor\projects\d\assets\avatar-lanyard-no-watermark.png"

$img = [System.Drawing.Image]::FromFile($in)
$bottom = 0.045
$right = 0.10
$newW = [int]($img.Width * (1 - $right))
$newH = [int]($img.Height * (1 - $bottom))

$bmp = New-Object System.Drawing.Bitmap $newW, $newH
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$src = New-Object System.Drawing.Rectangle 0, 0, $newW, $newH
$g.DrawImage($img, 0, 0, $src, [System.Drawing.GraphicsUnit]::Pixel)
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose(); $img.Dispose()

Write-Output "Saved: $out"
