Add-Type -AssemblyName System.Drawing

function Save-CroppedNoWatermark {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [double]$BottomCropRatio = 0.055,
        [double]$RightCropRatio = 0.0
    )

    $img = [System.Drawing.Image]::FromFile($InputPath)
    $newW = [int]($img.Width * (1 - $RightCropRatio))
    $newH = [int]($img.Height * (1 - $BottomCropRatio))

    $bmp = New-Object System.Drawing.Bitmap $newW, $newH
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

    $src = New-Object System.Drawing.Rectangle 0, 0, $newW, $newH
    $g.DrawImage($img, 0, 0, $src, [System.Drawing.GraphicsUnit]::Pixel)

    $bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose(); $img.Dispose()
}

$base = "C:\Users\HY\.cursor\projects\d\assets"
$items = @(
    @{
        In = "$base\c__Users_HY_AppData_Roaming_Cursor_User_workspaceStorage_4329c1b5e94952ef11948a41effc4ef4_images_______AI___2_-cc4d4b23-a55a-4e57-866a-9f3792a03acf.png"
        Out = "$base\avatar-anime-no-watermark.png"
        Bottom = 0.06
        Right = 0.0
    },
    @{
        In = "$base\c__Users_HY_AppData_Roaming_Cursor_User_workspaceStorage_4329c1b5e94952ef11948a41effc4ef4_images_______AI___3_-b15a97f5-b91c-422a-848a-95de85f79baa.png"
        Out = "$base\avatar-photo-no-watermark.png"
        Bottom = 0.04
        Right = 0.05
    }
)

foreach ($item in $items) {
    Save-CroppedNoWatermark -InputPath $item.In -OutputPath $item.Out -BottomCropRatio $item.Bottom -RightCropRatio $item.Right
    Write-Output "Saved: $($item.Out)"
}
