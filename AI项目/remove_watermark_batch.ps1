Add-Type -AssemblyName System.Drawing

function Save-CroppedNoWatermark {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [double]$BottomCropRatio = 0.045,
        [double]$RightCropRatio = 0.10
    )

    if (-not (Test-Path $InputPath)) {
        Write-Output "Skip missing: $InputPath"
        return
    }

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
    Write-Output "Saved: $OutputPath ($newW x $newH)"
}

$items = @(
    @{
        In = "C:\Users\HY\.cursor\projects\d\assets\c__Users_HY_AppData_Roaming_Cursor_User_workspaceStorage_4329c1b5e94952ef11948a41effc4ef4_images_G7______G7________1_-67e8b9dd-cb3d-419f-9733-54d145ae63d3.png"
        Out = "C:\Users\HY\.cursor\projects\d\assets\avatar-lanyard-v2-no-watermark.png"
        Bottom = 0.05
        Right = 0.12
    },
    @{
        In = "C:\Users\HY\.cursor\projects\d\assets\c__Users_HY_AppData_Roaming_Cursor_User_workspaceStorage_4329c1b5e94952ef11948a41effc4ef4_images_G7______G7________2_-7fa47f01-da35-4601-b386-db6bf55720a4.png"
        Out = "C:\Users\HY\.cursor\projects\d\assets\avatar-lanyard-v1-no-watermark.png"
        Bottom = 0.05
        Right = 0.12
    }
)

foreach ($item in $items) {
    Save-CroppedNoWatermark -InputPath $item.In -OutputPath $item.Out -BottomCropRatio $item.Bottom -RightCropRatio $item.Right
}
