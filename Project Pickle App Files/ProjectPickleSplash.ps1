param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ReadyFile,

    [int]$SafetyTimeoutSeconds = 180
)

$ErrorActionPreference =
    "Stop"

Add-Type `
    -AssemblyName System.Windows.Forms

Add-Type `
    -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# Read the application version directly from ProjectPickle.ps1 so the splash
# screen does not need a separately maintained hard-coded version number.
# ---------------------------------------------------------------------------
function Get-ProjectPickleSplashVersion {
    $appScriptPath =
        Join-Path `
            $PSScriptRoot `
            "ProjectPickle.ps1"

    if (
        -not (
            Test-Path `
                -LiteralPath $appScriptPath
        )
    ) {
        return ""
    }

    try {
        $openingText =
            [string]::Join(
                "`n",
                [string[]]@(
                    Get-Content `
                        -LiteralPath $appScriptPath `
                        -TotalCount 120
                )
            )

        $versionMatch =
            [regex]::Match(
                $openingText,
                '\$script:AppVersion\s*=\s*"([^"]+)"'
            )

        if ($versionMatch.Success) {
            return (
                [string]$versionMatch.Groups[1].Value
            ).Trim()
        }
    }
    catch {
    }

    return ""
}

# ---------------------------------------------------------------------------
# Load the best available pickle artwork.
#
# Preferred:
#   ProjectPickleSplash.png
#
# Fallbacks:
#   ProjectPickle.png
#   ProjectPickle.ico
#
# When no artwork exists, Project Pickle draws a simple pickle shape.
# ---------------------------------------------------------------------------
function New-ProjectPickleBuiltInBitmap {
    $bitmap =
        [System.Drawing.Bitmap]::new(
            220,
            220,
            [System.Drawing.Imaging.PixelFormat]::Format32bppArgb
        )

    $graphics =
        [System.Drawing.Graphics]::FromImage(
            $bitmap
        )

    $bodyPath =
        [System.Drawing.Drawing2D.GraphicsPath]::new()

    $shadowBrush =
        [System.Drawing.SolidBrush]::new(
            [System.Drawing.Color]::FromArgb(
                48,
                20,
                55,
                35
            )
        )

    $bodyBrush =
        [System.Drawing.SolidBrush]::new(
            [System.Drawing.Color]::FromArgb(
                92,
                164,
                84
            )
        )

    $outlinePen =
        [System.Drawing.Pen]::new(
            [System.Drawing.Color]::FromArgb(
                37,
                92,
                48
            ),
            [single]5
        )

    $highlightBrush =
        [System.Drawing.SolidBrush]::new(
            [System.Drawing.Color]::FromArgb(
                128,
                205,
                110
            )
        )

    $bumpBrush =
        [System.Drawing.SolidBrush]::new(
            [System.Drawing.Color]::FromArgb(
                66,
                132,
                62
            )
        )

    $eyeBrush =
        [System.Drawing.SolidBrush]::new(
            [System.Drawing.Color]::FromArgb(
                28,
                50,
                33
            )
        )

    $eyeHighlightBrush =
        [System.Drawing.SolidBrush]::new(
            [System.Drawing.Color]::White
        )

    $smilePen =
        [System.Drawing.Pen]::new(
            [System.Drawing.Color]::FromArgb(
                28,
                50,
                33
            ),
            [single]4
        )

    try {
        $graphics.SmoothingMode =
            [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

        $graphics.InterpolationMode =
            [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

        $graphics.PixelOffsetMode =
            [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

        $graphics.Clear(
            [System.Drawing.Color]::Transparent
        )

        # Soft shadow under the pickle.
        $graphics.FillEllipse(
            $shadowBrush,
            48,
            188,
            124,
            18
        )

        # Rounded vertical pickle body.
        $bodyPath.AddArc(
            55,
            18,
            110,
            80,
            180,
            180
        )

        $bodyPath.AddLine(
            165,
            58,
            165,
            162
        )

        $bodyPath.AddArc(
            55,
            122,
            110,
            80,
            0,
            180
        )

        $bodyPath.AddLine(
            55,
            162,
            55,
            58
        )

        $bodyPath.CloseFigure()

        $graphics.FillPath(
            $bodyBrush,
            $bodyPath
        )

        $graphics.DrawPath(
            $outlinePen,
            $bodyPath
        )

        # Left-side highlight.
        $graphics.FillEllipse(
            $highlightBrush,
            72,
            38,
            24,
            118
        )

        # Pickle bumps.
        $pickleBumps =
            @(
                @(105, 42, 16, 10),
                @(132, 61, 14, 11),
                @(91, 73, 13, 10),
                @(119, 118, 15, 11),
                @(82, 142, 14, 10),
                @(137, 151, 13, 9),
                @(70, 105, 12, 9)
            )

        foreach ($bump in $pickleBumps) {
            $graphics.FillEllipse(
                $bumpBrush,
                [int]$bump[0],
                [int]$bump[1],
                [int]$bump[2],
                [int]$bump[3]
            )
        }

        # Eyes.
        $graphics.FillEllipse(
            $eyeBrush,
            88,
            91,
            13,
            17
        )

        $graphics.FillEllipse(
            $eyeBrush,
            121,
            91,
            13,
            17
        )

        $graphics.FillEllipse(
            $eyeHighlightBrush,
            92,
            94,
            4,
            5
        )

        $graphics.FillEllipse(
            $eyeHighlightBrush,
            125,
            94,
            4,
            5
        )

        # Smile.
        $graphics.DrawArc(
            $smilePen,
            91,
            102,
            42,
            31,
            12,
            156
        )
    }
    finally {
        $smilePen.Dispose()
        $eyeHighlightBrush.Dispose()
        $eyeBrush.Dispose()
        $bumpBrush.Dispose()
        $highlightBrush.Dispose()
        $outlinePen.Dispose()
        $bodyBrush.Dispose()
        $shadowBrush.Dispose()
        $bodyPath.Dispose()
        $graphics.Dispose()
    }

    return $bitmap
}

function Get-ProjectPickleSplashSourceBitmap {
    # A real PNG is the safest format for resizing and rotation.
    #
    # ProjectPickle.ico is intentionally not used here. Some ICO files
    # contain PNG-compressed frames that Windows PowerShell 5.1 / GDI+
    # can decode as colored static when converted to a large Bitmap.
    $pngPath =
        Join-Path `
            $PSScriptRoot `
            "ProjectPickleSplash.png"

    if (
        Test-Path `
            -LiteralPath $pngPath
    ) {
        $stream =
            $null

        $loadedImage =
            $null

        try {
            $stream =
                [System.IO.File]::Open(
                    $pngPath,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite
                )

            $loadedImage =
                [System.Drawing.Image]::FromStream(
                    $stream,
                    $true,
                    $true
                )

            if (
                $loadedImage.Width -le 0 -or
                $loadedImage.Height -le 0
            ) {
                throw (
                    "ProjectPickleSplash.png did not contain " +
                    "valid image dimensions."
                )
            }

            # Clone the image so the PNG file and stream can be released.
            return (
                [System.Drawing.Bitmap]::new(
                    $loadedImage
                )
            )
        }
        catch {
            # A bad optional PNG must never stop Project Pickle startup.
            # Fall through to the built-in pickle.
        }
        finally {
            if ($null -ne $loadedImage) {
                $loadedImage.Dispose()
            }

            if ($null -ne $stream) {
                $stream.Dispose()
            }
        }
    }

    return (
        New-ProjectPickleBuiltInBitmap
    )
}

# ---------------------------------------------------------------------------
# Create one rotated animation frame.
#
# Frames are prepared once at startup. The timer only switches between the
# prepared images, which keeps the animation smooth and avoids repeated image
# allocation while Project Pickle is loading.
# ---------------------------------------------------------------------------
function New-ProjectPickleWiggleFrame {
    param(
        [Parameter(Mandatory = $true)]
        [System.Drawing.Image]$SourceImage,

        [single]$Angle
    )

    $canvasSize =
        196

    $targetSize =
        152

    $frame =
        [System.Drawing.Bitmap]::new(
            $canvasSize,
            $canvasSize,
            [System.Drawing.Imaging.PixelFormat]::Format32bppArgb
        )

    $graphics =
        [System.Drawing.Graphics]::FromImage(
            $frame
        )

    try {
        $graphics.Clear(
            [System.Drawing.Color]::Transparent
        )

        $graphics.SmoothingMode =
            [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

        $graphics.InterpolationMode =
            [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

        $graphics.PixelOffsetMode =
            [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

        $graphics.CompositingQuality =
            [System.Drawing.Drawing2D.CompositingQuality]::HighQuality

        $widthScale =
            $targetSize /
            [double]$SourceImage.Width

        $heightScale =
            $targetSize /
            [double]$SourceImage.Height

        $scale =
            [System.Math]::Min(
                $widthScale,
                $heightScale
            )

        $drawWidth =
            [single](
                $SourceImage.Width *
                $scale
            )

        $drawHeight =
            [single](
                $SourceImage.Height *
                $scale
            )

        $graphics.TranslateTransform(
            [single]($canvasSize / 2),
            [single]($canvasSize / 2)
        )

        $graphics.RotateTransform(
            $Angle
        )

        $graphics.DrawImage(
            $SourceImage,
            [single](-$drawWidth / 2),
            [single](-$drawHeight / 2),
            $drawWidth,
            $drawHeight
        )

        $graphics.ResetTransform()
    }
    finally {
        $graphics.Dispose()
    }

    return $frame
}

$script:ProjectPickleSplashReadyFile =
    $ReadyFile

$script:ProjectPickleSplashClosing =
    $false

$script:ProjectPickleSplashCleanupComplete =
    $false

$script:ProjectPickleSplashFrameIndex =
    0

$script:ProjectPickleSplashDotIndex =
    0

$script:ProjectPickleSplashSourceBitmap =
    Get-ProjectPickleSplashSourceBitmap

$script:ProjectPickleSplashFrames =
    [System.Collections.Generic.List[System.Drawing.Image]]::new()

# Rock from center to the right, back through center, then to the left.
$wiggleAngles =
    [single[]]@(
        0,
        3,
        6,
        8,
        6,
        3,
        0,
        -3,
        -6,
        -8,
        -6,
        -3
    )

foreach ($wiggleAngle in $wiggleAngles) {
    [void]$script:ProjectPickleSplashFrames.Add(
        (
            New-ProjectPickleWiggleFrame `
                -SourceImage $script:ProjectPickleSplashSourceBitmap `
                -Angle $wiggleAngle
        )
    )
}

$surfaceColor =
    [System.Drawing.Color]::FromArgb(
        250,
        252,
        248
    )

$accentColor =
    [System.Drawing.Color]::FromArgb(
        47,
        122,
        87
    )

$textColor =
    [System.Drawing.Color]::FromArgb(
        31,
        41,
        51
    )

$mutedTextColor =
    [System.Drawing.Color]::FromArgb(
        82,
        97,
        107
    )

$form =
    [System.Windows.Forms.Form]::new()

$form.Text =
    "Project Pickle"

$form.FormBorderStyle =
    [System.Windows.Forms.FormBorderStyle]::None

$form.StartPosition =
    [System.Windows.Forms.FormStartPosition]::CenterScreen

$form.ClientSize =
    [System.Drawing.Size]::new(
        430,
        342
    )

$form.BackColor =
    $accentColor

$form.ShowInTaskbar =
    $false

$form.TopMost =
    $true

$form.KeyPreview =
    $true

$contentPanel =
    [System.Windows.Forms.Panel]::new()

$contentPanel.Dock =
    [System.Windows.Forms.DockStyle]::Fill

$contentPanel.Padding =
    [System.Windows.Forms.Padding]::new(
        2
    )

$contentPanel.BackColor =
    $accentColor

$form.Controls.Add(
    $contentPanel
)

$innerPanel =
    [System.Windows.Forms.Panel]::new()

$innerPanel.Dock =
    [System.Windows.Forms.DockStyle]::Fill

$innerPanel.BackColor =
    $surfaceColor

$contentPanel.Controls.Add(
    $innerPanel
)

$pictureBox =
    [System.Windows.Forms.PictureBox]::new()

$pictureBox.Location =
    [System.Drawing.Point]::new(
        117,
        18
    )

$pictureBox.Size =
    [System.Drawing.Size]::new(
        196,
        196
    )

$pictureBox.SizeMode =
    [System.Windows.Forms.PictureBoxSizeMode]::CenterImage

$pictureBox.BackColor =
    $surfaceColor

$pictureBox.Image =
    $script:ProjectPickleSplashFrames[0]

$innerPanel.Controls.Add(
    $pictureBox
)

$script:ProjectPickleSplashPictureBox =
    $pictureBox

$titleLabel =
    [System.Windows.Forms.Label]::new()

$titleLabel.Text =
    "Project Pickle"

$titleLabel.Location =
    [System.Drawing.Point]::new(
        20,
        216
    )

$titleLabel.Size =
    [System.Drawing.Size]::new(
        390,
        34
    )

$titleLabel.TextAlign =
    [System.Drawing.ContentAlignment]::MiddleCenter

$titleLabel.Font =
    [System.Drawing.Font]::new(
        "Segoe UI",
        [single]18,
        [System.Drawing.FontStyle]::Bold
    )

$titleLabel.ForeColor =
    $accentColor

$titleLabel.BackColor =
    $surfaceColor

$innerPanel.Controls.Add(
    $titleLabel
)

$loadingLabel =
    [System.Windows.Forms.Label]::new()

$loadingLabel.Text =
    "Loading Project Pickle"

$loadingLabel.Location =
    [System.Drawing.Point]::new(
        20,
        254
    )

$loadingLabel.Size =
    [System.Drawing.Size]::new(
        390,
        28
    )

$loadingLabel.TextAlign =
    [System.Drawing.ContentAlignment]::MiddleCenter

$loadingLabel.Font =
    [System.Drawing.Font]::new(
        "Segoe UI",
        [single]11,
        [System.Drawing.FontStyle]::Regular
    )

$loadingLabel.ForeColor =
    $textColor

$loadingLabel.BackColor =
    $surfaceColor

$innerPanel.Controls.Add(
    $loadingLabel
)

$script:ProjectPickleSplashLoadingLabel =
    $loadingLabel

$version =
    Get-ProjectPickleSplashVersion

$versionLabel =
    [System.Windows.Forms.Label]::new()

$versionLabel.Text =
    if (
        [string]::IsNullOrWhiteSpace(
            $version
        )
    ) {
        "Preparing the application..."
    }
    else {
        "Version $version"
    }

$versionLabel.Location =
    [System.Drawing.Point]::new(
        20,
        288
    )

$versionLabel.Size =
    [System.Drawing.Size]::new(
        390,
        24
    )

$versionLabel.TextAlign =
    [System.Drawing.ContentAlignment]::MiddleCenter

$versionLabel.Font =
    [System.Drawing.Font]::new(
        "Segoe UI",
        [single]9,
        [System.Drawing.FontStyle]::Regular
    )

$versionLabel.ForeColor =
    $mutedTextColor

$versionLabel.BackColor =
    $surfaceColor

$innerPanel.Controls.Add(
    $versionLabel
)

$script:ProjectPickleSplashStopwatch =
    [System.Diagnostics.Stopwatch]::StartNew()

$script:ProjectPickleSplashTimer =
    [System.Windows.Forms.Timer]::new()

# About 12 animation frames per second.
$script:ProjectPickleSplashTimer.Interval =
    85

function Clear-ProjectPickleSplashResources {
    if ($script:ProjectPickleSplashCleanupComplete) {
        return
    }

    $script:ProjectPickleSplashCleanupComplete =
        $true

    if (
        $null -ne
        $script:ProjectPickleSplashTimer
    ) {
        try {
            $script:ProjectPickleSplashTimer.Stop()
        }
        catch {
        }

        try {
            $script:ProjectPickleSplashTimer.Dispose()
        }
        catch {
        }
    }

    if (
        $null -ne
        $script:ProjectPickleSplashPictureBox
    ) {
        try {
            $script:ProjectPickleSplashPictureBox.Image =
                $null
        }
        catch {
        }
    }

    if (
        $null -ne
        $script:ProjectPickleSplashFrames
    ) {
        foreach (
            $frame in
            $script:ProjectPickleSplashFrames
        ) {
            if ($null -ne $frame) {
                try {
                    $frame.Dispose()
                }
                catch {
                }
            }
        }

        $script:ProjectPickleSplashFrames.Clear()
    }

    if (
        $null -ne
        $script:ProjectPickleSplashSourceBitmap
    ) {
        try {
            $script:ProjectPickleSplashSourceBitmap.Dispose()
        }
        catch {
        }
    }

    if (
        $null -ne
        $script:ProjectPickleSplashStopwatch
    ) {
        try {
            $script:ProjectPickleSplashStopwatch.Stop()
        }
        catch {
        }
    }
}

$script:ProjectPickleSplashTimer.Add_Tick({
    try {
        if ($script:ProjectPickleSplashClosing) {
            return
        }

        # The main ProjectPickle.ps1 deletes this marker as soon as its
        # primary window is ready. That is the signal to close this splash.
        if (
            -not (
                [System.IO.File]::Exists(
                    $script:ProjectPickleSplashReadyFile
                )
            )
        ) {
            $script:ProjectPickleSplashClosing =
                $true

            $script:ProjectPickleSplashTimer.Stop()

            $form.Close()

            return
        }

        $script:ProjectPickleSplashFrameIndex =
            (
                $script:ProjectPickleSplashFrameIndex +
                1
            ) %
            $script:ProjectPickleSplashFrames.Count

        $script:ProjectPickleSplashPictureBox.Image =
            $script:ProjectPickleSplashFrames[
                $script:ProjectPickleSplashFrameIndex
            ]

        # Change the loading dots less frequently than the wiggle frames.
        if (
            (
                $script:ProjectPickleSplashFrameIndex %
                3
            ) -eq 0
        ) {
            $script:ProjectPickleSplashDotIndex =
                (
                    $script:ProjectPickleSplashDotIndex +
                    1
                ) %
                4

            $dots =
                "." *
                $script:ProjectPickleSplashDotIndex

            $script:ProjectPickleSplashLoadingLabel.Text =
                "Loading Project Pickle$dots"
        }

        if (
            $SafetyTimeoutSeconds -gt 0 -and
            $script:ProjectPickleSplashStopwatch.Elapsed.TotalSeconds -ge
            $SafetyTimeoutSeconds
        ) {
            # Do not leave a stuck splash on screen forever. The app process
            # is not terminated; only the splash marker and splash form close.
            try {
                if (
                    Test-Path `
                        -LiteralPath $script:ProjectPickleSplashReadyFile
                ) {
                    Remove-Item `
                        -LiteralPath $script:ProjectPickleSplashReadyFile `
                        -Force `
                        -ErrorAction SilentlyContinue
                }
            }
            catch {
            }

            $script:ProjectPickleSplashClosing =
                $true

            $script:ProjectPickleSplashTimer.Stop()

            $form.Close()
        }
    }
    catch {
        # A splash animation problem must never prevent the main app from
        # starting. Close only the splash.
        $script:ProjectPickleSplashClosing =
            $true

        try {
            $script:ProjectPickleSplashTimer.Stop()
        }
        catch {
        }

        try {
            $form.Close()
        }
        catch {
        }
    }
}.GetNewClosure())

$form.Add_Shown({
    $script:ProjectPickleSplashTimer.Start()
}.GetNewClosure())

$form.Add_FormClosed({
    Clear-ProjectPickleSplashResources
}.GetNewClosure())

try {
    [System.Windows.Forms.Application]::Run(
        $form
    )
}
finally {
    Clear-ProjectPickleSplashResources

    try {
        $form.Dispose()
    }
    catch {
    }
}