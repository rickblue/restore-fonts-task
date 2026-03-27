# restore-fonts.ps1
# 浣跨敤 SystemParametersInfo API 瀹炴椂搴旂敤绯荤粺瀛椾綋璁剧疆锛堟棤闇€娉ㄩ攢锛?
$script:LogPath = Join-Path -Path $PSScriptRoot -ChildPath "restore-fonts.log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message
    Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
}

function Exit-WithError {
    param(
        [string]$Message,
        [int]$Code = 1
    )

    Write-Host $Message -ForegroundColor Red
    Write-Log -Level "ERROR" -Message $Message
    exit $Code
}

Write-Log -Message "==== Execution started. PID=$PID User=$env:USERNAME Computer=$env:COMPUTERNAME ===="

try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
public struct LOGFONT {
    public int lfHeight;
    public int lfWidth;
    public int lfEscapement;
    public int lfOrientation;
    public int lfWeight;
    public byte lfItalic;
    public byte lfUnderline;
    public byte lfStrikeOut;
    public byte lfCharSet;
    public byte lfOutPrecision;
    public byte lfClipPrecision;
    public byte lfQuality;
    public byte lfPitchAndFamily;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
    public string lfFaceName;
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
public struct NONCLIENTMETRICS {
    public int cbSize;
    public int iBorderWidth;
    public int iScrollWidth;
    public int iScrollHeight;
    public int iCaptionWidth;
    public int iCaptionHeight;
    public LOGFONT lfCaptionFont;
    public int iSmCaptionWidth;
    public int iSmCaptionHeight;
    public LOGFONT lfSmCaptionFont;
    public int iMenuWidth;
    public int iMenuHeight;
    public LOGFONT lfMenuFont;
    public LOGFONT lfStatusFont;
    public LOGFONT lfMessageFont;
    public int iPaddedBorderWidth;
}

public class FontHelper {
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SystemParametersInfo(
        uint uiAction, uint uiParam, ref NONCLIENTMETRICS pvParam, uint fWinIni);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SystemParametersInfo(
        uint uiAction, uint uiParam, ref LOGFONT pvParam, uint fWinIni);

    public const uint SPI_GETNONCLIENTMETRICS = 0x0029;
    public const uint SPI_SETNONCLIENTMETRICS = 0x002A;
    public const uint SPI_GETICONTITLELOGFONT = 0x001F;
    public const uint SPI_SETICONTITLELOGFONT = 0x0022;
    public const uint SPIF_UPDATEINIFILE = 0x0001;
    public const uint SPIF_SENDCHANGE = 0x0002;
}
"@

    Write-Log -Message "Add-Type completed."

    function Set-LogFont {
        param(
            $Font,
            [string]$FaceName,
            [int]$Height,
            [int]$Weight = 400,
            [byte]$Quality = 5  # CLEARTYPE_QUALITY
        )
        $Font.lfHeight = $Height
        $Font.lfWeight = $Weight
        $Font.lfQuality = $Quality
        $Font.lfCharSet = 1  # DEFAULT_CHARSET
        $Font.lfFaceName = $FaceName
        return $Font
    }

    Write-Log -Message "Reading current NONCLIENTMETRICS."
    $metrics = New-Object NONCLIENTMETRICS
    $metrics.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf([type][NONCLIENTMETRICS])
    $ok = [FontHelper]::SystemParametersInfo(
        [FontHelper]::SPI_GETNONCLIENTMETRICS,
        [uint32]$metrics.cbSize,
        [ref]$metrics,
        0
    )
    if (-not $ok) {
        $lastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Exit-WithError -Message "Failed to get current metrics. Win32Error=$lastError"
    }

    Write-Log -Message "Preparing target font settings. Caption and small caption fonts will be left unchanged to avoid changing title bar height."
    $metrics.lfMenuFont = Set-LogFont -Font $metrics.lfMenuFont -FaceName "MiSans Medium" -Height -15
    $metrics.lfStatusFont = Set-LogFont -Font $metrics.lfStatusFont -FaceName "MiSans Medium" -Height -15
    $metrics.lfMessageFont = Set-LogFont -Font $metrics.lfMessageFont -FaceName "MiSans Medium" -Height -15

    Write-Log -Message "Applying NONCLIENTMETRICS. Caption=unchanged, SmallCaption=unchanged, Menu=MiSans Medium/-15, Status=MiSans Medium/-15, Message=MiSans Medium/-15."
    $ok = [FontHelper]::SystemParametersInfo(
        [FontHelper]::SPI_SETNONCLIENTMETRICS,
        [uint32]$metrics.cbSize,
        [ref]$metrics,
        ([FontHelper]::SPIF_UPDATEINIFILE -bor [FontHelper]::SPIF_SENDCHANGE)
    )
    if (-not $ok) {
        $lastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Exit-WithError -Message "Failed to set non-client metrics. Win32Error=$lastError"
    }

    Write-Log -Message "Applying icon title font. Face=MiSans Medium Height=-15 Weight=400."
    $iconFont = New-Object LOGFONT
    $iconFont.lfHeight = -15
    $iconFont.lfWidth = 0
    $iconFont.lfWeight = 400
    $iconFont.lfCharSet = 1
    $iconFont.lfQuality = 5
    $iconFont.lfFaceName = "MiSans Medium"

    $ok = [FontHelper]::SystemParametersInfo(
        [FontHelper]::SPI_SETICONTITLELOGFONT,
        [uint32][System.Runtime.InteropServices.Marshal]::SizeOf([type][LOGFONT]),
        [ref]$iconFont,
        ([FontHelper]::SPIF_UPDATEINIFILE -bor [FontHelper]::SPIF_SENDCHANGE)
    )
    if (-not $ok) {
        $lastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Exit-WithError -Message "Failed to set icon font. Win32Error=$lastError"
    }

    Write-Log -Message "Font settings applied successfully."
    Write-Host "Font settings applied successfully." -ForegroundColor Green
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Log -Level "ERROR" -Message "Unhandled exception: $errorMessage"
    throw
}
finally {
    Write-Log -Message "==== Execution finished. ===="
}
