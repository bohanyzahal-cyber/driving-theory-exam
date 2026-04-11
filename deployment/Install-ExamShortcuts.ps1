<#
.SYNOPSIS
    מתקין קיצורי דרך למערכת מבחן התאוריה הדיגיטלית של רישוי צה"ל.

.DESCRIPTION
    הסקריפט מוריד את האייקונים מ-GitHub, ממיר אותם ל-ICO, ויוצר קיצורי דרך
    בשולחן העבודה שפותחים את הדפים בדפדפן במצב אפליקציה (חלון נקי).

    Teacher → יוצר 4 קיצורים: בוחן, מורה, נבחן, תרגול
    Student → יוצר 2 קיצורים: נבחן, תרגול

.PARAMETER Role
    Teacher או Student - קובע אילו קיצורים ייווצרו

.PARAMETER BaseUrl
    כתובת הבסיס של האתר. ברירת מחדל: GitHub Pages של הפרויקט.

.EXAMPLE
    .\Install-ExamShortcuts.ps1 -Role Teacher
    .\Install-ExamShortcuts.ps1 -Role Student
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Teacher', 'Student')]
    [string]$Role,

    [string]$BaseUrl = 'https://bohanyzahal-cyber.github.io/driving-theory-exam'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ============================================================================
# הגדרת הקיצורים לכל תפקיד
# ============================================================================

$allShortcuts = @{
    Examiner = @{ Name = 'בוחן';  Page = 'examiner.html'; Icon = 'icon-examiner-192.png' }
    Teacher  = @{ Name = 'מורה';  Page = 'teacher.html';  Icon = 'icon-teacher-192.png'  }
    Examinee = @{ Name = 'נבחן';  Page = 'examinee.html'; Icon = 'icon-examinee-192.png' }
    Student  = @{ Name = 'תרגול'; Page = 'student.html';  Icon = 'icon-192.png'          }
}

$roleShortcuts = @{
    Teacher = @('Examiner', 'Teacher', 'Examinee', 'Student')
    Student = @('Examinee', 'Student')
}

# ============================================================================
# פונקציות עזר
# ============================================================================

function Test-IsAdmin {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-InstallPaths {
    $isAdmin = Test-IsAdmin
    if ($isAdmin) {
        return @{
            IconDir  = 'C:\ProgramData\ExamShortcuts\icons'
            Desktop  = [Environment]::GetFolderPath('CommonDesktopDirectory')
            Scope    = 'כל המשתמשים (Public Desktop)'
        }
    } else {
        return @{
            IconDir  = Join-Path $env:LOCALAPPDATA 'ExamShortcuts\icons'
            Desktop  = [Environment]::GetFolderPath('Desktop')
            Scope    = 'משתמש נוכחי בלבד'
        }
    }
}

function Find-Browser {
    $paths = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe"
    )
    foreach ($p in $paths) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    return $null
}

function Convert-PngToIco {
    param(
        [Parameter(Mandatory = $true)][string]$PngPath,
        [Parameter(Mandatory = $true)][string]$IcoPath
    )

    # ICO פורמט תומך ב-PNG ישירות החל מ-Windows Vista.
    # המבנה: ICONDIR (6B) + ICONDIRENTRY (16B) + PNG data
    $pngBytes = [System.IO.File]::ReadAllBytes($PngPath)

    # קריאת מימדים מתוך ה-IHDR של ה-PNG (big-endian)
    $wBytes = $pngBytes[16..19]; [Array]::Reverse($wBytes)
    $hBytes = $pngBytes[20..23]; [Array]::Reverse($hBytes)
    $width  = [System.BitConverter]::ToUInt32($wBytes, 0)
    $height = [System.BitConverter]::ToUInt32($hBytes, 0)

    # בשדה הגודל של ICO יש רק בייט אחד; 0 מייצג 256
    $icoW = if ($width  -ge 256) { [byte]0 } else { [byte]$width  }
    $icoH = if ($height -ge 256) { [byte]0 } else { [byte]$height }

    $stream = [System.IO.File]::Create($IcoPath)
    try {
        $bw = New-Object System.IO.BinaryWriter($stream)
        # ICONDIR
        $bw.Write([uint16]0)              # reserved
        $bw.Write([uint16]1)              # type: 1 = icon
        $bw.Write([uint16]1)              # image count
        # ICONDIRENTRY
        $bw.Write([byte]$icoW)            # width
        $bw.Write([byte]$icoH)            # height
        $bw.Write([byte]0)                # palette
        $bw.Write([byte]0)                # reserved
        $bw.Write([uint16]1)              # color planes
        $bw.Write([uint16]32)             # bits per pixel
        $bw.Write([uint32]$pngBytes.Length)  # image data size
        $bw.Write([uint32]22)             # image data offset (6 + 16)
        # Image data (PNG raw)
        $bw.Write($pngBytes)
        $bw.Flush()
    }
    finally {
        $stream.Dispose()
    }
}

function Get-RemoteIcon {
    param(
        [Parameter(Mandatory = $true)][string]$IconFile,
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$DestDir
    )

    if (-not (Test-Path $DestDir)) {
        New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    }

    $pngPath = Join-Path $DestDir $IconFile
    $icoPath = Join-Path $DestDir ([System.IO.Path]::ChangeExtension($IconFile, '.ico'))

    # תמיד מורידים מחדש כדי לקבל גרסה עדכנית
    $url = "$BaseUrl/$IconFile"
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $pngPath -UseBasicParsing -TimeoutSec 30
    }
    catch {
        throw "נכשלה הורדת האייקון מ-$url`n$($_.Exception.Message)"
    }

    Convert-PngToIco -PngPath $pngPath -IcoPath $icoPath
    return $icoPath
}

function New-ExamShortcut {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$IconPath,
        [Parameter(Mandatory = $true)][string]$DesktopPath,
        [string]$BrowserPath
    )

    $finalPath = Join-Path $DesktopPath "$Name.lnk"

    # WScript.Shell ב-PowerShell 5.1 משתמש ב-ANSI API ולכן לא תומך בשמות קבצים בעברית.
    # הפתרון: יוצרים את הקיצור בנתיב זמני באנגלית, שומרים אותו, ואז מעבירים
    # אותו לשם הסופי בעברית באמצעות .NET (System.IO.File.Move) שמטפל ב-Unicode כמו שצריך.
    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("ExamShortcut_" + [Guid]::NewGuid().ToString('N') + ".lnk")

    # מחיקת קיצור קודם אם קיים (ב-.NET כדי לתמוך בשם עברי)
    if ([System.IO.File]::Exists($finalPath)) {
        [System.IO.File]::Delete($finalPath)
    }
    if ([System.IO.File]::Exists($tempPath)) {
        [System.IO.File]::Delete($tempPath)
    }

    $wsh = New-Object -ComObject WScript.Shell
    try {
        $sc = $wsh.CreateShortcut($tempPath)

        if ($BrowserPath) {
            # מצב אפליקציה: חלון נקי ללא סרגל כתובות/כרטיסיות
            $sc.TargetPath       = $BrowserPath
            $sc.Arguments        = "--app=$Url"
            $sc.WorkingDirectory = Split-Path $BrowserPath -Parent
        }
        else {
            # Fallback: פותח בדפדפן ברירת המחדל
            $sc.TargetPath = $Url
        }

        $sc.IconLocation = "$IconPath,0"
        # תיאור באנגלית בלבד - WScript.Shell לא מטפל נכון ב-Unicode בשדה הזה
        $sc.Description  = "Driving Theory Exam System"
        $sc.Save()
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh) | Out-Null
    }

    # העברה לשם הסופי העברי באמצעות .NET
    [System.IO.File]::Move($tempPath, $finalPath)
}

# ============================================================================
# תהליך ראשי
# ============================================================================

Write-Host ''
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host '   התקנת קיצורי דרך - מערכת תאוריה דיגיטלית' -ForegroundColor Cyan
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host ''

$paths = Get-InstallPaths
Write-Host "תפקיד:       $Role"
Write-Host "היקף התקנה:  $($paths.Scope)"
Write-Host "שולחן עבודה: $($paths.Desktop)"
Write-Host "תיקיית אייקונים: $($paths.IconDir)"
Write-Host ''

$browser = Find-Browser
if ($browser) {
    Write-Host "דפדפן זוהה: $browser" -ForegroundColor Green
} else {
    Write-Host 'לא זוהה Edge/Chrome - ייעשה שימוש בדפדפן ברירת המחדל' -ForegroundColor Yellow
}
Write-Host ''

$shortcutKeys = $roleShortcuts[$Role]
$created = @()
$failed  = @()

foreach ($key in $shortcutKeys) {
    $cfg  = $allShortcuts[$key]
    $name = $cfg.Name
    $url  = "$BaseUrl/$($cfg.Page)"

    Write-Host "מעבד: $name ($($cfg.Page))..." -NoNewline

    try {
        $icoPath = Get-RemoteIcon -IconFile $cfg.Icon -BaseUrl $BaseUrl -DestDir $paths.IconDir
        New-ExamShortcut -Name $name -Url $url -IconPath $icoPath -DesktopPath $paths.Desktop -BrowserPath $browser
        Write-Host ' ✓' -ForegroundColor Green
        $created += $name
    }
    catch {
        Write-Host ' ✗' -ForegroundColor Red
        Write-Host "   שגיאה: $($_.Exception.Message)" -ForegroundColor Red
        $failed += $name
    }
}

Write-Host ''
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host "נוצרו בהצלחה: $($created.Count) קיצורים" -ForegroundColor Green
if ($created.Count -gt 0) {
    Write-Host "  [$($created -join ', ')]" -ForegroundColor Green
}
if ($failed.Count -gt 0) {
    Write-Host "נכשלו:        $($failed.Count) קיצורים" -ForegroundColor Red
    Write-Host "  [$($failed -join ', ')]" -ForegroundColor Red
}
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host ''

if ($failed.Count -gt 0) {
    exit 1
}
exit 0
