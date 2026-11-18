<#
.SYNOPSIS
Builds the K-PRICING workbook from your clone: the macro-free template, the
repository's VBA modules, the commit on the Home sheet and the demo sheet.
.DESCRIPTION
Like tools\Sync-SACCR-VBA in K-SACCR, it works from the files of your clone, so
the workbook always comes from the commit and the changes you have checked out.
It needs no GitHub download and runs no regression (tools\Check-KPR.ps1 does
both). Double-click tools\Build-KPR.cmd, or run this script. It:
1. finds your clone: -RepoRoot, the clone this script lives in, or the one
   remembered by an earlier run of this tool or of Check-KPR.ps1;
2. checks "Trust access to the VBA project object model" and, only when you
   answer Y, turns it on for your Windows user;
3. starts its own Excel (your open workbooks and add-ins are not touched) and
   saves src\workbook\KPR_Template.xlsx as a macro-enabled workbook in a
   temporary folder, next to a scratch workbook holding
   tests\modules\KPR_Dev_Import.bas;
4. imports the modules of the clone into it: the public, internal, ui and
   example modules (Production, the default), or also the regression and test
   modules (Development, which can run KPR_Test_RunAll);
5. compiles it (Debug > Compile), writes the commit on the Home sheet and
   builds the KPR Dates Demo sheet with KPR_Demo_BuildDates;
6. saves it as K-PRICING-<commit>.xlsm (K-PRICING-Development-<commit>.xlsm)
   in <clone>\build, which Git ignores; a file of that name is never replaced,
   the new one gets the time in its name;
7. opens it again, once no macro runs, and checks that Excel loads its
   K-PRICING tab and builds the Functions menu (-NoRibbonCheck skips this).
The clone is only read. A clone with uncommitted changes is built as it is and
the Home sheet says so.
.PARAMETER Mode
Production (default) or Development.
.PARAMETER RepoRoot
The clone to build from; remembered for the next runs.
.PARAMETER Open
Opens the finished workbook in Excel (macros need your confirmation).
.PARAMETER NoRibbonCheck
Skips the final Ribbon check (about 30 seconds).
.PARAMETER NoPause
Does not wait for Enter at the end.
.PARAMETER TimeoutMinutes
A watchdog closes this run's Excel after this many minutes, so a dialog nobody
answers cannot hang the build.
.EXAMPLE
tools\Build-KPR.cmd
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Build-KPR.ps1 -Mode Development
.EXAMPLE
&([scriptblock]::Create((Get-Content -Raw 'C:\path\to\K-PRICING\tools\Build-KPR.ps1'))) -RepoRoot 'C:\path\to\K-PRICING'

Inside a PowerShell window that is already open; it starts no other
PowerShell, so it also works where a policy forbids that or where scripts are
disabled.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("Production", "Development")]
    [string]$Mode = "Production",
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot,
    [Parameter(Mandatory = $false)]
    [int]$TimeoutMinutes = 10,
    [Parameter(Mandatory = $false)]
    [switch]$Open,
    [Parameter(Mandatory = $false)]
    [switch]$NoRibbonCheck,
    [Parameter(Mandatory = $false)]
    [switch]$NoPause
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$xlOpenXMLWorkbookMacroEnabled = 52
$compileControlId = 578
$templateCell = "D22"
$script:log = New-Object System.Collections.Generic.List[string]

function Say {
    param([string]$Text, [string]$Color = "Gray")
    $line = "{0}  {1}" -f (Get-Date).ToString("HH:mm:ss"), $Text
    Write-Host $line -ForegroundColor $Color
    $script:log.Add($line)
}

function Get-RegistryValue {
    param([string]$Path, [string]$Name)
    $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return $null
    }
    return $item.$Name
}

function Get-OfficeVersion {
    # "Excel.Application.16" -> "16.0", read without starting Excel.
    $curVer = Get-RegistryValue -Path "Registry::HKEY_CLASSES_ROOT\Excel.Application\CurVer" -Name "(default)"
    if ($curVer -match '\.(\d+)$') {
        return "$($Matches[1]).0"
    }
    return "16.0"
}

function Confirm-VbaProjectAccess {
    param([string]$OfficeVersion)
    $policy = Get-RegistryValue -Path "HKCU:\Software\Policies\Microsoft\Office\$OfficeVersion\Excel\Security" -Name "AccessVBOM"
    if ($null -ne $policy) {
        if ($policy -eq 1) {
            return "On (group policy)"
        }
        throw ("Your organization's policy turns off 'Trust access to the VBA project object model'. " +
               "The check cannot import modules; ask your IT department.")
    }
    $userKey = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Excel\Security"
    if ((Get-RegistryValue -Path $userKey -Name "AccessVBOM") -eq 1) {
        return "On"
    }
    Write-Host ""
    Write-Host "Excel's option 'Trust access to the VBA project object model' is off."
    Write-Host "This tool needs it to import the modules. It is the same switch as"
    Write-Host "File > Options > Trust Center > Trust Center Settings > Macro Settings."
    $answer = Read-Host "Turn it on now for your Windows user? (Y/N)"
    if ($answer -notmatch '^\s*[yYsS]') {
        throw "Stopped: the VBA project access stays off, so no module was imported."
    }
    if (-not (Test-Path -LiteralPath $userKey)) {
        New-Item -Path $userKey -Force | Out-Null
    }
    Set-ItemProperty -LiteralPath $userKey -Name "AccessVBOM" -Value 1 -Type DWord
    return "On (turned on by this run, at the operator's request)"
}

function Get-ExcelProcessId {
    param($Excel)
    if (-not ("KprCheck.Native" -as [type])) {
        Add-Type -Namespace KprCheck -Name Native -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@
    }
    $processId = [uint32]0
    [void][KprCheck.Native]::GetWindowThreadProcessId([System.IntPtr][int64]$Excel.Hwnd, [ref]$processId)
    return [int]$processId
}

function Test-WorkbookRibbon {
    # Opens the built workbook again, now that no macro runs, and waits up to
    # 20 seconds for its Ribbon: onLoad sets KPR_UI_RibbonLoaded, and the
    # Functions menu must list the KPR_ functions. Excel loads a workbook's
    # Ribbon only while no macro runs, so this cannot be done from VBA.
    param($Excel, [string]$Path)
    $book = $Excel.Workbooks.Open($Path)
    $own = "'" + $book.Name.Replace("'", "''") + "'!"
    try {
        for ($attempt = 0; $attempt -lt 40; $attempt++) {
            Start-Sleep -Milliseconds 500
            try {
                if ([bool]$Excel.Run($own + "KPR_UI_RibbonLoaded")) {
                    $menu = [string]$Excel.Run($own + "KPR_UI_RibbonMenuXml", "functions", "Date primitives")
                    if ($menu.Contains("KPR_Fn_")) {
                        return "PASS (the K-PRICING tab loaded in $($book.Name) and its Functions menu lists the functions)"
                    }
                    return "FAIL (the K-PRICING tab loaded but its Functions menu is empty)"
                }
            }
            catch {
                # Excel can refuse a call while it is still opening the workbook.
            }
        }
        return "FAIL (Excel did not load the K-PRICING tab in 20 seconds: onLoad never ran)"
    }
    finally {
        $book.Close($false)
    }
}

function New-FromTemplate {
    # Opens the macro-free template read-only and saves it as an .xlsm at Path;
    # the new workbook stays open.
    param($Excel, [string]$Template, [string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
    $book = $Excel.Workbooks.Open($Template, 0, $true)
    $book.SaveAs($Path, $xlOpenXMLWorkbookMacroEnabled)
    return $book
}

function Invoke-Compile {
    # Debug > Compile VBAProject on Book's project. The command is greyed out
    # once the project is compiled, so it still being enabled means an error.
    param($Excel, $Book)
    $vbe = $Excel.VBE
    $vbe.MainWindow.Visible = $true
    $vbe.ActiveVBProject = $Book.VBProject
    $command = $vbe.CommandBars.FindControl(1, $compileControlId)
    if ($null -eq $command) {
        return $false
    }
    if ($command.Enabled) {
        $command.Execute()
    }
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        Start-Sleep -Milliseconds 500
        if (-not $vbe.CommandBars.FindControl(1, $compileControlId).Enabled) {
            return $true
        }
    }
    return $false
}

function Resolve-Clone {
    # -RepoRoot, else the clone this script lives in, else the clone remembered
    # (shared with Check-KPR.ps1). The clone must hold the template and the loader.
    param([string]$Given)
    $key = "HKCU:\Software\K-PRICING\Check"
    $root = $null
    if (-not [string]::IsNullOrWhiteSpace($Given)) {
        if (-not (Test-Path -LiteralPath (Join-Path $Given ".git"))) {
            throw "-RepoRoot is not a clone of the repository (no .git folder): $Given"
        }
        $root = (Resolve-Path -LiteralPath $Given).Path
    }
    elseif (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $candidate = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
        if (Test-Path -LiteralPath (Join-Path $candidate ".git")) {
            $root = $candidate
        }
    }
    if ($null -eq $root) {
        $remembered = Get-RegistryValue -Path $key -Name "RepoRoot"
        if ($null -ne $remembered -and (Test-Path -LiteralPath (Join-Path $remembered ".git"))) {
            $root = $remembered
        }
    }
    if ($null -eq $root) {
        throw ("No clone of K-PRICING was found. Run tools\Build-KPR.cmd from your clone, " +
               "or pass -RepoRoot 'C:\path\to\K-PRICING' once; it is remembered afterwards.")
    }
    foreach ($relative in @("src\workbook\KPR_Template.xlsx", "tests\modules\KPR_Dev_Import.bas", ".github\repository-profile.json")) {
        if (-not (Test-Path -LiteralPath (Join-Path $root $relative) -PathType Leaf)) {
            throw "The clone is incomplete or too old: $relative is missing in $root. Update it with git pull."
        }
    }
    if (-not (Test-Path -LiteralPath $key)) {
        New-Item -Path $key -Force | Out-Null
    }
    Set-ItemProperty -LiteralPath $key -Name "RepoRoot" -Value $root
    return $root
}

function Get-CloneCommit {
    # The commit the clone has checked out, and whether it has uncommitted
    # changes. git when it is installed, else the files of .git; the commit is
    # empty when neither answers.
    param([string]$Root)
    $ErrorActionPreference = "Continue"
    $sha = ""
    $dirty = $false
    $git = Get-Command git.exe -ErrorAction SilentlyContinue
    if ($null -ne $git) {
        $head = (& $git.Source -C $Root rev-parse HEAD 2>$null | Out-String).Trim()
        if ($head -match '^[0-9a-f]{40}$') {
            $sha = $head
            $status = (& $git.Source -C $Root status --porcelain 2>$null | Out-String).Trim()
            $dirty = ($status.Length -gt 0)
        }
    }
    if ($sha.Length -eq 0) {
        $headFile = Join-Path $Root ".git\HEAD"
        if (Test-Path -LiteralPath $headFile -PathType Leaf) {
            $head = (Get-Content -LiteralPath $headFile -Raw).Trim()
            if ($head -match '^[0-9a-f]{40}$') {
                $sha = $head
            }
            elseif ($head -match '^ref:\s*(\S+)$') {
                $ref = $Matches[1]
                $refFile = Join-Path $Root (".git\" + ($ref -replace '/', '\'))
                $packed = Join-Path $Root ".git\packed-refs"
                if (Test-Path -LiteralPath $refFile -PathType Leaf) {
                    $value = (Get-Content -LiteralPath $refFile -Raw).Trim()
                    if ($value -match '^[0-9a-f]{40}$') {
                        $sha = $value
                    }
                }
                elseif (Test-Path -LiteralPath $packed -PathType Leaf) {
                    foreach ($line in (Get-Content -LiteralPath $packed)) {
                        if ($line -match ('^([0-9a-f]{40}) ' + [regex]::Escape($ref) + '$')) {
                            $sha = $Matches[1]
                        }
                    }
                }
            }
        }
    }
    return @{ Sha = $sha; Dirty = $dirty }
}

function Write-Loader {
    # The loader of the clone with the CRLF line ends the VBE imports.
    param([string]$Source, [string]$Destination)
    $text = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($Source))
    $text = $text -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($Destination, $text, [System.Text.Encoding]::ASCII)
}

# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
$code = 1
$excel = $null
$work = $null
$bookPath = $null
$ribbonResult = "NOT RUN"
$watchdog = $null
$doneFlag = $null
$timeoutFlag = $null
$startedAt = Get-Date
try {
    Write-Host "K-PRICING workbook build ($Mode)" -ForegroundColor Cyan
    Write-Host "Leave Excel alone until this window says it is finished." -ForegroundColor Cyan
    Write-Host ""

    $root = Resolve-Clone -Given $RepoRoot
    $commit = Get-CloneCommit -Root $root
    $outDir = Join-Path $root "build"
    $label = $(if ($commit.Sha.Length -gt 0) { $commit.Sha } else { "unknown (neither git nor .git could name the commit)" })
    if ($commit.Dirty) {
        $label = $label + " + uncommitted changes"
    }
    $short = $(if ($commit.Sha.Length -gt 0) { $commit.Sha.Substring(0, 8) } else { "unknown" })
    Say "Clone: $root"
    Say "Commit: $label"
    $roles = $(if ($Mode -eq "Development") { "public,internal,ui,example,test" } else { "public,internal,ui,example" })
    $baseName = $(if ($Mode -eq "Development") { "K-PRICING-Development-$short" } else { "K-PRICING-$short" })

    $officeVersion = Get-OfficeVersion
    $access = Confirm-VbaProjectAccess -OfficeVersion $officeVersion
    Say "Office $officeVersion; VBA project access: $access"

    $work = Join-Path $env:TEMP ("kpr-build-" + $startedAt.ToString("yyyyMMdd-HHmmss"))
    New-Item -ItemType Directory -Path $work | Out-Null
    $loaderPath = Join-Path $work "KPR_Dev_Import.bas"
    Write-Loader -Source (Join-Path $root "tests\modules\KPR_Dev_Import.bas") -Destination $loaderPath
    $templatePath = Join-Path $root "src\workbook\KPR_Template.xlsx"

    # A private Excel: no add-in or Personal Macro Workbook is loaded in it
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    Say "Excel $($excel.Version) (build $($excel.Build)) started for this build"
    $bookName = $baseName + ".xlsm"
    $book = New-FromTemplate -Excel $excel -Template $templatePath -Path (Join-Path $work $bookName)
    $scratch = $excel.Workbooks.Add()
    try {
        [void]$scratch.VBProject.VBComponents.Import($loaderPath)
    }
    catch {
        throw ("Excel refused access to the VBA project: " + $_.Exception.Message +
               " Close every Excel window and run the build again.")
    }
    # Debug > Compile reports its state only with the VBA editor shown.
    $excel.VBE.MainWindow.Visible = $true

    # Watchdog: closes this run's Excel if the build does not finish in time
    $doneFlag = Join-Path $work "done.flag"
    $timeoutFlag = Join-Path $work "timeout.flag"
    $excelPid = Get-ExcelProcessId -Excel $excel
    $doneQuoted = $doneFlag.Replace("'", "''")
    $timeoutQuoted = $timeoutFlag.Replace("'", "''")
    $watch = "Start-Sleep -Seconds $($TimeoutMinutes * 60); " +
             "if (-not (Test-Path -LiteralPath '$doneQuoted')) { " +
             "Set-Content -LiteralPath '$timeoutQuoted' -Value timeout; Stop-Process -Id $excelPid -Force }"
    # A runspace in this process, not a child powershell.exe: some policies
    # forbid PowerShell from starting another PowerShell.
    $watchdog = [powershell]::Create()
    [void]$watchdog.AddScript($watch)
    [void]$watchdog.BeginInvoke()

    Say "Importing the modules ($roles) from the clone..."
    $qualified = "'" + $scratch.Name + "'!"
    $imported = [int]$excel.Run($qualified + "KPR_Dev_ImportModules", $root, $bookName, $roles)
    if ($imported -lt 0) {
        throw "The modules could not be imported; the VBA editor's Immediate window (Ctrl+G) says why."
    }
    Say "Imported $imported modules"
    Say "Compiling..."
    if (-not (Invoke-Compile -Excel $excel -Book $book)) {
        throw "The workbook does not compile; the VBA editor shows the line."
    }
    $book.Worksheets.Item("Home").Range($templateCell).Value2 = $label
    [void]$book.Activate()
    $own = "'" + $bookName.Replace("'", "''") + "'!"
    if (-not [bool]$excel.Run($own + "KPR_Demo_BuildDates")) {
        throw ("The demo sheet was not built: " + [string]$excel.Run($own + "KPR_Demo_LastReport"))
    }
    [void]$book.Worksheets.Item("Home").Activate()
    $book.Save()
    $book.Close($false)

    # Built in the temporary folder and copied out only when complete, under a
    # new name when the target already exists: a saved workbook is never replaced.
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    $bookPath = Join-Path $outDir $bookName
    if (Test-Path -LiteralPath $bookPath) {
        $bookPath = Join-Path $outDir ($baseName + "-" + (Get-Date).ToString("yyyyMMdd-HHmmss") + ".xlsm")
    }
    Copy-Item -LiteralPath (Join-Path $work $bookName) -Destination $bookPath
    Say "Workbook: $bookPath" "Green"
    $code = 0

    if (-not $NoRibbonCheck) {
        $ribbonResult = Test-WorkbookRibbon -Excel $excel -Path $bookPath
        Say ("Ribbon: " + $ribbonResult) $(if ($ribbonResult -like "PASS*") { "Green" } else { "Red" })
        if ($ribbonResult -notlike "PASS*") {
            $code = 1
        }
    }
    Set-Content -LiteralPath $doneFlag -Value done
}
catch {
    if ($null -ne $timeoutFlag -and (Test-Path -LiteralPath $timeoutFlag)) {
        Say "STOPPED: the build did not finish in $TimeoutMinutes minutes; its Excel was closed." "Red"
    }
    else {
        Say ("STOPPED: " + $_.Exception.Message) "Red"
    }
    $code = 1
}
finally {
    if ($null -ne $doneFlag) {
        Set-Content -LiteralPath $doneFlag -Value done -ErrorAction SilentlyContinue
    }
    if ($null -ne $watchdog) {
        try {
            $watchdog.Stop()
            $watchdog.Dispose()
        }
        catch {
        }
    }
    if ($null -ne $excel) {
        try {
            foreach ($leftOpen in @($excel.Workbooks)) {
                $leftOpen.Close($false)
            }
            $excel.Quit()
        }
        catch {
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Say ("Finished in {0:N0} seconds" -f ((Get-Date) - $startedAt).TotalSeconds)
Write-Host ""
if ($null -ne $bookPath -and (Test-Path -LiteralPath $bookPath)) {
    Write-Host "Your workbook: $bookPath" -ForegroundColor Green
    Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$bookPath`""
    if ($Open -and $code -eq 0) {
        Start-Process -FilePath $bookPath
    }
}
else {
    Write-Host "No workbook was built. Copy this window's text (right-click the title bar, Edit > Select All, Enter) and send it." -ForegroundColor Yellow
}
if ($code -eq 0 -and $null -ne $work) {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host $(if ($code -eq 0) { "RESULT=PASS" } else { "RESULT=FAIL" }) -ForegroundColor $(if ($code -eq 0) { "Green" } else { "Red" })
if (-not $NoPause) {
    [void](Read-Host "Press Enter to close this window")
}
exit $code
