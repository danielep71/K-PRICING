<#
.SYNOPSIS
One-click Excel check and build of K-PRICING: from one commit and the workbook
template, it checks a development workbook (compile, regression twice, Ribbon)
and builds the K-PRICING workbook with its demo sheet, unattended.
.DESCRIPTION
Run it by pasting one line into Windows Run (Win+R), by the desktop shortcut
it creates on the first run, or with tools\Check-KPR.cmd from a clone. It:
1. asks GitHub for the exact commit the branch points to;
2. checks "Trust access to the VBA project object model" and, only when you
   answer Y, turns it on for your Windows user;
3. starts its own Excel (your open workbooks and installed add-ins are not
   touched) and saves that commit's template src\workbook\KPR_Template.xlsx
   as K-PRICING_Development.xlsm in a temporary folder, next to a scratch
   workbook holding the commit's tests\modules\KPR_Dev_Import.bas (a commit
   without the template gets an empty KPR.xlam instead);
4. runs KPR_Dev_CheckBuild unattended on it: import of every module, compile,
   save and KPR_Test_RunAll twice;
5. builds the workbook to use, K-PRICING-<commit>.xlsm in the build folder of
   your clone (-RepoRoot, remembered), or in Documents\K-PRICING without one:
   the template with the production and example modules, compiled, the
   commit written on the Home sheet and the KPR Dates Demo sheet built once
   by KPR_Demo_BuildDates; it then opens the workbook again, once no macro
   runs, and checks that Excel loads its K-PRICING tab and builds the
   Functions menu (Excel loads a workbook's Ribbon only while no macro runs);
6. adds check.log (Windows, Excel, branch, commit, each step) to the evidence
   folder Documents\KPR-evidence\<date>-<commit>, zips it, copies the summary
   and both evidence files to the clipboard and shows the workbook in
   Explorer.
A watchdog closes this run's Excel after -TimeoutMinutes, so a dialog nobody
answers cannot hang the check. The script never pushes, and changes no other
Excel or Windows setting.
.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -Command "&([scriptblock]::Create((irm https://raw.githubusercontent.com/danielep71/K-PRICING/main/tools/Check-KPR.ps1)))"
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Check-KPR.ps1 -Branch my-branch
.EXAMPLE
&([scriptblock]::Create((irm https://raw.githubusercontent.com/danielep71/K-PRICING/main/tools/Check-KPR.ps1))) -RepoRoot 'C:\path\to\K-PRICING'

Inside a PowerShell window that is already open; it starts no other
PowerShell, so it also works where a policy forbids that. -RepoRoot names
your clone once: the workbook goes into its build folder, and later runs
remember it.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Branch = "main",
    [Parameter(Mandatory = $false)]
    [int]$TimeoutMinutes = 30,
    [Parameter(Mandatory = $false)]
    [switch]$NoShortcut,
    [Parameter(Mandatory = $false)]
    [switch]$NoPause,
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repository = "danielep71/K-PRICING"
$rawBase = "https://raw.githubusercontent.com/$repository"
$scriptUrl = "$rawBase/main/tools/Check-KPR.ps1"
$xlOpenXMLAddIn = 55
$xlOpenXMLWorkbookMacroEnabled = 52
$compileControlId = 578
$productionRoles = "public,internal,ui,example"
$templateCell = "D22"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
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
    Write-Host "The check needs it to import the modules. It is the same switch as"
    Write-Host "File > Options > Trust Center > Trust Center Settings > Macro Settings."
    $answer = Read-Host "Turn it on now for your Windows user? (Y/N)"
    if ($answer -notmatch '^\s*[yYsS]') {
        throw "Stopped: the VBA project access stays off, so nothing was checked."
    }
    if (-not (Test-Path -LiteralPath $userKey)) {
        New-Item -Path $userKey -Force | Out-Null
    }
    Set-ItemProperty -LiteralPath $userKey -Name "AccessVBOM" -Value 1 -Type DWord
    return "On (turned on by this check, at the operator's request)"
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

function Resolve-OutputFolder {
    # The workbook goes into <clone>\build (ignored by git) when the clone is
    # known: -RepoRoot, then the clone this script runs from, then the clone
    # remembered from an earlier -RepoRoot; otherwise Documents\K-PRICING.
    param([string]$Given)
    $key = "HKCU:\Software\K-PRICING\Check"
    $root = $null
    if (-not [string]::IsNullOrWhiteSpace($Given)) {
        if (-not (Test-Path -LiteralPath (Join-Path $Given ".git"))) {
            throw "-RepoRoot is not a clone of the repository (no .git folder): $Given"
        }
        $root = (Resolve-Path -LiteralPath $Given).Path
        if (-not (Test-Path -LiteralPath $key)) {
            New-Item -Path $key -Force | Out-Null
        }
        Set-ItemProperty -LiteralPath $key -Name "RepoRoot" -Value $root
        Say "Clone remembered for the next checks: $root"
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
    if ($null -ne $root) {
        return (Join-Path $root "build")
    }
    return (Join-Path ([Environment]::GetFolderPath("MyDocuments")) "K-PRICING")
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

function New-DesktopShortcut {
    $desktop = [Environment]::GetFolderPath("Desktop")
    $path = Join-Path $desktop "K-PRICING check.lnk"
    if (Test-Path -LiteralPath $path) {
        return
    }
    $shell = New-Object -ComObject WScript.Shell
    $link = $shell.CreateShortcut($path)
    $link.TargetPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    $link.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"&([scriptblock]::Create((Invoke-RestMethod '$scriptUrl')))`""
    $link.Description = "Excel check of the latest K-PRICING main"
    $link.IconLocation = "shell32.dll,167"
    $link.Save()
    Say "Desktop shortcut 'K-PRICING check' created: next time, double-click it."
}

function Send-ToClipboard {
    param([string]$Folder)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($name in @("summary.txt", "check.log", "run-1\kpr-test-evidence.json", "run-2\kpr-test-evidence.json")) {
        $file = Join-Path $Folder $name
        if (Test-Path -LiteralPath $file) {
            $parts.Add("===== $name =====")
            $parts.Add([System.IO.File]::ReadAllText($file))
        }
    }
    $text = $parts -join "`r`n"
    if ($text.Length -gt 1000000) {
        $text = [System.IO.File]::ReadAllText((Join-Path $Folder "summary.txt"))
        Say "The evidence is too large for the clipboard: only summary.txt was copied; send the zip too."
    }
    Set-Clipboard -Value $text
}

# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
$code = 1
$excel = $null
$work = $null
$folder = ""
$workbookPath = $null
$ribbonResult = "NOT RUN (no workbook was built)"
$watchdog = $null
$doneFlag = $null
$timeoutFlag = $null
$startedAt = Get-Date
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Write-Host "K-PRICING Excel check" -ForegroundColor Cyan
    Write-Host "Leave Excel alone until this window says it is finished." -ForegroundColor Cyan
    Write-Host ""

    # Exact commit
    $commit = Invoke-RestMethod -UseBasicParsing -Uri "https://api.github.com/repos/$repository/commits/$Branch" `
        -Headers @{ "User-Agent" = "kpr-check"; "Accept" = "application/vnd.github+json" }
    $sha = [string]$commit.sha
    if ($sha -notmatch '^[0-9a-f]{40}$') {
        throw "GitHub did not return a commit for the branch '$Branch'."
    }
    Say "Branch $Branch is at commit $sha"
    $outDir = Resolve-OutputFolder -Given $RepoRoot
    Say "The workbook will be saved in $outDir"

    # Environment, and the one Excel option the check needs
    $officeVersion = Get-OfficeVersion
    $access = Confirm-VbaProjectAccess -OfficeVersion $officeVersion
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $platform = Get-RegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "Platform"
    $build = Get-RegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "VersionToReport"
    Say ("Windows: {0} ({1}); Office {2}, {3}; VBA project access: {4}" -f
        $os.Caption, $os.Version, $(if ($build) { $build } else { $officeVersion }),
        $(switch ($platform) { "x64" { "64-bit" } "x86" { "32-bit" } default { "bitness unknown" } }), $access)

    # The loader of the same commit, with the CRLF line ends the VBE imports
    $work = Join-Path $env:TEMP ("kpr-check-" + $startedAt.ToString("yyyyMMdd-HHmmss"))
    New-Item -ItemType Directory -Path $work | Out-Null
    $loaderPath = Join-Path $work "KPR_Dev_Import.bas"
    $loader = (Invoke-WebRequest -UseBasicParsing -Uri "$rawBase/$sha/tests/modules/KPR_Dev_Import.bas").Content
    if ($loader -is [byte[]]) {
        $loader = [System.Text.Encoding]::ASCII.GetString($loader)
    }
    $loader = $loader -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($loaderPath, $loader, [System.Text.Encoding]::ASCII)
    Say "Downloaded the loader of $($sha.Substring(0, 8))"
    $templatePath = Join-Path $work "KPR_Template.xlsx"
    try {
        Invoke-WebRequest -UseBasicParsing -Uri "$rawBase/$sha/src/workbook/KPR_Template.xlsx" -OutFile $templatePath
        Say "Downloaded the workbook template of $($sha.Substring(0, 8))"
    }
    catch {
        # Only a confirmed 404 means the commit predates the template; any
        # other failure stops the run rather than checking the wrong build.
        $response = $_.Exception.Response
        if ($null -eq $response -or [int]$response.StatusCode -ne 404) {
            throw ("The workbook template could not be downloaded: " + $_.Exception.Message)
        }
        $templatePath = $null
        Say "This commit has no workbook template: the check uses an empty KPR.xlam and builds no workbook."
    }

    # A private Excel: no add-in or Personal Macro Workbook is loaded in it
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    Say "Excel $($excel.Version) (build $($excel.Build)) started for this check"
    if ($null -ne $templatePath) {
        $devName = "K-PRICING_Development.xlsm"
        $target = New-FromTemplate -Excel $excel -Template $templatePath -Path (Join-Path $work $devName)
    }
    else {
        $devName = "KPR.xlam"
        $target = $excel.Workbooks.Add()
        $target.SaveAs((Join-Path $work $devName), $xlOpenXMLAddIn)
    }
    $scratch = $excel.Workbooks.Add()
    try {
        [void]$scratch.VBProject.VBComponents.Import($loaderPath)
    }
    catch {
        throw ("Excel refused access to the VBA project: " + $_.Exception.Message +
               " Close every Excel window and run the check again.")
    }
    # Debug > Compile reports its state only with the VBA editor shown.
    $excel.VBE.MainWindow.Visible = $true

    # Watchdog: closes this run's Excel if the check does not finish in time
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

    Say "Importing, compiling and running the regression twice: this takes a few minutes."
    Say "If a 'Compile error' box appears in the VBA editor, click OK: the check then reports it."
    $qualified = "'" + $scratch.Name + "'!"
    $passed = [bool]$excel.Run($qualified + "KPR_Dev_CheckBuild", $sha, $devName, $true)
    $folder = [string]$excel.Run($qualified + "KPR_Dev_LastCheckFolder")
    Say ("KPR_Dev_CheckBuild: " + $(if ($passed) { "PASS" } else { "FAIL" })) $(if ($passed) { "Green" } else { "Red" })
    $code = $(if ($passed) { 0 } else { 1 })
    $target.Close($false)

    # The workbook to use: production and example modules, commit, demo sheet
    if ($null -ne $templatePath) {
        Say "Building the workbook: production modules, the commit on Home and the demo sheet."
        # Built in the temporary folder and copied out only when complete, under
        # a new name when the folder already has one: a workbook you saved there
        # is never replaced.
        $bookName = "K-PRICING-" + $sha.Substring(0, 8) + ".xlsm"
        $book = New-FromTemplate -Excel $excel -Template $templatePath -Path (Join-Path $work $bookName)
        $imported = [int]$excel.Run($qualified + "KPR_Dev_ImportModules", $sha, $bookName, $productionRoles)
        if ($imported -lt 0) {
            throw "The workbook's modules could not be imported."
        }
        if (-not (Invoke-Compile -Excel $excel -Book $book)) {
            throw "The workbook does not compile with the production and example modules; the VBA editor shows the line."
        }
        $book.Worksheets.Item("Home").Range($templateCell).Value2 = $sha
        [void]$book.Activate()
        $own = "'" + $bookName.Replace("'", "''") + "'!"
        if (-not [bool]$excel.Run($own + "KPR_Demo_BuildDates")) {
            throw ("The demo sheet was not built: " + [string]$excel.Run($own + "KPR_Demo_LastReport"))
        }
        [void]$book.Worksheets.Item("Home").Activate()
        $book.Save()
        $book.Close($false)
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
        $bookPath = Join-Path $outDir $bookName
        if (Test-Path -LiteralPath $bookPath) {
            $bookPath = Join-Path $outDir ("K-PRICING-" + $sha.Substring(0, 8) + "-" +
                (Get-Date).ToString("yyyyMMdd-HHmmss") + ".xlsm")
        }
        Copy-Item -LiteralPath (Join-Path $work $bookName) -Destination $bookPath
        $workbookPath = $bookPath
        Say "Workbook: PASS ($imported modules, demo sheet built) $bookPath" "Green"
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
        Say "STOPPED: the check did not finish in $TimeoutMinutes minutes; its Excel was closed." "Red"
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
            foreach ($book in @($excel.Workbooks)) {
                $book.Close($false)
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

# ----------------------------------------------------------------------------
# Evidence: check.log, zip, clipboard, Explorer
# ----------------------------------------------------------------------------
Say ("Finished in {0:N0} minutes" -f ((Get-Date) - $startedAt).TotalMinutes)
if (-not [string]::IsNullOrWhiteSpace($folder) -and (Test-Path -LiteralPath $folder)) {
    [System.IO.File]::WriteAllText((Join-Path $folder "check.log"), (($script:log -join "`r`n") + "`r`n"), $utf8NoBom)
    $workbookLine = $(if ($null -ne $workbookPath) { "Workbook: PASS ($workbookPath)" } else { "Workbook: NOT BUILT (see check.log)" })
    [System.IO.File]::AppendAllText((Join-Path $folder "summary.txt"),
        $workbookLine + "`r`n" + "Ribbon (workbook): " + $ribbonResult + "`r`n", $utf8NoBom)
    $zip = $folder.TrimEnd("\") + ".zip"
    Compress-Archive -Path (Join-Path $folder "*") -DestinationPath $zip -Force
    Write-Host ""
    Write-Host ([System.IO.File]::ReadAllText((Join-Path $folder "summary.txt")))
    Send-ToClipboard -Folder $folder
    $shown = $(if ($null -ne $workbookPath) { $workbookPath } else { $zip })
    Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$shown`""
    if ($null -ne $workbookPath) {
        Write-Host "Your workbook: $workbookPath" -ForegroundColor Green
    }
    Write-Host "The result is copied to the clipboard: paste it (Ctrl+V) in the chat with Claude." -ForegroundColor Cyan
    Write-Host "The same files are in $zip" -ForegroundColor Cyan
    if ($code -eq 0 -and $null -ne $work) {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}
else {
    Write-Host ""
    Write-Host "No evidence folder was written. Copy this window's text (right-click the title bar, Edit > Select All, Enter) and send it." -ForegroundColor Yellow
}
if (-not $NoShortcut) {
    try {
        New-DesktopShortcut
    }
    catch {
        Say ("The desktop shortcut could not be created: " + $_.Exception.Message)
    }
}
Write-Host ""
Write-Host $(if ($code -eq 0) { "RESULT=PASS" } else { "RESULT=FAIL" }) -ForegroundColor $(if ($code -eq 0) { "Green" } else { "Red" })
if (-not $NoPause) {
    [void](Read-Host "Press Enter to close this window")
}
exit $code
