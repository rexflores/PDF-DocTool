$PDF24 = "C:\Program Files\PDF24\pdf24-DocTool.exe"

# --- HELPERS -----------------------------------------------------------------

function Write-ErrorLog {
    param([string]$Message, [Exception]$Ex = $null)
    $sd = $PSScriptRoot
    if (-not $sd) { $sd = Split-Path $MyInvocation.MyCommand.Path -Parent }
    $logFile = Join-Path $sd "error_log.txt"
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $fullMsg = "[$ts] ERROR: $Message"
    if ($Ex) {
        $fullMsg += "`r`nException: $($Ex.Message)`r`nStackTrace: $($Ex.StackTrace)`r`n"
    }
    try { Add-Content -Path $logFile -Value $fullMsg -Force } catch {}
}

function Clean-UserPathInput {
    param([string]$RawInput)
    $clean = $RawInput
    while ($true) {
        $previous = $clean
        $clean = $clean.Trim().Trim('"').Trim("'").Trim()
        if ($clean -eq $previous) { break }
    }
    return $clean
}

function Get-FolderPathFromUser {
    param([string]$Prompt = "Enter folder path (quotes optional)")
    while ($true) {
        Write-Host ""
        $raw = Read-Host "  $Prompt"
        if ([string]::IsNullOrWhiteSpace($raw)) {
            Write-Host "  Path cannot be empty." -ForegroundColor Yellow
            continue
        }
        $clean = Clean-UserPathInput -RawInput $raw
        if (-not (Test-Path $clean -PathType Container)) {
            Write-Host "  Folder not found: $clean" -ForegroundColor Red
            continue
        }
        return (Resolve-Path $clean).Path.TrimEnd("\")
    }
}

function Get-FilePathFromUser {
    param([string]$Prompt = "Enter file path (quotes optional)")
    while ($true) {
        Write-Host ""
        $raw = Read-Host "  $Prompt"
        if ([string]::IsNullOrWhiteSpace($raw)) {
            Write-Host "  Path cannot be empty." -ForegroundColor Yellow
            continue
        }
        $clean = Clean-UserPathInput -RawInput $raw
        if (-not (Test-Path $clean -PathType Leaf)) {
            Write-Host "  File not found: $clean" -ForegroundColor Red
            continue
        }
        return (Resolve-Path $clean).Path
    }
}

function Show-FilePicker {
    param(
        [string]$Title       = "Select a PDF file",
        [string]$InitialDir  = "",
        [string]$Filter      = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
    )

    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title       = $Title
    $dialog.Filter      = $Filter
    $dialog.Multiselect = $false

    if ($InitialDir -and (Test-Path $InitialDir -PathType Container)) {
        $dialog.InitialDirectory = $InitialDir
    }

    $result = $dialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

function Show-MultiFilePicker {
    param(
        [string]$Title      = "Select PDF files to queue",
        [string]$InitialDir = "",
        [string]$Filter     = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
    )

    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title       = $Title
    $dialog.Filter      = $Filter
    $dialog.Multiselect = $true

    if ($InitialDir -and (Test-Path $InitialDir -PathType Container)) {
        $dialog.InitialDirectory = $InitialDir
    }

    $result = $dialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileNames
    }
    return @()
}


function Send-ToRecycleBin {
    param([string]$Path)
    $shell = New-Object -ComObject Shell.Application
    $parent = Split-Path $Path -Parent
    $name   = Split-Path $Path -Leaf
    $item   = $shell.Namespace($parent).ParseName($name)
    if ($null -eq $item) { throw "Shell could not locate: $Path" }
    $item.InvokeVerb("delete")
}

function Get-YearInput {
    param([string]$Label)
    while ($true) {
        $raw = Read-Host "  $Label (leave blank to skip)"
        $raw = $raw.Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { return "" }
        if ($raw -match '^\d{4}$') {
            $year = [int]$raw
            $maxYear = (Get-Date).Year + 2
            if ($year -gt $maxYear) {
                Write-Host "  Year cannot exceed $maxYear." -ForegroundColor Yellow
                continue
            }
            return $raw
        }
        Write-Host "  Enter a 4-digit year (e.g. 2025)." -ForegroundColor Yellow
    }
}

function Get-YesNo {
    param([string]$Prompt)
    while ($true) {
        $input = (Read-Host "  $Prompt (Y/N)").Trim().ToUpper()
        if ($input -eq "Y") { return $true }
        if ($input -eq "N") { return $false }
        Write-Host "  Please type Y or N." -ForegroundColor Yellow
    }
}

function Get-QueueExitChoice {
    Write-Host ""
    while ($true) {
        Write-Host "  What would you like to do?" -ForegroundColor Yellow
        Write-Host "    S -> Skip this item and continue with the next" -ForegroundColor Yellow
        Write-Host "    X -> Exit the entire queue and return to main menu" -ForegroundColor Yellow
        $choice = (Read-Host "  Enter choice").Trim().ToUpper()
        if ($choice -eq "S") { return "SKIP" }
        if ($choice -eq "X") {
            $script:queueAborted = $true
            return "EXIT"
        }
        Write-Host "  Please type S or X." -ForegroundColor Yellow
    }
}

# Script-level abort flag -- set to $true to stop the current queue loop
$script:queueAborted = $false

function Build-RecordFilename {
    param(
        [string]$Name,
        [string]$LatestYear,
        [string]$OldestYear,
        [string]$Department
    )

    $deptSuffix = switch ($Department) {
        "PT"     { "_PT" }
        "Dental" { "_Dental" }
        default  { "" }
    }

    if ([string]::IsNullOrWhiteSpace($LatestYear) -and [string]::IsNullOrWhiteSpace($OldestYear)) {
        return "$Name$deptSuffix.pdf"
    }

    if ([string]::IsNullOrWhiteSpace($OldestYear) -or $LatestYear -eq $OldestYear) {
        return "${Name}_${LatestYear}$deptSuffix.pdf"
    }

    $latest  = [Math]::Max([int]$LatestYear,  [int]$OldestYear)
    $oldest  = [Math]::Min([int]$LatestYear,  [int]$OldestYear)
    return "${Name}_${latest}_${oldest}$deptSuffix.pdf"
}

function Build-RecordFolderPath {
    param(
        [string]$RootPath,
        [string]$Status,
        [string]$Name,
        [string]$Letter = ""
    )
    if (-not $Letter) {
        $Letter = if ($Name.Length -gt 0) { $Name[0].ToString().ToUpper() } else { "#" }
    }
    return Join-Path $RootPath "$Status\$Letter\$Name"
}

function Get-RecordMetadataFromUser {
    param(
        [string]$DefaultName        = "",
        [string]$DefaultLatestYear  = "",
        [string]$DefaultOldestYear  = "",
        [string]$DefaultStatus      = "",
        [string]$DefaultDepartment  = ""
    )

    Write-Host ""
    Write-Host "  --- Record Metadata ---" -ForegroundColor Cyan
    Write-Host "  (Type CANCEL at the name prompt to go back)" -ForegroundColor DarkGray

    # Name
    if ($DefaultName) {
        $nameInput = (Read-Host "  Employee name (Enter to keep '$DefaultName', or CANCEL to go back)").Trim()
        if ($nameInput.ToUpper() -eq "CANCEL") {
            $exitChoice = Get-QueueExitChoice
            return [PSCustomObject]@{ Cancelled = $true; ExitQueue = ($exitChoice -eq "EXIT") }
        }
        if ([string]::IsNullOrWhiteSpace($nameInput)) { $nameInput = $DefaultName }
    } else {
        while ($true) {
            $nameInput = (Read-Host "  Employee name (e.g. Dela Cruz, Juan) or CANCEL to go back").Trim()
            if ($nameInput.ToUpper() -eq "CANCEL") {
                $exitChoice = Get-QueueExitChoice
                return [PSCustomObject]@{ Cancelled = $true; ExitQueue = ($exitChoice -eq "EXIT") }
            }
            if ($nameInput) { break }
            Write-Host "  Name is required." -ForegroundColor Yellow
        }
    }

    # Status
    $statusDefault = if ($DefaultStatus -eq "RETIREE") { "2" } else { "1" }
    $statusLabel   = if ($DefaultStatus -eq "RETIREE") { "default: 2" } else { "default: 1" }
    $status = if ($DefaultStatus -eq "RETIREE") { "RETIREE" } else { "ACTIVE" }
    while ($true) {
        Write-Host ""
        Write-Host "  Status:" -ForegroundColor Cyan
        Write-Host "    1 -> Active" -ForegroundColor Cyan
        Write-Host "    2 -> Retiree" -ForegroundColor Cyan
        $s = (Read-Host "  Enter choice ($statusLabel)").Trim()
        if ($s -eq "")  { break }
        if ($s -eq "1") { $status = "ACTIVE";  break }
        if ($s -eq "2") { $status = "RETIREE"; break }
        Write-Host "  Enter 1 or 2." -ForegroundColor Yellow
    }

    # Department
    $deptDefault = switch ($DefaultDepartment) {
        "PT"     { "2" }
        "Dental" { "3" }
        default  { "1" }
    }
    $deptLabel = switch ($DefaultDepartment) {
        "PT"     { "default: 2" }
        "Dental" { "default: 3" }
        default  { "default: 1" }
    }
    $dept = if ($DefaultDepartment -in @("PT","Dental","Medical")) { $DefaultDepartment } else { "Medical" }
    while ($true) {
        Write-Host ""
        Write-Host "  Department:" -ForegroundColor Cyan
        Write-Host "    1 -> Medical (default)" -ForegroundColor Cyan
        Write-Host "    2 -> PT" -ForegroundColor Cyan
        Write-Host "    3 -> Dental" -ForegroundColor Cyan
        $d = (Read-Host "  Enter choice ($deptLabel)").Trim()
        if ($d -eq "")  { break }
        if ($d -eq "1") { $dept = "Medical"; break }
        if ($d -eq "2") { $dept = "PT";      break }
        if ($d -eq "3") { $dept = "Dental";  break }
        Write-Host "  Enter 1, 2, or 3." -ForegroundColor Yellow
    }

    # Years -- pre-filled from existing filename if available
    Write-Host ""
    Write-Host "  Year range (at least Latest Year is required):" -ForegroundColor DarkGray
    if ($DefaultLatestYear) {
        Write-Host "  Press Enter to keep existing years." -ForegroundColor DarkGray
    } else {
        Write-Host "  If Oldest Year is left blank, it will match Latest Year (single-year filename)." -ForegroundColor DarkGray
    }

    $latestYear = ""
    while ([string]::IsNullOrWhiteSpace($latestYear)) {
        $latestPrompt = if ($DefaultLatestYear) { "Latest year (Enter to keep '$DefaultLatestYear')" } else { "Latest year" }
        $latestYear = Get-YearInput -Label $latestPrompt
        if ([string]::IsNullOrWhiteSpace($latestYear)) {
            if ($DefaultLatestYear) {
                $latestYear = $DefaultLatestYear
            } else {
                Write-Host "  Latest year is required." -ForegroundColor Yellow
            }
        }
    }

    $oldestPrompt = if ($DefaultOldestYear) { "Oldest year (Enter to keep '$DefaultOldestYear')" } else { "Oldest year (blank = same as latest)" }
    $oldestYear = Get-YearInput -Label $oldestPrompt
    if ([string]::IsNullOrWhiteSpace($oldestYear)) {
        if ($DefaultOldestYear) {
            $oldestYear = $DefaultOldestYear
            Write-Host "  Oldest year kept as $oldestYear." -ForegroundColor DarkGray
        } else {
            $oldestYear = $latestYear
            Write-Host "  Oldest year defaulted to $latestYear (single-year filename)." -ForegroundColor DarkGray
        }
    }

    return [PSCustomObject]@{
        Cancelled  = $false
        ExitQueue  = $false
        Name       = $nameInput
        Status     = $status
        Department = $dept
        LatestYear = $latestYear
        OldestYear = $oldestYear
    }
}

function Invoke-PDF24Merge {
    param(
        [string]$NewPdfPath,
        [string]$ExistingPdfPath,
        [string]$OutputPath
    )

    # PDF24 always appends .pdf to -outputFile regardless of what extension
    # we give it. So we pass a temp base name and expect PDF24 to produce
    # that name with .pdf appended.
    $tempBase    = $OutputPath
    $expectedOut = $tempBase + ".pdf"

    $pdf24Args = "-join -noProgress -profile `"default/good`" " +
                 "-outputFile `"$tempBase`" " +
                 "`"$NewPdfPath`" `"$ExistingPdfPath`""

    $stdoutLog = [System.IO.Path]::GetTempFileName()
    $stderrLog = [System.IO.Path]::GetTempFileName()

    # We remove -Wait so we can loop and show a progress bar
    $proc = Start-Process `
        -FilePath               $PDF24 `
        -ArgumentList           $pdf24Args `
        -PassThru `
        -WindowStyle            Hidden `
        -RedirectStandardOutput $stdoutLog `
        -RedirectStandardError  $stderrLog

    $timer = 0
    while (-not $proc.HasExited) {
        # Create an oscillating progress bar (0 to 100 and back)
        $cycle = $timer % 20
        $percent = if ($cycle -le 10) { $cycle * 10 } else { (20 - $cycle) * 10 }
        
        Write-Progress -Activity "Merging via PDF24" -Status "Processing files... Please wait." -PercentComplete $percent
        Start-Sleep -Milliseconds 200
        $timer++
        
        # If running in GUI mode, allow the UI to remain responsive
        if (("System.Windows.Forms.Application" -as [type]) -ne $null) {
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    Write-Progress -Activity "Merging via PDF24" -Completed

    # Clean up the log files quietly
    Remove-Item $stdoutLog -Force -ErrorAction SilentlyContinue
    Remove-Item $stderrLog -Force -ErrorAction SilentlyContinue

    if ($null -ne $proc.ExitCode -and $proc.ExitCode -ne 0) {
        throw "PDF24 exited with code $($proc.ExitCode). Merge may have failed."
    }

    # Wait for PDF24 to flush the output file
    $waited = 0
    while (-not (Test-Path $expectedOut) -and $waited -lt 20) {
        Start-Sleep -Seconds 1
        $waited++
    }

    if (-not (Test-Path $expectedOut)) {
        throw "Merged file was not created. Expected: $expectedOut"
    }

    return $expectedOut
}

function Invoke-SiblingRename {
    param(
        [string]$EmployeeFolder,
        [string]$CorrectName,
        [string]$ExcludeFile = ""
    )

    # Find all PDFs in the folder except the one currently being merged/saved
    $siblings = Get-ChildItem -Path $EmployeeFolder -File -Filter "*.pdf" |
        Where-Object {
            $_.FullName -ne $ExcludeFile -and
            $_.FullName -ne ($ExcludeFile + ".merging.tmp.pdf")
        }

    if ($siblings.Count -eq 0) { return }

    Write-Host ""
    Write-Host "  The employee folder contains $($siblings.Count) other PDF file(s):" -ForegroundColor Yellow
    foreach ($s in $siblings) {
        Write-Host "    $($s.Name)" -ForegroundColor White
    }
    Write-Host ""

    if (-not (Get-YesNo -Prompt "Rename these files to match the corrected employee name")) {
        Write-Host "  Sibling files left unchanged." -ForegroundColor DarkGray
        return
    }

    foreach ($sibling in $siblings) {
        $siblingBase = [System.IO.Path]::GetFileNameWithoutExtension($sibling.Name)

        # Pre-fill years from the sibling's own filename
        $sibLatest = ""
        $sibOldest = ""
        if ($siblingBase -match '_(\d{4})_(\d{4})(_PT|_Dental)?$') {
            $y1 = [int]$Matches[1]; $y2 = [int]$Matches[2]
            $sibLatest = [string][Math]::Max($y1, $y2)
            $sibOldest = [string][Math]::Min($y1, $y2)
        } elseif ($siblingBase -match '_(\d{4})(_PT|_Dental)?$') {
            $sibLatest = $Matches[1]
            $sibOldest = $Matches[1]
        }

        # Pre-fill department from sibling filename
        $sibDept = "Medical"
        if ($siblingBase -match '_PT$')     { $sibDept = "PT" }
        if ($siblingBase -match '_Dental$') { $sibDept = "Dental" }

        Write-Host ""
        Write-Host "  --- Renaming: $($sibling.Name) ---" -ForegroundColor DarkGray
        Write-Host "  Name defaults to the corrected name. Press Enter to keep it." -ForegroundColor DarkGray

        $sibMeta = Get-RecordMetadataFromUser `
            -DefaultName        $CorrectName `
            -DefaultLatestYear  $sibLatest `
            -DefaultOldestYear  $sibOldest `
            -DefaultDepartment  $sibDept

        if ($sibMeta.Cancelled) {
            Write-Host "  Skipped: $($sibling.Name)" -ForegroundColor DarkGray
            if ($sibMeta.ExitQueue) { return }
            continue
        }

        $newSibFilename = Build-RecordFilename `
            -Name       $sibMeta.Name `
            -LatestYear $sibMeta.LatestYear `
            -OldestYear $sibMeta.OldestYear `
            -Department $sibMeta.Department

        $newSibPath = Join-Path $EmployeeFolder $newSibFilename

        if ($newSibPath -eq $sibling.FullName) {
            Write-Host "  No change for: $($sibling.Name)" -ForegroundColor DarkGray
            continue
        }

        if (Test-Path $newSibPath) {
            Write-Host "  WARNING: $newSibFilename already exists. Skipping to avoid overwrite." -ForegroundColor Yellow
            continue
        }

        try {
            Rename-Item -Path $sibling.FullName -NewName $newSibFilename -Force
            Write-Host "  Renamed: $($sibling.Name)" -ForegroundColor DarkGray
            Write-Host "       --> $newSibFilename" -ForegroundColor Green
        } catch {
            Write-Host "  ERROR renaming $($sibling.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# --- WORKFLOW: MERGE EXISTING RECORD -----------------------------------------

function Start-MergeWorkflow {
    param([string]$RootPath)

    $script:QueueAborted = $false
    $downloadsPath = Join-Path $env:USERPROFILE "Downloads"

    Write-Host ""
    Write-Host "  === MERGE INTO EXISTING RECORD ===" -ForegroundColor Cyan
    Write-Host "  New PDF pages go FIRST, existing record pages go AFTER." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  A file picker will open. Select one or more NEW PDFs to queue." -ForegroundColor DarkGray
    Write-Host "  You can hold Ctrl or Shift to select multiple files at once." -ForegroundColor DarkGray
    Write-Host ""

    # Step 1: pick new PDFs from Downloads (multi-select to build a queue)
    $selectedFiles = Show-MultiFilePicker `
        -Title      "Select new PDF(s) to merge - hold Ctrl/Shift for multiple" `
        -InitialDir $downloadsPath

    if (-not $selectedFiles -or $selectedFiles.Count -eq 0) {
        Write-Host "  No files selected. Cancelled." -ForegroundColor Yellow
        return
    }

    $queue = @($selectedFiles | Where-Object { $_.ToLower().EndsWith(".pdf") })

    if ($queue.Count -eq 0) {
        Write-Host "  No PDF files in selection. Cancelled." -ForegroundColor Yellow
        return
    }

    Write-Host ""
    Write-Host "  Queue: $($queue.Count) PDF(s) to process" -ForegroundColor Cyan
    for ($i = 0; $i -lt $queue.Count; $i++) {
        Write-Host "    $($i+1). $(Split-Path $queue[$i] -Leaf)" -ForegroundColor White
    }

    # Process each queued PDF one by one
    $processed = 0
    $failed    = 0
    $script:queueAborted = $false

    foreach ($newPdf in $queue) {
        if ($script:queueAborted) { break }

        $newPdfName = Split-Path $newPdf -Leaf
        Write-Host ""
        Write-Host "  -------------------------------------" -ForegroundColor DarkGray
        Write-Host "  Processing: $newPdfName" -ForegroundColor Cyan
        Write-Host "  -------------------------------------" -ForegroundColor DarkGray

        # Step 2: find the existing employee PDF via search or file picker
        $existingPdf = $null

        while (-not $existingPdf) {
            Write-Host ""
            Write-Host "  Locate the EXISTING employee PDF to merge '$newPdfName' into." -ForegroundColor Cyan
            Write-Host "    S  ->  Search by name in the records root" -ForegroundColor Cyan
            Write-Host "    B  ->  Browse via file picker (opens at root folder)" -ForegroundColor Cyan
            Write-Host "    V  ->  View the new PDF first (opens '$newPdfName')" -ForegroundColor Cyan
            Write-Host "    K  ->  Skip this file" -ForegroundColor Cyan
            Write-Host "    X  ->  Exit queue and return to main menu" -ForegroundColor Cyan
            $locateChoice = (Read-Host "  Enter choice").Trim().ToUpper()

            switch ($locateChoice) {
                "X" {
                    $script:queueAborted = $true
                    $existingPdf = "SKIP"
                    Write-Host "  Queue exited." -ForegroundColor DarkGray
                }
                "V" {
                    Write-Host "  Opening $newPdfName..." -ForegroundColor DarkGray
                    try {
                        Start-Process $newPdf
                    } catch {
                        Write-Host "  Could not open file: $($_.Exception.Message)" -ForegroundColor Red
                    }
                    # Don't set $existingPdf -- loop continues so user can then pick S/B/K
                }
                "S" {
                    $searchName = (Read-Host "  Search name (with or without comma)").Trim()
                    if (-not $searchName) { continue }

                    # Normalize search: strip commas and collapse spaces so
                    # "prasad jay", "prasad, jay", or "jay prasad" all match
                    # "Prasad, Jayneeta" in either the filename or folder name.
                    $searchTokens = ($searchName -replace ',', ' ') -split '\s+' |
                        Where-Object { $_ -ne '' }

                    $found = Get-ChildItem -Path $RootPath -Recurse -File -Filter "*.pdf" |
                        Where-Object {
                            $target = ($_.Name + ' ' + $_.Directory.Name).ToLower()
                            $allMatch = $true
                            foreach ($token in $searchTokens) {
                                if ($target -notlike "*$token*") {
                                    $allMatch = $false
                                    break
                                }
                            }
                            $allMatch
                        }

                    if ($found.Count -eq 0) {
                        Write-Host "  No matches for '$searchName'. Try again or choose B to browse." -ForegroundColor Yellow
                    } elseif ($found.Count -eq 1) {
                        $rel = $found[0].FullName.Substring($RootPath.Length).TrimStart("\")
                        Write-Host "  Found: $rel" -ForegroundColor Green
                        $c = if (Get-YesNo -Prompt "Use this file") { "Y" } else { "N" }
                        if ($c -eq "Y") {
                            $existingPdf = $found[0].FullName
                            # Offer to view the existing PDF before metadata entry
                            if (Get-YesNo -Prompt "View this existing PDF before proceeding") {
                                Write-Host "  Opening $rel..." -ForegroundColor DarkGray
                                try { Start-Process $existingPdf } catch {
                                    Write-Host "  Could not open file: $($_.Exception.Message)" -ForegroundColor Red
                                }
                            }
                        }
                    } else {
                        Write-Host ""
                        Write-Host "  Multiple matches:" -ForegroundColor White
                        for ($i = 0; $i -lt [Math]::Min($found.Count, 20); $i++) {
                            $rel = $found[$i].FullName.Substring($RootPath.Length).TrimStart("\")
                            Write-Host "    $($i+1)  $rel" -ForegroundColor White
                        }
                        Write-Host ""
                        while ($true) {
                            $pick = (Read-Host "  Enter number, or 0 to search again").Trim()
                            if ($pick -match '^\d+$') {
                                $idx = [int]$pick
                                if ($idx -eq 0) { break }
                                if ($idx -ge 1 -and $idx -le $found.Count) {
                                    $existingPdf = $found[$idx - 1].FullName
                                    # Offer to view the existing PDF before metadata entry
                                    $relChosen = $found[$idx - 1].FullName.Substring($RootPath.Length).TrimStart("\")
                                    if (Get-YesNo -Prompt "View this existing PDF before proceeding") {
                                        Write-Host "  Opening $relChosen..." -ForegroundColor DarkGray
                                        try { Start-Process $existingPdf } catch {
                                            Write-Host "  Could not open file: $($_.Exception.Message)" -ForegroundColor Red
                                        }
                                    }
                                    break
                                }
                            }
                            Write-Host "  Enter a valid number." -ForegroundColor Yellow
                        }
                    }
                }
                "B" {
                    $picked = Show-FilePicker `
                        -Title      "Select the EXISTING employee PDF to merge into" `
                        -InitialDir $RootPath
                    if ($picked -and $picked.ToLower().EndsWith(".pdf")) {
                        $existingPdf = $picked
                        $pickedRel = $picked.Substring($RootPath.Length).TrimStart("\")
                        Write-Host "  Selected: $pickedRel" -ForegroundColor Green
                        # Offer to view the existing PDF before metadata entry
                        if (Get-YesNo -Prompt "View this existing PDF before proceeding") {
                            Write-Host "  Opening $pickedRel..." -ForegroundColor DarkGray
                            try { Start-Process $existingPdf } catch {
                                Write-Host "  Could not open file: $($_.Exception.Message)" -ForegroundColor Red
                            }
                        }
                    } elseif ($picked) {
                        Write-Host "  Selected file is not a PDF. Try again." -ForegroundColor Yellow
                    } else {
                        Write-Host "  No file selected." -ForegroundColor Yellow
                    }
                }
                "K" {
                    Write-Host "  Skipped: $newPdfName" -ForegroundColor DarkGray
                    $failed++
                    $existingPdf = "SKIP"
                }
                "X" {
                    $script:queueAborted = $true
                    $existingPdf = "SKIP"
                }
                default {
                    Write-Host "  Enter S, B, V, K, or X." -ForegroundColor Yellow
                }
            }
        }

        if ($existingPdf -eq "SKIP") {
            if ($script:queueAborted) { break }
            continue
        }

        # Step 3: metadata - pre-fill name and years from existing filename
        Write-Host ""
        Write-Host "  Enter details for the MERGED output file." -ForegroundColor Cyan
        $existingBase = [System.IO.Path]::GetFileNameWithoutExtension($existingPdf)
        $defaultName  = $existingBase -replace '(_\d{4})+(_PT|_Dental)?$', '' -replace '(_PT|_Dental)$', ''

        # Extract years from filename pattern Name_YYYY or Name_YYYY_YYYY
        $defaultLatest = ""
        $defaultOldest = ""
        if ($existingBase -match '_(\d{4})_(\d{4})(_PT|_Dental)?$') {
            $yr1 = [int]$Matches[1]
            $yr2 = [int]$Matches[2]
            $defaultLatest = [string][Math]::Max($yr1, $yr2)
            $defaultOldest = [string][Math]::Min($yr1, $yr2)
        } elseif ($existingBase -match '_(\d{4})(_PT|_Dental)?$') {
            $defaultLatest = $Matches[1]
            $defaultOldest = $Matches[1]
        }

        # Derive the existing status and root from the existing file's path
        # so we can pre-fill status and detect folder changes.
        $existingEmployeeFolder = Split-Path $existingPdf -Parent
        $existingEmployeeName   = Split-Path $existingEmployeeFolder -Leaf

        # Walk up the path to find which status folder (ACTIVE/RETIREE) this
        # file lives under, so we can pre-fill the status prompt correctly.
        $detectedStatus = "ACTIVE"
        $pathParts = $existingPdf -split [regex]::Escape([System.IO.Path]::DirectorySeparatorChar)
        foreach ($part in $pathParts) {
            if ($part.ToUpper() -eq "ACTIVE")  { $detectedStatus = "ACTIVE";  break }
            if ($part.ToUpper() -eq "RETIREE") { $detectedStatus = "RETIREE"; break }
        }

        $meta = Get-RecordMetadataFromUser `
            -DefaultName        $defaultName `
            -DefaultLatestYear  $defaultLatest `
            -DefaultOldestYear  $defaultOldest `
            -DefaultStatus      $detectedStatus

        if ($meta.Cancelled) {
            $failed++
            if ($meta.ExitQueue) { break }
            continue
        }

        $outputFilename = Build-RecordFilename `
            -Name       $meta.Name `
            -LatestYear $meta.LatestYear `
            -OldestYear $meta.OldestYear `
            -Department $meta.Department

        # Check if name or status changed -- if so, the employee folder
        # needs to move to its correct location under the new name/status.
        $nameChanged   = $meta.Name   -ne $defaultName
        $statusChanged = $meta.Status -ne $detectedStatus

        if ($nameChanged -or $statusChanged) {
            # Derive the root path by stripping status\letter\employeeName from the existing path
            $existingFolderFull = [System.IO.Path]::GetFullPath($existingEmployeeFolder)
            $rootPathFull       = [System.IO.Path]::GetFullPath($RootPath)

            $newEmployeeFolder = Build-RecordFolderPath `
                -RootPath $rootPathFull `
                -Status   $meta.Status `
                -Name     $meta.Name

            $outputPath = Join-Path $newEmployeeFolder $outputFilename
        } else {
            $newEmployeeFolder = $existingEmployeeFolder
            $outputPath        = Join-Path $existingEmployeeFolder $outputFilename
        }

        Write-Host ""
        Write-Host "  --- Merge Summary ---" -ForegroundColor DarkGray
        Write-Host "  New PDF    : $newPdfName" -ForegroundColor White
        Write-Host "  Existing   : $(Split-Path $existingPdf -Leaf)" -ForegroundColor White
        Write-Host "  Output     : $outputPath" -ForegroundColor White
        if ($nameChanged -or $statusChanged) {
            Write-Host "  Folder move: $existingEmployeeFolder" -ForegroundColor Yellow
            Write-Host "          --> $newEmployeeFolder" -ForegroundColor Yellow
        }
        Write-Host ""

        if (Test-Path $outputPath) {
            Write-Host "  Output filename already exists and will be overwritten." -ForegroundColor Yellow
        }

        $confirmChoice = ""
        while ($confirmChoice -notin @("Y","N","X")) {
            Write-Host "  Y -> Proceed   N -> Skip this file   X -> Exit queue" -ForegroundColor DarkGray
            $confirmChoice = (Read-Host "  Enter choice").Trim().ToUpper()
            if ($confirmChoice -notin @("Y","N","X")) {
                Write-Host "  Please type Y, N, or X." -ForegroundColor Yellow
            }
        }
        if ($confirmChoice -eq "X") {
            $script:queueAborted = $true
            Write-Host "  Queue exited." -ForegroundColor DarkGray
            break
        }
        if ($confirmChoice -eq "N") {
            Write-Host "  Skipped." -ForegroundColor DarkGray
            $failed++
            continue
        }

        $tempBase   = $outputPath + ".merging.tmp"
        $tempActual = $tempBase + ".pdf"

        try {
            # If name/status changed, move the employee folder first so the
            # existing PDF travels with it before we try to merge into it.
            $movedExistingPdf = $existingPdf
            if (($nameChanged -or $statusChanged) -and
                ([System.IO.Path]::GetFullPath($existingEmployeeFolder).ToLower() -ne
                 [System.IO.Path]::GetFullPath($newEmployeeFolder).ToLower())) {

                if (Test-Path $newEmployeeFolder) {
                    Write-Host "  WARNING: Destination folder already exists. Files will be merged into it." -ForegroundColor Yellow
                    # Move individual files rather than the whole folder to avoid collision
                    Get-ChildItem $existingEmployeeFolder -Recurse | ForEach-Object {
                        $destItem = Join-Path $newEmployeeFolder $_.FullName.Substring($existingEmployeeFolder.Length).TrimStart("\")
                        $destDir  = Split-Path $destItem -Parent
                        if (-not (Test-Path $destDir)) {
                            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                        }
                        if ($_.PSIsContainer) { return }
                        Move-Item -Path $_.FullName -Destination $destItem -Force
                    }
                    # Remove old folder if now empty
                    if (-not (Get-ChildItem $existingEmployeeFolder -Recurse -File)) {
                        Remove-Item $existingEmployeeFolder -Recurse -Force -ErrorAction SilentlyContinue
                    }
                } else {
                    $parentDir = Split-Path $newEmployeeFolder -Parent
                    if (-not (Test-Path $parentDir)) {
                        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
                    }
                    Move-Item -Path $existingEmployeeFolder -Destination $newEmployeeFolder -Force
                }

                # Update the existing PDF path to its new location after the move
                $movedExistingPdf = Join-Path $newEmployeeFolder (Split-Path $existingPdf -Leaf)
                Write-Host "  Employee folder moved to: $newEmployeeFolder" -ForegroundColor DarkGray

                # Offer to rename sibling files in the moved folder
                Invoke-SiblingRename `
                    -EmployeeFolder $newEmployeeFolder `
                    -CorrectName    $meta.Name `
                    -ExcludeFile    $movedExistingPdf
            }

            if (-not (Test-Path $newEmployeeFolder)) {
                New-Item -ItemType Directory -Path $newEmployeeFolder -Force | Out-Null
            }

            $actualMergedPath = Invoke-PDF24Merge `
                -NewPdfPath      $newPdf `
                -ExistingPdfPath $movedExistingPdf `
                -OutputPath      $tempBase

            $existingNorm = [System.IO.Path]::GetFullPath($movedExistingPdf).ToLower()
            $outputNorm   = [System.IO.Path]::GetFullPath($outputPath).ToLower()

            if ($existingNorm -eq $outputNorm) {
                Send-ToRecycleBin -Path $movedExistingPdf
            } elseif (Test-Path $outputPath) {
                Remove-Item $outputPath -Force
            }

            Move-Item -Path $actualMergedPath -Destination $outputPath -Force

            Write-Host "  Merged and saved:" -ForegroundColor Green
            Write-Host "    $outputPath" -ForegroundColor Green

            if ($existingNorm -ne $outputNorm -and (Test-Path $movedExistingPdf)) {
                Send-ToRecycleBin -Path $movedExistingPdf
                Write-Host "  Old record recycled: $(Split-Path $movedExistingPdf -Leaf)" -ForegroundColor DarkGray
            }

            if (Get-YesNo -Prompt "Move '$newPdfName' from Downloads to Recycle Bin") {
                Send-ToRecycleBin -Path $newPdf
                Write-Host "  Source PDF recycled." -ForegroundColor DarkGray
            }

            $processed++
            Write-Host "  Done. $processed of $($queue.Count) processed." -ForegroundColor Green

        } catch {
            $errMsg = $_.Exception.Message
            Write-ErrorLog "Merge operation failed for '$newPdfName'" $_.Exception
            Write-Host "  ERROR: $errMsg" -ForegroundColor Red
            if (Test-Path $tempActual) { Remove-Item $tempActual -Force -ErrorAction SilentlyContinue }
            if (Test-Path $tempBase)   { Remove-Item $tempBase   -Force -ErrorAction SilentlyContinue }
            Write-Host "  No files were changed for this item." -ForegroundColor Yellow
            $failed++
        }
    }

    Write-Host ""
    Write-Host "  =====================================" -ForegroundColor DarkGray
    if ($script:queueAborted) {
        Write-Host "  Queue exited early." -ForegroundColor Yellow
    } else {
        Write-Host "  Queue complete." -ForegroundColor Cyan
    }
    Write-Host "  Processed : $processed" -ForegroundColor Green
    if ($failed -gt 0) {
        Write-Host "  Skipped/failed : $failed" -ForegroundColor Yellow
    }
    Write-Host "  =====================================" -ForegroundColor DarkGray
}

function Start-NewRecordWorkflow {
    param([string]$RootPath)

    $downloadsPath = Join-Path $env:USERPROFILE "Downloads"

    Write-Host ""
    Write-Host "  === SAVE NEW RECORD ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  A file picker will open. Select one or more new PDFs to queue." -ForegroundColor DarkGray
    Write-Host "  You can hold Ctrl or Shift to select multiple files at once." -ForegroundColor DarkGray
    Write-Host ""

    # Step 1: multi-select new PDFs from Downloads
    $selectedFiles = Show-MultiFilePicker `
        -Title      "Select new PDF(s) to save as new employee records - hold Ctrl/Shift for multiple" `
        -InitialDir $downloadsPath

    if (-not $selectedFiles -or $selectedFiles.Count -eq 0) {
        Write-Host "  No files selected. Cancelled." -ForegroundColor Yellow
        return
    }

    $queue = @($selectedFiles | Where-Object { $_.ToLower().EndsWith(".pdf") })

    if ($queue.Count -eq 0) {
        Write-Host "  No PDF files in selection. Cancelled." -ForegroundColor Yellow
        return
    }

    Write-Host ""
    Write-Host "  Queue: $($queue.Count) PDF(s) to process" -ForegroundColor Cyan
    for ($i = 0; $i -lt $queue.Count; $i++) {
        Write-Host "    $($i+1). $(Split-Path $queue[$i] -Leaf)" -ForegroundColor White
    }

    $processed = 0
    $failed    = 0
    $script:queueAborted = $false

    foreach ($sourcePdf in $queue) {
        if ($script:queueAborted) { break }

        $sourceName = Split-Path $sourcePdf -Leaf

        Write-Host ""
        Write-Host "  -------------------------------------" -ForegroundColor DarkGray
        Write-Host "  Processing: $sourceName" -ForegroundColor Cyan
        Write-Host "  -------------------------------------" -ForegroundColor DarkGray

        # Show filename and offer to view before naming
        Write-Host ""
        Write-Host "  Selected: $sourceName" -ForegroundColor Green
        if (Get-YesNo -Prompt "View this PDF before naming it") {
            Write-Host "  Opening $sourceName..." -ForegroundColor DarkGray
            try { Start-Process $sourcePdf } catch {
                Write-Host "  Could not open file: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # Step 2: metadata
        $meta = Get-RecordMetadataFromUser

        if ($meta.Cancelled) {
            $failed++
            if ($meta.ExitQueue) { break }
            continue
        }

        $outputFilename = Build-RecordFilename `
            -Name       $meta.Name `
            -LatestYear $meta.LatestYear `
            -OldestYear $meta.OldestYear `
            -Department $meta.Department

        $employeeFolder = Build-RecordFolderPath `
            -RootPath $RootPath `
            -Status   $meta.Status `
            -Name     $meta.Name

        $outputPath = Join-Path $employeeFolder $outputFilename

        Write-Host ""
        Write-Host "  --- Save Summary ---" -ForegroundColor DarkGray
        Write-Host "  Source     : $sourceName" -ForegroundColor White
        Write-Host "  Destination: $outputPath" -ForegroundColor White
        Write-Host ""

        # Warn if folder already exists
        if (Test-Path $employeeFolder) {
            $existing = Get-ChildItem $employeeFolder -File | Select-Object -ExpandProperty Name
            Write-Host "  WARNING: Employee folder already exists." -ForegroundColor Yellow
            Write-Host "  Existing files in that folder:" -ForegroundColor Yellow
            foreach ($f in $existing) { Write-Host "    $f" -ForegroundColor Yellow }
            Write-Host ""
            if (-not (Get-YesNo -Prompt "Continue saving here")) {
                Write-Host "  Skipped." -ForegroundColor DarkGray
                $failed++
                continue
            }
        }

        if (Test-Path $outputPath) {
            Write-Host "  A file with that name already exists at the destination." -ForegroundColor Red
            if (-not (Get-YesNo -Prompt "Overwrite it")) {
                Write-Host "  Skipped." -ForegroundColor DarkGray
                $failed++
                continue
            }
        }

        $confirmChoice = ""
        while ($confirmChoice -notin @("Y","N","X")) {
            Write-Host "  Y -> Proceed   N -> Skip this file   X -> Exit queue" -ForegroundColor DarkGray
            $confirmChoice = (Read-Host "  Enter choice").Trim().ToUpper()
            if ($confirmChoice -notin @("Y","N","X")) {
                Write-Host "  Please type Y, N, or X." -ForegroundColor Yellow
            }
        }
        if ($confirmChoice -eq "X") {
            $script:queueAborted = $true
            Write-Host "  Queue exited." -ForegroundColor DarkGray
            break
        }
        if ($confirmChoice -eq "N") {
            Write-Host "  Skipped." -ForegroundColor DarkGray
            $failed++
            continue
        }

        try {
            if (-not (Test-Path $employeeFolder)) {
                New-Item -ItemType Directory -Path $employeeFolder -Force | Out-Null
                Write-Host "  Created new folder: $employeeFolder" -ForegroundColor DarkGray
            }

            Copy-Item -Path $sourcePdf -Destination $outputPath -Force
            Write-Host "  Saved to: $outputPath" -ForegroundColor Green

            if (Get-YesNo -Prompt "Move '$sourceName' from Downloads to Recycle Bin") {
                Send-ToRecycleBin -Path $sourcePdf
                Write-Host "  Source PDF moved to Recycle Bin." -ForegroundColor DarkGray
            }

            $processed++
            Write-Host "  Done. $processed of $($queue.Count) processed." -ForegroundColor Green

        } catch {
            $errMsg = $_.Exception.Message
            Write-ErrorLog "New record operation failed for '$sourceName'" $_.Exception
            Write-Host "  ERROR: $errMsg" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host ""
    Write-Host "  =====================================" -ForegroundColor DarkGray
    if ($script:queueAborted) {
        Write-Host "  Queue exited early." -ForegroundColor Yellow
    } else {
        Write-Host "  Queue complete." -ForegroundColor Cyan
    }
    Write-Host "  Processed : $processed" -ForegroundColor Green
    if ($failed -gt 0) {
        Write-Host "  Skipped/failed : $failed" -ForegroundColor Yellow
    }
    Write-Host "  =====================================" -ForegroundColor DarkGray
}

# --- MAIN MENU ----------------------------------------------------------------
# Only run the terminal menu when this script is executed directly.
# When loaded by PDFRecordTool-GUI.ps1, $script:PDFRecordToolGUIMode is set
# to $true before dot-sourcing, so the menu loop below is skipped entirely.

if (-not $script:PDFRecordToolGUIMode) {

    if (-not (Test-Path $PDF24)) {
        Write-Host ""
        Write-Host "  ERROR: PDF24 not found at $PDF24" -ForegroundColor Red
        Write-Host "  Please verify PDF24 Creator is installed." -ForegroundColor Red
        Write-Host ""
        exit
    }

    Write-Host ""
    Write-Host "  =====================================" -ForegroundColor Cyan
    Write-Host "       PDF Record Tool                 " -ForegroundColor Cyan
    Write-Host "       Powered by PDF24 DocTool        " -ForegroundColor Cyan
    Write-Host "  =====================================" -ForegroundColor Cyan

    $rootPath = Get-FolderPathFromUser -Prompt "Enter Records Root Folder path"

    Write-Host ""
    Write-Host "  Root folder set to:" -ForegroundColor Cyan
    Write-Host "    $rootPath" -ForegroundColor White

    while ($true) {
        Write-Host ""
        Write-Host "  =====================================" -ForegroundColor Cyan
        Write-Host "       PDF Record Tool                 " -ForegroundColor Cyan
        Write-Host "  =====================================" -ForegroundColor Cyan
        Write-Host "  1  ->  Merge into existing record    " -ForegroundColor Cyan
        Write-Host "  2  ->  Save as new record            " -ForegroundColor Cyan
        Write-Host "  R  ->  Change root folder            " -ForegroundColor Cyan
        Write-Host "  Q  ->  Quit                          " -ForegroundColor Cyan
        Write-Host "  =====================================" -ForegroundColor Cyan
        Write-Host ""

        $choice = Read-Host "  Enter choice"

        switch ($choice.Trim().ToUpper()) {
            "1"     { Start-MergeWorkflow   -RootPath $rootPath }
            "2"     { Start-NewRecordWorkflow -RootPath $rootPath }
            "R"     {
                $rootPath = Get-FolderPathFromUser -Prompt "Enter new Records Root Folder path"
                Write-Host ""
                Write-Host "  Root folder updated to:" -ForegroundColor Cyan
                Write-Host "    $rootPath" -ForegroundColor White
            }
            "Q"     { Write-Host "  Bye!" -ForegroundColor DarkGray; exit }
            default { Write-Host "  Invalid choice. Enter 1, 2, R, or Q." -ForegroundColor Yellow }
        }
    }
}
