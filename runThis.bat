<# :
@echo off
title PDF Record Tool
cd /d "%~dp0"
powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression $([System.IO.File]::ReadAllText('%~f0'))"
goto :eof
#>
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# PDF Record Tool - GUI Version
# Self-contained: does not dot-source PDFRecordTool.ps1
# Requires PDFRecordTool.ps1 in same folder only for Invoke-PDF24Merge,
# Send-ToRecycleBin, Build-RecordFilename, Build-RecordFolderPath.
# Those are loaded safely below.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# LOAD CORE FUNCTIONS FROM PDFRecordTool.ps1 SAFELY
# ---------------------------------------------------------------------------
$coreScript = Join-Path (Get-Location).Path "PDFRecordTool.ps1"
if (-not (Test-Path $coreScript)) {
    [System.Windows.Forms.MessageBox]::Show(
        "PDFRecordTool.ps1 not found in the same folder.`nPlease keep both files together.",
        "Missing File", "OK", "Error") | Out-Null
    exit
}
$script:PDFRecordToolGUIMode = $true
$coreBytes = [System.IO.File]::ReadAllText($coreScript, [System.Text.Encoding]::UTF8)
$coreSB = [scriptblock]::Create($coreBytes)
. $coreSB

# ---------------------------------------------------------------------------
# THEME - defined after dot-source so they are not overwritten
# ---------------------------------------------------------------------------
$C_BG       = [System.Drawing.ColorTranslator]::FromHtml("#0F172A")
$C_SURFACE  = [System.Drawing.ColorTranslator]::FromHtml("#1E293B")
$C_SURFACE2 = [System.Drawing.ColorTranslator]::FromHtml("#24344F")
$C_ACCENT   = [System.Drawing.ColorTranslator]::FromHtml("#2D7FF9")
$C_TEXT     = [System.Drawing.ColorTranslator]::FromHtml("#E2E8F0")
$C_SUBTEXT  = [System.Drawing.ColorTranslator]::FromHtml("#94A3B8")
$C_BORDER   = [System.Drawing.ColorTranslator]::FromHtml("#334155")
$C_SUCCESS  = [System.Drawing.ColorTranslator]::FromHtml("#22C55E")
$C_WARNING  = [System.Drawing.ColorTranslator]::FromHtml("#F59E0B")
$C_DANGER   = [System.Drawing.ColorTranslator]::FromHtml("#EF4444")
$C_FIELD    = [System.Drawing.ColorTranslator]::FromHtml("#131C2F")
$C_MUTED    = [System.Drawing.ColorTranslator]::FromHtml("#334155")
$C_WHITE    = [System.Drawing.Color]::White

$F_REG   = New-Object System.Drawing.Font("Segoe UI", 9)
$F_SM    = New-Object System.Drawing.Font("Segoe UI", 8)
$F_BOLD  = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$F_TITLE = New-Object System.Drawing.Font("Segoe UI Semibold", 12)

# ---------------------------------------------------------------------------
# WIDGET HELPERS
# ---------------------------------------------------------------------------
function G-Label {
    param([string]$T, [int]$X, [int]$Y, [int]$W=120, [int]$H=20,
          [bool]$Bold=$false, [bool]$Muted=$false, [bool]$Small=$false)
    $l = New-Object System.Windows.Forms.Label
    $l.Text=$T; $l.Left=$X; $l.Top=$Y; $l.Width=$W; $l.Height=$H
    $l.BackColor=[System.Drawing.Color]::Transparent
    $l.ForeColor = if ($Muted) { $C_SUBTEXT } else { $C_TEXT }
    $l.Font = if ($Bold) { $F_BOLD } elseif ($Small) { $F_SM } else { $F_REG }
    $l.AutoSize=$false
    return $l
}
function G-TextBox {
    param([int]$X,[int]$Y,[int]$W,[int]$H=24,[bool]$ReadOnly=$false)
    $t=New-Object System.Windows.Forms.TextBox
    $t.Left=$X;$t.Top=$Y;$t.Width=$W;$t.Height=$H
    $t.BackColor=$C_FIELD;$t.ForeColor=$C_TEXT
    $t.BorderStyle="FixedSingle";$t.Font=$F_REG
    $t.ReadOnly=$ReadOnly
    return $t
}
function G-ComboBox {
    param([int]$X,[int]$Y,[int]$W,[int]$H=24)
    $c=New-Object System.Windows.Forms.ComboBox
    $c.Left=$X;$c.Top=$Y;$c.Width=$W;$c.Height=$H
    $c.BackColor=$C_FIELD;$c.ForeColor=$C_TEXT
    $c.FlatStyle="Flat";$c.Font=$F_REG
    $c.DropDownStyle="DropDown"
    return $c
}
function G-DropDown {
    param([int]$X,[int]$Y,[int]$W,[int]$H=24,[string[]]$Items)
    $c=New-Object System.Windows.Forms.ComboBox
    $c.Left=$X;$c.Top=$Y;$c.Width=$W;$c.Height=$H
    $c.BackColor=$C_FIELD;$c.ForeColor=$C_TEXT
    $c.FlatStyle="Flat";$c.Font=$F_REG
    $c.DropDownStyle="DropDownList"
    foreach ($i in $Items) { [void]$c.Items.Add($i) }
    $c.SelectedIndex=0
    return $c
}
function G-AccentBtn {
    param([string]$T,[int]$X,[int]$Y,[int]$W=120,[int]$H=30)
    $b=New-Object System.Windows.Forms.Button
    $b.Text=$T;$b.Left=$X;$b.Top=$Y;$b.Width=$W;$b.Height=$H
    $b.BackColor=$C_ACCENT;$b.ForeColor=$C_WHITE
    $b.FlatStyle="Flat";$b.Font=$F_BOLD
    $b.FlatAppearance.BorderSize=0
    $b.Cursor=[System.Windows.Forms.Cursors]::Hand
    return $b
}
function G-MutedBtn {
    param([string]$T,[int]$X,[int]$Y,[int]$W=100,[int]$H=30)
    $b=New-Object System.Windows.Forms.Button
    $b.Text=$T;$b.Left=$X;$b.Top=$Y;$b.Width=$W;$b.Height=$H
    $b.BackColor=$C_MUTED;$b.ForeColor=$C_TEXT
    $b.FlatStyle="Flat";$b.Font=$F_REG
    $b.FlatAppearance.BorderSize=0
    $b.Cursor=[System.Windows.Forms.Cursors]::Hand
    return $b
}
function G-Panel {
    param([int]$X,[int]$Y,[int]$W,[int]$H,[string]$BG="BG")
    $p=New-Object System.Windows.Forms.Panel
    $p.Left=$X;$p.Top=$Y;$p.Width=$W;$p.Height=$H
    $p.BackColor = switch($BG){
        "SURFACE"  { $C_SURFACE }
        "SURFACE2" { $C_SURFACE2 }
        default    { $C_BG }
    }
    return $p
}
function G-ListBox {
    param([int]$X,[int]$Y,[int]$W,[int]$H)
    $l=New-Object System.Windows.Forms.ListBox
    $l.Left=$X;$l.Top=$Y;$l.Width=$W;$l.Height=$H
    $l.BackColor=$C_FIELD;$l.ForeColor=$C_TEXT
    $l.BorderStyle="FixedSingle";$l.Font=$F_REG
    return $l
}
function G-Divider {
    param([int]$X,[int]$Y,[int]$W)
    $d=New-Object System.Windows.Forms.Label
    $d.Left=$X;$d.Top=$Y;$d.Width=$W;$d.Height=1
    $d.BackColor=$C_BORDER
    return $d
}
function G-StatusBar {
    param($Form,[string]$Init="Ready")
    $bar=New-Object System.Windows.Forms.StatusStrip
    $bar.BackColor=$C_SURFACE
    $lbl=New-Object System.Windows.Forms.ToolStripStatusLabel
    $lbl.Text=$Init;$lbl.ForeColor=$C_SUBTEXT;$lbl.Font=$F_SM
    [void]$bar.Items.Add($lbl)
    $Form.Controls.Add($bar)
    return $lbl
}
function G-SetStatus {
    param($Lbl,[string]$Text,[string]$Type="info")
    $Lbl.Text=$Text
    $Lbl.ForeColor=switch($Type){
        "success"{$C_SUCCESS}"error"{$C_DANGER}"warning"{$C_WARNING}default{$C_SUBTEXT}
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# ---------------------------------------------------------------------------
# INPUT VALIDATION HELPERS
# ---------------------------------------------------------------------------
$script:InvalidNameChars = '[\\/:*?"<>|]'
$script:MaxYear = (Get-Date).Year + 2

function Test-ValidYear {
    param([string]$Val)
    if($Val -notmatch '^\d{4}$'){return "Must be a 4-digit year (e.g. 2025)."}
    if([int]$Val -gt $script:MaxYear){return "Year cannot exceed $($script:MaxYear)."}
    if([int]$Val -lt 1900){return "Year cannot be before 1900."}
    return $null
}
function Test-ValidName {
    param([string]$Val)
    if(-not $Val){return "Employee name is required."}
    if($Val -match $script:InvalidNameChars){return 'Name contains invalid characters: \ / : * ? " < > |'}
    if($Val.Length -gt 200){return "Name is too long (max 200 characters)."}
    return $null
}
function Limit-YearInput {
    param($TextBox)
    $TextBox.MaxLength=4
    $TextBox.Add_KeyPress({
        param($sender,$e)
        # Allow digits, backspace, delete, and control keys only
        if(-not [char]::IsDigit($e.KeyChar) -and -not [char]::IsControl($e.KeyChar)){
            $e.Handled=$true
            G-SetStatus $script:currentStatusBar "Invalid character - year must be digits only." "warning"
        }
    })
}

# ---------------------------------------------------------------------------
# NAME SUGGESTIONS
# ---------------------------------------------------------------------------
$script:Names = @()
function Load-Names {
    param([string]$Dir)
    $t=Join-Path $Dir "names.txt"
    $c=Join-Path $Dir "names.csv"
    if (Test-Path $t) {
        $script:Names=Get-Content $t -Encoding UTF8|Where-Object{$_.Trim()-ne""}|Sort-Object
        return "names.txt ($($script:Names.Count) names loaded)"
    }
    if (Test-Path $c) {
        $script:Names=Import-Csv $c|ForEach-Object{$_.Name}|Where-Object{$_-ne""}|Sort-Object
        return "names.csv ($($script:Names.Count) names loaded)"
    }
    $script:Names=@()
    return "No names file found (create names.txt in script folder)"
}
function Wire-AC {
    param($CB)
    $CB.AutoCompleteMode="SuggestAppend"
    $CB.AutoCompleteSource="CustomSource"
    $ac=New-Object System.Windows.Forms.AutoCompleteStringCollection
    foreach ($n in $script:Names){[void]$ac.Add($n)}
    $CB.AutoCompleteCustomSource=$ac
    $CB.Add_TextChanged({
        $q=$CB.Text
        if($q.Length -lt 2){return}
        $m=$script:Names|Where-Object{$_ -like "*$q*"}|Select-Object -First 30
        $CB.Items.Clear()
        foreach($x in $m){[void]$CB.Items.Add($x)}
    })
}

# ---------------------------------------------------------------------------
# METADATA FROM FILENAME
# ---------------------------------------------------------------------------
function Get-FileMeta {
    param([string]$Path)
    $b=[System.IO.Path]::GetFileNameWithoutExtension($Path)
    $name=$b -replace '([ _-]+\d{4})+([ _-]*(PT|Dental))?$','' -replace '([ _-]+(PT|Dental))$',''
    $lat="";$old=""
    if($b -match '[ _-]?(\d{4})[ _-]+(\d{4})[ _-]*(PT|Dental)?$'){
        $y1=[int]$Matches[1];$y2=[int]$Matches[2]
        $lat=[string][Math]::Max($y1,$y2);$old=[string][Math]::Min($y1,$y2)
    } elseif($b -match '[ _-]?(\d{4})[ _-]*(PT|Dental)?$'){
        $lat=$Matches[1];$old=$Matches[1]
    }
    $dept="Medical"
    if($b -match '(?i)[ _-]PT$'){$dept="PT"}
    if($b -match '(?i)[ _-]Dental$'){$dept="Dental"}
    $status="ACTIVE"
    $parts=$Path -split [regex]::Escape([System.IO.Path]::DirectorySeparatorChar)
    foreach($p in $parts){
        if($p.ToUpper()-eq"ACTIVE"){$status="ACTIVE";break}
        if($p.ToUpper()-eq"RETIREE"){$status="RETIREE";break}
    }
    return [PSCustomObject]@{Name=$name;Lat=$lat;Old=$old;Dept=$dept;Status=$status}
}
function Get-Preview {
    param([string]$Root,[string]$Name,[string]$Status,[string]$Dept,[string]$Lat,[string]$Old,[string]$Letter="")
    if(-not $Name.Trim()){return ""}
    if(-not $Lat.Trim()){return ""}
    $oldest = if($Old){$Old}else{$Lat}
    $fn=Build-RecordFilename -Name $Name -LatestYear $Lat -OldestYear $oldest -Department $Dept
    $fd=Build-RecordFolderPath -RootPath $Root -Status $Status -Name $Name -Letter $Letter
    return Join-Path $fd $fn
}
function Show-FPicker {
    param([string]$Title,[string]$Init)
    $d=New-Object System.Windows.Forms.OpenFileDialog
    $d.Title=$Title;$d.Filter="PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
    $d.Multiselect=$false
    if($Init -and (Test-Path $Init -PathType Container)){$d.InitialDirectory=$Init}
    if($d.ShowDialog()-eq"OK"){return $d.FileName}
    return $null
}
function Show-MFPicker {
    param([string]$Title,[string]$Init)
    $d=New-Object System.Windows.Forms.OpenFileDialog
    $d.Title=$Title;$d.Filter="PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
    $d.Multiselect=$true
    if($Init -and (Test-Path $Init -PathType Container)){$d.InitialDirectory=$Init}
    if($d.ShowDialog()-eq"OK"){return $d.FileNames}
    return @()
}

# ---------------------------------------------------------------------------
# SIBLING RENAME DIALOG
# ---------------------------------------------------------------------------
function Show-SiblingDlg {
    param([string]$Folder,[string]$CorrectName,[string]$ExcludeFile,[string]$Root)
    $siblings=Get-ChildItem $Folder -File -Filter "*.pdf"|Where-Object{$_.FullName -ne $ExcludeFile}
    if($siblings.Count -eq 0){return}
    $ask=[System.Windows.Forms.MessageBox]::Show(
        "$($siblings.Count) other PDF file(s) found in the moved folder.`nRename them to match '$CorrectName'?",
        "Rename Sibling Files","YesNo","Question")
    if($ask -ne "Yes"){return}

    $dlg=New-Object System.Windows.Forms.Form
    $dlg.Text="Rename Sibling Files";$dlg.Width=520;$dlg.Height=460
    $dlg.BackColor=$C_BG;$dlg.ForeColor=$C_TEXT;$dlg.Font=$F_REG
    $dlg.StartPosition="CenterScreen";$dlg.FormBorderStyle="FixedDialog";$dlg.MaximizeBox=$false

    $dlg.Controls.Add((G-Label "Rename each file to match the corrected employee name." 10 10 490 20 -Bold $true))
    $dlg.Controls.Add((G-Label "Files in folder:" 10 36 490 18 -Bold $true))
    $lb=G-ListBox 10 56 490 80;$dlg.Controls.Add($lb)
    foreach($s in $siblings){[void]$lb.Items.Add($s.Name)}
    $lb.SelectedIndex=0

    $dlg.Controls.Add((G-Label "Employee Name:" 10 148 150 18 -Bold $true))
    $nameBox=G-TextBox 10 168 490;$nameBox.Text=$CorrectName;$dlg.Controls.Add($nameBox)
    $dlg.Controls.Add((G-Label "Latest Year:" 10 200 150 18 -Bold $true))
    $dlg.Controls.Add((G-Label "Oldest Year:" 260 200 150 18 -Bold $true))
    $latBox=G-TextBox 10 220 240;$dlg.Controls.Add($latBox)
    $oldBox=G-TextBox 260 220 240;$dlg.Controls.Add($oldBox)
    $script:currentStatusBar = $msgLbl
    Limit-YearInput $latBox
    Limit-YearInput $oldBox
    $dlg.Controls.Add((G-Label "Department:" 10 252 150 18 -Bold $true))
    $deptDD=G-DropDown 10 270 200 24 @("Medical","PT","Dental");$dlg.Controls.Add($deptDD)

    $msgLbl=G-Label "" 10 305 490 20 -Muted $true -Small $true;$dlg.Controls.Add($msgLbl)
    $renBtn=G-AccentBtn "Rename This File" 10 330 150 30;$dlg.Controls.Add($renBtn)
    $skipBtn=G-MutedBtn "Skip" 168 330 80 30;$dlg.Controls.Add($skipBtn)
    $closeBtn=G-MutedBtn "Done" 400 330 100 30;$dlg.Controls.Add($closeBtn)

    $lb.Add_SelectedIndexChanged({
        $idx=$lb.SelectedIndex
        if($idx -lt 0){return}
        $m=Get-FileMeta $siblings[$idx].FullName
        $nameBox.Text=$CorrectName
        $latBox.Text=$m.Lat;$oldBox.Text=$m.Old
        $deptDD.SelectedItem=$m.Dept
    })
    # fill first
    $m0=Get-FileMeta $siblings[0].FullName
    $latBox.Text=$m0.Lat;$oldBox.Text=$m0.Old;$deptDD.SelectedItem=$m0.Dept

    $renBtn.Add_Click({
        $idx=$lb.SelectedIndex;if($idx -lt 0){return}
        $sib=$siblings[$idx]
        $n=$nameBox.Text.Trim();$l=$latBox.Text.Trim();$o=$oldBox.Text.Trim()
        if(-not $n){$msgLbl.ForeColor=$C_DANGER;$msgLbl.Text="Name required.";return}
        if(-not $l){$msgLbl.ForeColor=$C_DANGER;$msgLbl.Text="Latest year required.";return}
        if(-not $o){$o=$l}
        $fn=Build-RecordFilename -Name $n -LatestYear $l -OldestYear $o -Department $deptDD.SelectedItem
        $dest=Join-Path $Folder $fn
        if(Test-Path $dest){$msgLbl.ForeColor=$C_WARNING;$msgLbl.Text="$fn already exists - skipped."}
        else {
            try{Rename-Item $sib.FullName $fn -Force;$msgLbl.ForeColor=$C_SUCCESS;$msgLbl.Text="Renamed to $fn"}
            catch{$msgLbl.ForeColor=$C_DANGER;$msgLbl.Text=$_.Exception.Message;return}
        }
        if($idx+1 -lt $lb.Items.Count){$lb.SelectedIndex=$idx+1}
        else{$renBtn.Enabled=$false;$skipBtn.Text="Close";$msgLbl.Text+="  All done."}
    })
    $skipBtn.Add_Click({
        if($skipBtn.Text -eq "Close" -or $skipBtn.Text -eq "Done"){$dlg.Close();return}
        $idx=$lb.SelectedIndex
        if($idx+1 -lt $lb.Items.Count){$lb.SelectedIndex=$idx+1}
        else{$dlg.Close()}
    })
    $closeBtn.Add_Click({$dlg.Close()})
    [void]$dlg.ShowDialog()
}

# ---------------------------------------------------------------------------
# MERGE WINDOW
# ---------------------------------------------------------------------------
function Show-MergeWin {
    param([string]$Root, [string]$DefDept="Medical")
    $win=New-Object System.Windows.Forms.Form
    $win.Text="Merge into Existing Record";$win.Width=960;$win.Height=680
    $win.BackColor=$C_BG;$win.ForeColor=$C_TEXT;$win.Font=$F_REG
    $win.StartPosition="CenterScreen";$win.MinimumSize=[System.Drawing.Size]::new(800,580)
    $sb=G-StatusBar $win "Add PDFs to the queue to begin."

    # LEFT - queue panel
    $lp=G-Panel 0 0 240 600 "SURFACE";$win.Controls.Add($lp)
    $lp.Controls.Add((G-Label "Merge Queue" 10 10 200 22 -Bold $true))
    $ql=G-ListBox 10 36 218 460;$lp.Controls.Add($ql)
    $addBtn=G-AccentBtn "Add Files" 10 504 110 26;$lp.Controls.Add($addBtn)
    $remBtn=G-MutedBtn "Remove"    126 504 102 26;$lp.Controls.Add($remBtn)
    $pgLbl=G-Label "0 of 0 processed" 10 534 218 18 -Muted $true -Small $true;$lp.Controls.Add($pgLbl)

    # RIGHT - details panel
    $rp=G-Panel 245 0 700 620 "SURFACE";$win.Controls.Add($rp)

    # New PDF row
    $rp.Controls.Add((G-Label "New PDF (from Downloads / email):" 10 10 450 20 -Bold $true))
    $newLbl=G-Label "(none)" 10 32 540 20 -Muted $true -Small $true;$rp.Controls.Add($newLbl)
    $viewNBtn=G-MutedBtn "View" 590 28 80 26;$viewNBtn.Enabled=$false;$rp.Controls.Add($viewNBtn)
    $rp.Controls.Add((G-Divider 10 62 670))

    # Existing PDF row
    $rp.Controls.Add((G-Label "Existing Employee PDF:" 10 70 300 20 -Bold $true))
    $searchBox=G-TextBox 10 92 488;$searchBox.Text="Type name to search...";$searchBox.ForeColor=$C_SUBTEXT;$rp.Controls.Add($searchBox)
    $searchBox.Add_Enter({if($searchBox.Text-eq"Type name to search..."){$searchBox.Text="";$searchBox.ForeColor=$C_TEXT}})
    $searchBox.Add_Leave({if($searchBox.Text-eq""){$searchBox.Text="Type name to search...";$searchBox.ForeColor=$C_SUBTEXT}})
    $schBtn=G-AccentBtn "Search" 504 90 80 28;$rp.Controls.Add($schBtn)
    $brwBtn=G-MutedBtn "Browse" 590 90 80 28;$rp.Controls.Add($brwBtn)
    $resLB=G-ListBox 10 124 660 90;$rp.Controls.Add($resLB)
    $exLbl=G-Label "(none selected)" 10 220 530 18 -Muted $true -Small $true;$rp.Controls.Add($exLbl)
    $viewEBtn=G-MutedBtn "View Existing" 550 216 120 26;$viewEBtn.Enabled=$false;$rp.Controls.Add($viewEBtn)
    $rp.Controls.Add((G-Divider 10 250 670))

    # Metadata fields
    $y=258
    $hw=320
    $rp.Controls.Add((G-Label "Employee Name:" 10 $y 200 18 -Bold $true))
    $y+=20
    $nameCB=G-ComboBox 10 $y 660;Wire-AC $nameCB;$rp.Controls.Add($nameCB)
    $y+=28

    $rp.Controls.Add((G-Label "Status:" 10 $y 140 18 -Bold $true))
    $rp.Controls.Add((G-Label "Department:" 160 $y 160 18 -Bold $true))
    $rp.Controls.Add((G-Label "Letter:" 330 $y 70 18 -Bold $true))
    $rp.Controls.Add((G-Label "Latest Year:" 410 $y 125 18 -Bold $true))
    $rp.Controls.Add((G-Label "Oldest Year:" 545 $y 125 18 -Bold $true))
    $y+=20
    $stDD=G-DropDown 10 $y 140 24 @("ACTIVE","RETIREE");$rp.Controls.Add($stDD)
    $dtDD=G-DropDown 160 $y 160 24 @("Medical","PT","Dental");$rp.Controls.Add($dtDD)
    if ($DefDept) { $dtDD.SelectedItem = $DefDept }
    $ltDD=G-DropDown 330 $y 70 24 @("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","#");$rp.Controls.Add($ltDD)
    $latBox=G-TextBox 410 $y 125;$rp.Controls.Add($latBox)
    $oldBox=G-TextBox 545 $y 125;$rp.Controls.Add($oldBox)
    $y+=30
    $script:currentStatusBar=$sb
    Limit-YearInput $latBox
    Limit-YearInput $oldBox

    $rp.Controls.Add((G-Label "Output Path Preview:" 10 $y 200 18 -Bold $true))
    $y+=20
    $prevLbl=G-Label "" 10 $y 660 32 -Muted $true -Small $true;$prevLbl.AutoEllipsis=$true;$rp.Controls.Add($prevLbl)
    $y+=36

    $ryclChk=New-Object System.Windows.Forms.CheckBox
    $ryclChk.Text="Move source PDF to Recycle Bin after merge"
    $ryclChk.Left=10;$ryclChk.Top=$y;$ryclChk.Width=400;$ryclChk.Height=20
    $ryclChk.BackColor=$C_SURFACE;$ryclChk.ForeColor=$C_TEXT;$ryclChk.Font=$F_SM;$ryclChk.Checked=$true
    $rp.Controls.Add($ryclChk)
    $y+=28

    $rp.Controls.Add((G-Divider 10 $y 670))
    $y+=8
    $mergeBtn=G-AccentBtn "Merge and Save" 10 $y 160 32;$mergeBtn.Enabled=$false;$rp.Controls.Add($mergeBtn)
    $skipFBtn=G-MutedBtn "Skip File" 178 $y 100 32;$skipFBtn.Enabled=$false;$rp.Controls.Add($skipFBtn)
    $closeBtn=G-MutedBtn "Close" 570 $y 100 32;$rp.Controls.Add($closeBtn)

    # State
    $q=[System.Collections.Generic.List[string]]::new()
    $script:exPdf=""
    $script:qIdx=0
    $script:qDone=0
    $script:qTotal=0
    $script:foundPdfs=@()
    $script:allPdfs=$null

    function UpdatePreview {
        $p=Get-Preview -Root $Root -Name $nameCB.Text.Trim() -Status $stDD.SelectedItem `
            -Dept $dtDD.SelectedItem -Lat $latBox.Text.Trim() -Old $oldBox.Text.Trim() `
            -Letter $ltDD.SelectedItem
        if($p){
            try{$rel=[System.IO.Path]::GetRelativePath($Root,$p)}catch{$rel=$p}
            $prevLbl.Text=$rel;$prevLbl.ForeColor=$C_ACCENT
        } else {
            $prevLbl.Text="Fill in name and years to preview path";$prevLbl.ForeColor=$C_SUBTEXT
        }
    }
    function AutoSetLetter {
        $n=$nameCB.Text.Trim()
        if($n.Length -gt 0){
            $ch=$n[0].ToString().ToUpper()
            if($ch -match '[A-Z]'){$ltDD.SelectedItem=$ch}else{$ltDD.SelectedItem="#"}
        }
    }
    $nameCB.Add_TextChanged({AutoSetLetter;UpdatePreview})
    $stDD.Add_SelectedIndexChanged({UpdatePreview})
    $dtDD.Add_SelectedIndexChanged({UpdatePreview})
    $ltDD.Add_SelectedIndexChanged({UpdatePreview})
    $latBox.Add_TextChanged({UpdatePreview})
    $oldBox.Add_TextChanged({UpdatePreview})

    function UpdateProgress { $pgLbl.Text="$($script:qDone) of $($script:qTotal) processed" }

    function FillMeta { param([string]$Path)
        $m=Get-FileMeta $Path
        $nameCB.Text=$m.Name;$stDD.SelectedItem=$m.Status
        $dtDD.SelectedItem=$m.Dept;$latBox.Text=$m.Lat;$oldBox.Text=$m.Old
        UpdatePreview
    }

    function LoadCurrent {
        if($q.Count -eq 0){$newLbl.Text="(none)";$newLbl.ForeColor=$C_SUBTEXT;$viewNBtn.Enabled=$false;return}
        $idx=$script:qIdx
        if($idx -ge $q.Count){return}
        $path=$q[$idx];$name=Split-Path $path -Leaf
        $newLbl.Text=$name;$newLbl.ForeColor=$C_TEXT;$viewNBtn.Enabled=$true
        $ql.SelectedIndex=$idx
        $script:exPdf=""
        $exLbl.Text="(none selected)";$exLbl.ForeColor=$C_SUBTEXT;$viewEBtn.Enabled=$false
        $mergeBtn.Enabled=$false;$skipFBtn.Enabled=$true
        $resLB.Items.Clear();UpdateProgress
        
        # Clear form fields for the next record
        $nameCB.Text="";$latBox.Text="";$oldBox.Text=""
        $searchBox.Text="Type name to search...";$searchBox.ForeColor=$C_SUBTEXT

        G-SetStatus $sb "Processing: $name"
    }

    function SelectExisting { param([string]$Path)
        if(-not $Path -or -not (Test-Path $Path)){return}
        $script:exPdf=$Path
        try{$rel=[System.IO.Path]::GetRelativePath($Root,$Path)}catch{$rel=$Path}
        $exLbl.Text=$rel;$exLbl.ForeColor=$C_ACCENT;$viewEBtn.Enabled=$true
        $mergeBtn.Enabled=$true
        FillMeta $Path
        G-SetStatus $sb "Existing selected: $(Split-Path $Path -Leaf)"
    }

    $addBtn.Add_Click({
        $dl=Join-Path $env:USERPROFILE "Downloads"
        $files=Show-MFPicker "Select new PDFs to merge" $dl
        $added = 0
        foreach($f in $files){
            if($f.ToLower().EndsWith(".pdf") -and -not $q.Contains($f)){
                $q.Add($f);[void]$ql.Items.Add([System.IO.Path]::GetFileName($f))
                $added++
            }
        }
        $script:qTotal += $added
        if($q.Count -gt 0 -and $script:qIdx -eq 0){LoadCurrent}
        UpdateProgress;G-SetStatus $sb "$($q.Count) file(s) in queue."
    })

    $remBtn.Add_Click({
        $idx=$ql.SelectedIndex;if($idx -lt 0){return}
        $q.RemoveAt($idx);$ql.Items.RemoveAt($idx)
        $script:qTotal--
        $script:qIdx=[Math]::Min($idx,[Math]::Max(0,$q.Count-1))
        LoadCurrent;UpdateProgress
    })

    $ql.Add_SelectedIndexChanged({
        $idx=$ql.SelectedIndex
        if($idx -ge 0 -and $idx -ne $script:qIdx){$script:qIdx=$idx;LoadCurrent}
    })

    $viewNBtn.Add_Click({
        if($script:qIdx -lt $q.Count){try{Start-Process $q[$script:qIdx]}catch{}}
    })
    $viewEBtn.Add_Click({
        if($script:exPdf -and (Test-Path $script:exPdf)){try{Start-Process $script:exPdf}catch{}}
    })

    $doSearch={
        $query=$searchBox.Text.Trim()
        if(-not $query -or $query -eq "Type name to search..."){
            $resLB.Items.Clear()
            return
        }
        if($script:allPdfs -eq $null){
            $resLB.Items.Clear()
            [void]$resLB.Items.Add("Loading files...")
            G-SetStatus $sb "Caching file list for fast search..."
            [System.Windows.Forms.Application]::DoEvents()
            $script:allPdfs=@(Get-ChildItem $Root -Recurse -File -Filter "*.pdf")
            G-SetStatus $sb "Ready."
        }
        $tokens=($query -replace ',', ' ') -split '\s+'|Where-Object{$_ -ne ''}
        $script:foundPdfs=$script:allPdfs|Where-Object{
            $t=($_.Name+' '+$_.Directory.Name).ToLower()
            $ok=$true;foreach($tok in $tokens){if($t -notlike "*$tok*"){$ok=$false;break}};$ok
        }
        $resLB.Items.Clear()
        if($script:foundPdfs){
            foreach($f in $script:foundPdfs|Select-Object -First 50){
                try{$rel=[System.IO.Path]::GetRelativePath($Root,$f.FullName)}catch{$rel=$f.Name}
                [void]$resLB.Items.Add($rel)
            }
            $cnt=$script:foundPdfs.Count
            G-SetStatus $sb "$cnt result(s)$(if($cnt -gt 50){' (showing first 50)'}else{''})"
        } else {
            G-SetStatus $sb "No files found for '$query'." "warning"
        }
    }
    $schBtn.Add_Click($doSearch)
    $searchBox.Add_TextChanged($doSearch)
    $searchBox.Add_KeyDown({param($s,$e)if($e.KeyCode -eq "Return"){& $doSearch}})

    $resLB.Add_SelectedIndexChanged({
        $idx=$resLB.SelectedIndex
        if($idx -ge 0 -and $idx -lt $script:foundPdfs.Count){SelectExisting $script:foundPdfs[$idx].FullName}
    })

    $brwBtn.Add_Click({
        $p=Show-FPicker "Select existing employee PDF" $Root
        if($p){SelectExisting $p}
    })

    $skipFBtn.Add_Click({
        $script:qIdx++;
        if($script:qIdx -ge $q.Count){
            G-SetStatus $sb "Queue complete. $($script:qDone) processed." "success"
            $mergeBtn.Enabled=$false;$skipFBtn.Enabled=$false
        } else {LoadCurrent}
    })

    $mergeBtn.Add_Click({
        $idx=$script:qIdx;if($idx -ge $q.Count){return}
        $newPdf=$q[$idx];$exPdf=$script:exPdf
        $name=$nameCB.Text.Trim();$lat=$latBox.Text.Trim()
        $old=$oldBox.Text.Trim();$st=$stDD.SelectedItem;$dept=$dtDD.SelectedItem;$letter=$ltDD.SelectedItem

        # --- Validation ---
        if(-not $exPdf -or -not (Test-Path $exPdf)){
            G-SetStatus $sb "Select an existing employee PDF first." "error";return}
        $nameErr=Test-ValidName $name
        if($nameErr){G-SetStatus $sb $nameErr "error";$nameCB.Focus();return}
        if(-not $lat){G-SetStatus $sb "Latest year is required." "error";$latBox.Focus();return}
        $latErr=Test-ValidYear $lat
        if($latErr){G-SetStatus $sb "Latest year: $latErr" "error";$latBox.Focus();return}
        if($old){
            $oldErr=Test-ValidYear $old
            if($oldErr){G-SetStatus $sb "Oldest year: $oldErr" "error";$oldBox.Focus();return}
        } else {$old=$lat}

        $exDir=Split-Path $exPdf -Parent

        $outFn=Build-RecordFilename -Name $name -LatestYear $lat -OldestYear $old -Department $dept
        $outDir=Build-RecordFolderPath -RootPath $Root -Status $st -Name $name -Letter $letter
        $outPath=Join-Path $outDir $outFn
        $folderMismatch=[System.IO.Path]::GetFullPath($exDir).ToLower() -ne [System.IO.Path]::GetFullPath($outDir).ToLower()

        $msg="New:      $(Split-Path $newPdf -Leaf)`nExisting: $(Split-Path $exPdf -Leaf)`nOutput:   $outPath"
        if($folderMismatch){$msg+="`n`nFolder will move to:`n$outDir"}
        $ok=[System.Windows.Forms.MessageBox]::Show($msg+"`n`nProceed?","Confirm Merge","YesNo","Question")
        if($ok -ne "Yes"){return}

        $mergeBtn.Enabled=$false;G-SetStatus $sb "Merging via PDF24..."
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $movedEx=$exPdf
            if($folderMismatch){
                $par=Split-Path $outDir -Parent
                if(-not(Test-Path $par)){New-Item -ItemType Directory $par -Force|Out-Null}
                if(Test-Path $outDir){
                    Get-ChildItem $exDir -Recurse|ForEach-Object{
                        $dst=Join-Path $outDir $_.FullName.Substring($exDir.Length).TrimStart("\")
                        $dd=Split-Path $dst -Parent
                        if(-not(Test-Path $dd)){New-Item -ItemType Directory $dd -Force|Out-Null}
                        if(-not $_.PSIsContainer){Move-Item $_.FullName $dst -Force}
                    }
                    if(-not(Get-ChildItem $exDir -Recurse -File)){Remove-Item $exDir -Recurse -Force -EA SilentlyContinue}
                } else {Move-Item $exDir $outDir -Force}
                $movedEx=Join-Path $outDir (Split-Path $exPdf -Leaf)
            }
            if(-not(Test-Path $outDir)){New-Item -ItemType Directory $outDir -Force|Out-Null}

            $tmp=$outPath+".merging.tmp"
            $actual=Invoke-PDF24Merge -NewPdfPath $newPdf -ExistingPdfPath $movedEx -OutputPath $tmp

            $exN=[System.IO.Path]::GetFullPath($movedEx).ToLower()
            $outN=[System.IO.Path]::GetFullPath($outPath).ToLower()
            if($exN -eq $outN){Send-ToRecycleBin $movedEx}
            elseif(Test-Path $outPath){Remove-Item $outPath -Force}
            Move-Item $actual $outPath -Force
            if($exN -ne $outN -and (Test-Path $movedEx)){Send-ToRecycleBin $movedEx}
            if($ryclChk.Checked -and (Test-Path $newPdf)){Send-ToRecycleBin $newPdf}

            if($folderMismatch){
                Show-SiblingDlg $outDir $name $outPath $Root
            }

            $script:qDone++;UpdateProgress
            G-SetStatus $sb "Merged: $(Split-Path $outPath -Leaf)" "success"
            $script:allPdfs=$null

            # Remove the processed item from the visible queue list
            if($script:qIdx -lt $ql.Items.Count){
                $ql.Items.RemoveAt($script:qIdx)
                $q.RemoveAt($script:qIdx)
                # Don't increment qIdx -- after removal the next item
                # slides into the same index position
            }

            if($q.Count -eq 0){
                G-SetStatus $sb "Queue complete. $($script:qDone) file(s) merged." "success"
                $win.Close()
            } else {
                # Cap index in case we just removed the last item
                if($script:qIdx -ge $q.Count){$script:qIdx=$q.Count-1}
                LoadCurrent
            }
        } catch {
            $mergeBtn.Enabled=$true
            Write-ErrorLog "Merge operation failed" $_.Exception
            G-SetStatus $sb "Error: $($_.Exception.Message)" "error"
            [System.Windows.Forms.MessageBox]::Show("Merge failed:`n$($_.Exception.Message)","Error","OK","Error")|Out-Null
        }
    })

    $closeBtn.Add_Click({$win.Close()})
    UpdatePreview
    [void]$win.ShowDialog()
}

# ---------------------------------------------------------------------------
# NEW RECORD WINDOW
# ---------------------------------------------------------------------------
function Show-NewRecordWin {
    param([string]$Root, [string]$DefDept="Medical")
    $win=New-Object System.Windows.Forms.Form
    $win.Text="Save as New Record";$win.Width=640;$win.Height=660
    $win.BackColor=$C_BG;$win.ForeColor=$C_TEXT;$win.Font=$F_REG
    $win.StartPosition="CenterScreen";$win.FormBorderStyle="FixedDialog";$win.MaximizeBox=$false
    $sb=G-StatusBar $win "Add PDFs to queue and fill in details."

    $win.Controls.Add((G-Label "New Record Queue" 10 10 580 20 -Bold $true))
    $ql=G-ListBox 10 32 604 110;$win.Controls.Add($ql)
    $addBtn=G-AccentBtn "Add Files" 10 148 120 28;$win.Controls.Add($addBtn)
    $remBtn=G-MutedBtn "Remove"    138 148 100 28;$win.Controls.Add($remBtn)
    $pgLbl=G-Label "0 of 0 saved" 454 152 160 20 -Muted $true -Small $true;$win.Controls.Add($pgLbl)
    $win.Controls.Add((G-Divider 10 182 604))

    $win.Controls.Add((G-Label "Current:" 10 190 70 18 -Bold $true))
    $curLbl=G-Label "(none)" 85 190 440 18 -Muted $true;$win.Controls.Add($curLbl)
    $viewBtn=G-MutedBtn "View" 534 186 80 26;$viewBtn.Enabled=$false;$win.Controls.Add($viewBtn)
    $win.Controls.Add((G-Divider 10 220 604))

    $y=228
    $win.Controls.Add((G-Label "Employee Name:" 10 $y 200 18 -Bold $true));$y+=20
    $nameCB=G-ComboBox 10 $y 604;Wire-AC $nameCB;$win.Controls.Add($nameCB);$y+=28

    $win.Controls.Add((G-Label "Status:" 10 $y 140 18 -Bold $true))
    $win.Controls.Add((G-Label "Department:" 160 $y 140 18 -Bold $true))
    $win.Controls.Add((G-Label "Letter:" 310 $y 64 18 -Bold $true))
    $win.Controls.Add((G-Label "Latest Year:" 384 $y 110 18 -Bold $true))
    $win.Controls.Add((G-Label "Oldest Year:" 504 $y 110 18 -Bold $true));$y+=20
    $stDD=G-DropDown 10 $y 140 24 @("ACTIVE","RETIREE");$win.Controls.Add($stDD)
    $dtDD=G-DropDown 160 $y 140 24 @("Medical","PT","Dental");$win.Controls.Add($dtDD)
    if ($DefDept) { $dtDD.SelectedItem = $DefDept }
    $ltDD=G-DropDown 310 $y 64 24 @("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","#");$win.Controls.Add($ltDD)
    $latBox=G-TextBox 384 $y 110;$win.Controls.Add($latBox)
    $oldBox=G-TextBox 504 $y 110;$win.Controls.Add($oldBox);$y+=30
    $script:currentStatusBar=$sb
    Limit-YearInput $latBox
    Limit-YearInput $oldBox

    $win.Controls.Add((G-Label "Output Path Preview:" 10 $y 200 18 -Bold $true));$y+=20
    $prevLbl=G-Label "" 10 $y 604 32 -Muted $true -Small $true;$prevLbl.AutoEllipsis=$true;$win.Controls.Add($prevLbl);$y+=38

    $ryclChk=New-Object System.Windows.Forms.CheckBox
    $ryclChk.Text="Move source PDF to Recycle Bin after save"
    $ryclChk.Left=10;$ryclChk.Top=$y;$ryclChk.Width=400;$ryclChk.Height=20
    $ryclChk.BackColor=$C_BG;$ryclChk.ForeColor=$C_TEXT;$ryclChk.Font=$F_SM;$ryclChk.Checked=$true
    $win.Controls.Add($ryclChk);$y+=30

    $win.Controls.Add((G-Divider 10 $y 604));$y+=8
    $saveBtn=G-AccentBtn "Save Record" 10 $y 140 32;$saveBtn.Enabled=$false;$win.Controls.Add($saveBtn)
    $skipBtn=G-MutedBtn "Skip" 158 $y 80 32;$skipBtn.Enabled=$false;$win.Controls.Add($skipBtn)
    $closeBtn=G-MutedBtn "Close" 514 $y 100 32;$win.Controls.Add($closeBtn)

    $q=[System.Collections.Generic.List[string]]::new()
    $script:nIdx=0;$script:nDone=0;$script:nTotal=0

    function UpdateNPrev {
        $p=Get-Preview -Root $Root -Name $nameCB.Text.Trim() -Status $stDD.SelectedItem `
            -Dept $dtDD.SelectedItem -Lat $latBox.Text.Trim() -Old $oldBox.Text.Trim() `
            -Letter $ltDD.SelectedItem
        if($p){
            try{$rel=[System.IO.Path]::GetRelativePath($Root,$p)}catch{$rel=$p}
            $prevLbl.Text=$rel;$prevLbl.ForeColor=$C_ACCENT
        } else {$prevLbl.Text="Fill in name and years to preview.";$prevLbl.ForeColor=$C_SUBTEXT}
    }
    function AutoSetLetterN {
        $n=$nameCB.Text.Trim()
        if($n.Length -gt 0){
            $ch=$n[0].ToString().ToUpper()
            if($ch -match '[A-Z]'){$ltDD.SelectedItem=$ch}else{$ltDD.SelectedItem="#"}
        }
    }
    $nameCB.Add_TextChanged({AutoSetLetterN;UpdateNPrev})
    $stDD.Add_SelectedIndexChanged({UpdateNPrev})
    $dtDD.Add_SelectedIndexChanged({UpdateNPrev})
    $ltDD.Add_SelectedIndexChanged({UpdateNPrev})
    $latBox.Add_TextChanged({UpdateNPrev})
    $oldBox.Add_TextChanged({UpdateNPrev})

    function UpdateNProg {$pgLbl.Text="$($script:nDone) of $($script:nTotal) saved"}

    function LoadNCurrent {
        if($q.Count -eq 0){$curLbl.Text="(none)";$curLbl.ForeColor=$C_SUBTEXT;$viewBtn.Enabled=$false
            $saveBtn.Enabled=$false;$skipBtn.Enabled=$false;return}
        $idx=$script:nIdx;if($idx -ge $q.Count){return}
        $p=$q[$idx];$n=Split-Path $p -Leaf
        $curLbl.Text=$n;$curLbl.ForeColor=$C_TEXT;$viewBtn.Enabled=$true
        $saveBtn.Enabled=$true;$skipBtn.Enabled=$true
        $ql.SelectedIndex=$idx;UpdateNProg

        # Clear form fields for the next record
        $nameCB.Text="";$latBox.Text="";$oldBox.Text=""

        G-SetStatus $sb "Current: $n"
    }

    $addBtn.Add_Click({
        $dl=Join-Path $env:USERPROFILE "Downloads"
        $files=Show-MFPicker "Select new PDFs to save as new records" $dl
        $added = 0
        foreach($f in $files){
            if($f.ToLower().EndsWith(".pdf") -and -not $q.Contains($f)){
                $q.Add($f);[void]$ql.Items.Add([System.IO.Path]::GetFileName($f))
                $added++
            }
        }
        $script:nTotal += $added
        LoadNCurrent;G-SetStatus $sb "$($q.Count) file(s) queued."
    })
    $remBtn.Add_Click({
        $idx=$ql.SelectedIndex;if($idx -lt 0){return}
        $q.RemoveAt($idx);$ql.Items.RemoveAt($idx)
        $script:nTotal--
        $script:nIdx=[Math]::Min($idx,[Math]::Max(0,$q.Count-1))
        LoadNCurrent;UpdateNProg
    })
    $ql.Add_SelectedIndexChanged({
        $idx=$ql.SelectedIndex
        if($idx -ge 0 -and $idx -ne $script:nIdx){$script:nIdx=$idx;LoadNCurrent}
    })
    $viewBtn.Add_Click({
        if($script:nIdx -lt $q.Count){try{Start-Process $q[$script:nIdx]}catch{}}
    })
    $skipBtn.Add_Click({
        $script:nIdx++;
        if($script:nIdx -ge $q.Count){
            G-SetStatus $sb "Queue complete. $($script:nDone) saved." "success"
            $saveBtn.Enabled=$false;$skipBtn.Enabled=$false
        } else {LoadNCurrent}
    })
    $saveBtn.Add_Click({
        $idx=$script:nIdx;if($idx -ge $q.Count){return}
        $src=$q[$idx];$name=$nameCB.Text.Trim();$lat=$latBox.Text.Trim()
        $old=$oldBox.Text.Trim();$st=$stDD.SelectedItem;$dept=$dtDD.SelectedItem

        # --- Validation ---
        $nameErr=Test-ValidName $name
        if($nameErr){G-SetStatus $sb $nameErr "error";$nameCB.Focus();return}
        if(-not $lat){G-SetStatus $sb "Latest year is required." "error";$latBox.Focus();return}
        $latErr=Test-ValidYear $lat
        if($latErr){G-SetStatus $sb "Latest year: $latErr" "error";$latBox.Focus();return}
        if($old){
            $oldErr=Test-ValidYear $old
            if($oldErr){G-SetStatus $sb "Oldest year: $oldErr" "error";$oldBox.Focus();return}
        } else {$old=$lat}
        $fn=Build-RecordFilename -Name $name -LatestYear $lat -OldestYear $old -Department $dept
        $fd=Build-RecordFolderPath -RootPath $Root -Status $st -Name $name -Letter $ltDD.SelectedItem
        $outPath=Join-Path $fd $fn
        $msg="Source:      $(Split-Path $src -Leaf)`nDestination: $outPath"
        if(Test-Path $fd){
            $ex=(Get-ChildItem $fd -File).Name -join ", "
            $msg+="`n`nFolder exists. Current files: $ex"
        }
        $ok=[System.Windows.Forms.MessageBox]::Show($msg+"`n`nProceed?","Confirm Save","YesNo","Question")
        if($ok -ne "Yes"){return}
        try{
            if(-not(Test-Path $fd)){New-Item -ItemType Directory $fd -Force|Out-Null}
            if(Test-Path $outPath){
                $ow=[System.Windows.Forms.MessageBox]::Show("$fn exists. Overwrite?","File Exists","YesNo","Warning")
                if($ow -ne "Yes"){return}
            }
            Copy-Item $src $outPath -Force
            if($ryclChk.Checked -and (Test-Path $src)){Send-ToRecycleBin $src}
            $script:nDone++;UpdateNProg
            G-SetStatus $sb "Saved: $fn" "success"

            # Remove the processed item from the visible queue list
            if($script:nIdx -lt $ql.Items.Count){
                $ql.Items.RemoveAt($script:nIdx)
                $q.RemoveAt($script:nIdx)
                # Don't increment nIdx -- next item slides into same position
            }

            if($q.Count -eq 0){
                G-SetStatus $sb "Queue complete. $($script:nDone) file(s) saved." "success"
                $win.Close()
            } else {
                if($script:nIdx -ge $q.Count){$script:nIdx=$q.Count-1}
                LoadNCurrent
            }
        } catch {
            Write-ErrorLog "Save New Record failed" $_.Exception
            G-SetStatus $sb "Error: $($_.Exception.Message)" "error"
            [System.Windows.Forms.MessageBox]::Show("Save failed:`n$($_.Exception.Message)","Error","OK","Error")|Out-Null
        }
    })
    $closeBtn.Add_Click({$win.Close()})
    UpdateNPrev
    [void]$win.ShowDialog()
}

# ---------------------------------------------------------------------------
# EDIT RECORD WINDOW
# ---------------------------------------------------------------------------
function Show-EditRecordWin {
    param([string]$Root, [string]$DefDept="Medical")
    $win=New-Object System.Windows.Forms.Form
    $win.Text="Edit Record";$win.Width=640;$win.Height=520
    $win.BackColor=$C_BG;$win.ForeColor=$C_TEXT;$win.Font=$F_REG
    $win.StartPosition="CenterScreen";$win.FormBorderStyle="FixedSingle";$win.MaximizeBox=$false
    $sb=G-StatusBar $win "Select a record to edit."

    # Top Section: Select Existing
    $win.Controls.Add((G-Label "Select Existing Employee PDF:" 10 10 300 20 -Bold $true))
    $searchBox=G-TextBox 10 32 430;$searchBox.Text="Type name to search...";$searchBox.ForeColor=$C_SUBTEXT;$win.Controls.Add($searchBox)
    $searchBox.Add_Enter({if($searchBox.Text-eq"Type name to search..."){$searchBox.Text="";$searchBox.ForeColor=$C_TEXT}})
    $searchBox.Add_Leave({if($searchBox.Text-eq""){$searchBox.Text="Type name to search...";$searchBox.ForeColor=$C_SUBTEXT}})
    $schBtn=G-AccentBtn "Search" 448 30 80 28;$win.Controls.Add($schBtn)
    $brwBtn=G-MutedBtn "Browse" 534 30 80 28;$win.Controls.Add($brwBtn)
    $resLB=G-ListBox 10 64 604 90;$win.Controls.Add($resLB)
    $exLbl=G-Label "(none selected)" 10 160 514 18 -Muted $true -Small $true;$exLbl.AutoEllipsis=$true;$win.Controls.Add($exLbl)
    $viewEBtn=G-MutedBtn "View PDF" 534 156 80 26;$viewEBtn.Enabled=$false;$win.Controls.Add($viewEBtn)
    $win.Controls.Add((G-Divider 10 190 604))

    # Middle Section: Metadata fields
    $y=200
    $win.Controls.Add((G-Label "Employee Name:" 10 $y 200 18 -Bold $true));$y+=20
    $nameCB=G-ComboBox 10 $y 604;Wire-AC $nameCB;$win.Controls.Add($nameCB);$y+=28

    $win.Controls.Add((G-Label "Status:" 10 $y 140 18 -Bold $true))
    $win.Controls.Add((G-Label "Department:" 160 $y 140 18 -Bold $true))
    $win.Controls.Add((G-Label "Letter:" 310 $y 64 18 -Bold $true))
    $win.Controls.Add((G-Label "Latest Year:" 384 $y 110 18 -Bold $true))
    $win.Controls.Add((G-Label "Oldest Year:" 504 $y 110 18 -Bold $true));$y+=20
    $stDD=G-DropDown 10 $y 140 24 @("ACTIVE","RETIREE");$win.Controls.Add($stDD)
    $dtDD=G-DropDown 160 $y 140 24 @("Medical","PT","Dental");$win.Controls.Add($dtDD)
    if ($DefDept) { $dtDD.SelectedItem = $DefDept }
    $ltDD=G-DropDown 310 $y 64 24 @("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","#");$win.Controls.Add($ltDD)
    $latBox=G-TextBox 384 $y 110;$win.Controls.Add($latBox)
    $oldBox=G-TextBox 504 $y 110;$win.Controls.Add($oldBox);$y+=30
    $script:currentStatusBar=$sb
    Limit-YearInput $latBox
    Limit-YearInput $oldBox

    $win.Controls.Add((G-Label "Output Path Preview:" 10 $y 200 18 -Bold $true));$y+=20
    $prevLbl=G-Label "" 10 $y 604 32 -Muted $true -Small $true;$prevLbl.AutoEllipsis=$true;$win.Controls.Add($prevLbl);$y+=38

    $win.Controls.Add((G-Divider 10 $y 604));$y+=8
    $saveBtn=G-AccentBtn "Save Changes" 10 $y 140 32;$saveBtn.Enabled=$false;$win.Controls.Add($saveBtn)
    $closeBtn=G-MutedBtn "Cancel" 514 $y 100 32;$win.Controls.Add($closeBtn)

    $script:editPdf=""
    $script:foundPdfs=@()
    $script:allPdfs=$null

    function UpdateEPrev {
        $p=Get-Preview -Root $Root -Name $nameCB.Text.Trim() -Status $stDD.SelectedItem `
            -Dept $dtDD.SelectedItem -Lat $latBox.Text.Trim() -Old $oldBox.Text.Trim() `
            -Letter $ltDD.SelectedItem
        if($p){
            try{$rel=[System.IO.Path]::GetRelativePath($Root,$p)}catch{$rel=$p}
            $prevLbl.Text=$rel;$prevLbl.ForeColor=$C_ACCENT
        } else {$prevLbl.Text="Fill in name and years to preview.";$prevLbl.ForeColor=$C_SUBTEXT}
    }
    function AutoSetLetterE {
        $n=$nameCB.Text.Trim()
        if($n.Length -gt 0){
            $ch=$n[0].ToString().ToUpper()
            if($ch -match '[A-Z]'){$ltDD.SelectedItem=$ch}else{$ltDD.SelectedItem="#"}
        }
    }
    $nameCB.Add_TextChanged({AutoSetLetterE;UpdateEPrev})
    $stDD.Add_SelectedIndexChanged({UpdateEPrev})
    $dtDD.Add_SelectedIndexChanged({UpdateEPrev})
    $ltDD.Add_SelectedIndexChanged({UpdateEPrev})
    $latBox.Add_TextChanged({UpdateEPrev})
    $oldBox.Add_TextChanged({UpdateEPrev})

    function SelectExistingE {
        param([string]$Path)
        $script:editPdf=$Path
        try{$rel=[System.IO.Path]::GetRelativePath($Root,$Path)}catch{$rel=$Path}
        $exLbl.Text=$rel;$viewEBtn.Enabled=$true;$saveBtn.Enabled=$true
        $m=Get-FileMeta $Path
        $nameCB.Text=$m.Name;$stDD.SelectedItem=$m.Status
        $dtDD.SelectedItem=$m.Dept;$latBox.Text=$m.Lat;$oldBox.Text=$m.Old
        # Auto-detect letter from existing path (parent folder name)
        $pDir=Split-Path (Split-Path $Path -Parent) -Leaf
        if($pDir.Length -eq 1){$ltDD.SelectedItem=$pDir.ToUpper()}
        UpdateEPrev
        G-SetStatus $sb "Loaded metadata from selected file."
    }
    $viewEBtn.Add_Click({try{Start-Process $script:editPdf}catch{}})

    $doSearch={
        $query=$searchBox.Text.Trim()
        if(-not $query -or $query -eq "Type name to search..."){
            $resLB.Items.Clear()
            return
        }
        if($script:allPdfs -eq $null){
            $resLB.Items.Clear()
            [void]$resLB.Items.Add("Loading files...")
            G-SetStatus $sb "Caching file list for fast search..."
            [System.Windows.Forms.Application]::DoEvents()
            $script:allPdfs=@(Get-ChildItem $Root -Recurse -File -Filter "*.pdf")
            G-SetStatus $sb "Ready."
        }
        $tokens=($query -replace ',', ' ') -split '\s+'|Where-Object{$_ -ne ''}
        $script:foundPdfs=$script:allPdfs|Where-Object{
            $t=($_.Name+' '+$_.Directory.Name).ToLower()
            $ok=$true;foreach($tok in $tokens){if($t -notlike "*$tok*"){$ok=$false;break}};$ok
        }
        $resLB.Items.Clear()
        if($script:foundPdfs){
            foreach($f in $script:foundPdfs|Select-Object -First 50){
                try{$rel=[System.IO.Path]::GetRelativePath($Root,$f.FullName)}catch{$rel=$f.Name}
                [void]$resLB.Items.Add($rel)
            }
            $cnt=$script:foundPdfs.Count
            G-SetStatus $sb "$cnt result(s)$(if($cnt -gt 50){' (showing first 50)'}else{''})"
        } else {
            G-SetStatus $sb "No files found for '$query'." "warning"
        }
    }
    $schBtn.Add_Click($doSearch)
    $searchBox.Add_TextChanged($doSearch)
    $searchBox.Add_KeyDown({param($s,$e)if($e.KeyCode -eq "Return"){& $doSearch}})

    $resLB.Add_SelectedIndexChanged({
        $idx=$resLB.SelectedIndex
        if($idx -ge 0 -and $idx -lt $script:foundPdfs.Count){SelectExistingE $script:foundPdfs[$idx].FullName}
    })
    $brwBtn.Add_Click({
        $p=Show-FPicker "Select existing employee PDF to edit" $Root
        if($p){SelectExistingE $p}
    })

    $saveBtn.Add_Click({
        if(-not $script:editPdf -or -not (Test-Path $script:editPdf)){
            G-SetStatus $sb "Select an existing employee PDF first." "error";return}
        
        $name=$nameCB.Text.Trim();$lat=$latBox.Text.Trim()
        $old=$oldBox.Text.Trim();$st=$stDD.SelectedItem;$dept=$dtDD.SelectedItem;$letter=$ltDD.SelectedItem

        # --- Validation ---
        $nameErr=Test-ValidName $name
        if($nameErr){G-SetStatus $sb $nameErr "error";$nameCB.Focus();return}
        if(-not $lat){G-SetStatus $sb "Latest year is required." "error";$latBox.Focus();return}
        $latErr=Test-ValidYear $lat
        if($latErr){G-SetStatus $sb "Latest year: $latErr" "error";$latBox.Focus();return}
        if($old){
            $oldErr=Test-ValidYear $old
            if($oldErr){G-SetStatus $sb "Oldest year: $oldErr" "error";$oldBox.Focus();return}
        } else {$old=$lat}

        $exPdf=$script:editPdf
        $exDir=Split-Path $exPdf -Parent
        $outFn=Build-RecordFilename -Name $name -LatestYear $lat -OldestYear $old -Department $dept
        $outDir=Build-RecordFolderPath -RootPath $Root -Status $st -Name $name -Letter $letter
        $outPath=Join-Path $outDir $outFn

        $folderMismatch=[System.IO.Path]::GetFullPath($exDir).ToLower() -ne [System.IO.Path]::GetFullPath($outDir).ToLower()
        $fileMismatch=[System.IO.Path]::GetFullPath($exPdf).ToLower() -ne [System.IO.Path]::GetFullPath($outPath).ToLower()

        if(-not $folderMismatch -and -not $fileMismatch){
            G-SetStatus $sb "No changes made." "warning";return
        }

        $msg="Old: $(Split-Path $exPdf -Leaf)`nNew: $outFn"
        if($folderMismatch){$msg+="`n`nFolder will move to:`n$outDir"}
        $ok=[System.Windows.Forms.MessageBox]::Show($msg+"`n`nProceed?","Confirm Changes","YesNo","Question")
        if($ok -ne "Yes"){return}

        try {
            $movedEx=$exPdf
            if($folderMismatch){
                $par=Split-Path $outDir -Parent
                if(-not(Test-Path $par)){New-Item -ItemType Directory $par -Force|Out-Null}
                if(Test-Path $outDir){
                    Get-ChildItem $exDir -Recurse|ForEach-Object{
                        $dst=Join-Path $outDir $_.FullName.Substring($exDir.Length).TrimStart("\")
                        $dd=Split-Path $dst -Parent
                        if(-not(Test-Path $dd)){New-Item -ItemType Directory $dd -Force|Out-Null}
                        if(-not $_.PSIsContainer){Move-Item $_.FullName $dst -Force}
                    }
                    if(-not(Get-ChildItem $exDir -Recurse -File)){Remove-Item $exDir -Recurse -Force -EA SilentlyContinue}
                } else {Move-Item $exDir $outDir -Force}
                $movedEx=Join-Path $outDir (Split-Path $exPdf -Leaf)
            }
            if(-not(Test-Path $outDir)){New-Item -ItemType Directory $outDir -Force|Out-Null}
            
            # File rename (if filename actually changed, or if we moved and outputpath is different)
            if([System.IO.Path]::GetFullPath($movedEx).ToLower() -ne [System.IO.Path]::GetFullPath($outPath).ToLower()){
                if(Test-Path $outPath){
                    $ow=[System.Windows.Forms.MessageBox]::Show("$outFn exists. Overwrite?","File Exists","YesNo","Warning")
                    if($ow -eq "Yes"){
                        Remove-Item $outPath -Force
                        Rename-Item $movedEx $outFn -Force
                    }
                } else {
                    Rename-Item $movedEx $outFn -Force
                }
            }

            if($folderMismatch){
                Show-SiblingDlg $outDir $name $outPath $Root
            }

            [System.Windows.Forms.MessageBox]::Show("Record updated successfully.","Success","OK","Information")|Out-Null
            $script:allPdfs=$null
            $win.Close()
        } catch {
            Write-ErrorLog "Edit Record failed" $_.Exception
            G-SetStatus $sb "Error: $($_.Exception.Message)" "error"
            [System.Windows.Forms.MessageBox]::Show("Update failed:`n$($_.Exception.Message)","Error","OK","Error")|Out-Null
        }
    })
    $closeBtn.Add_Click({$win.Close()})
    UpdateEPrev
    [void]$win.ShowDialog()
}

# ---------------------------------------------------------------------------
# MAIN WINDOW
# ---------------------------------------------------------------------------
$scriptDir=(Get-Location).Path
if(-not $scriptDir){$scriptDir=Split-Path $MyInvocation.MyCommand.Path -Parent}

$main=New-Object System.Windows.Forms.Form
$main.Text="PDF Record Tool";$main.Width=730;$main.Height=310
$main.BackColor=$C_BG;$main.ForeColor=$C_TEXT;$main.Font=$F_REG
$main.StartPosition="CenterScreen";$main.FormBorderStyle="FixedSingle";$main.MaximizeBox=$false

$hp=G-Panel 0 0 730 58 "SURFACE";$main.Controls.Add($hp)
$tl=G-Label "PDF Record Tool" 14 10 400 26 -Bold $true;$tl.Font=$F_TITLE;$hp.Controls.Add($tl)
$sl=G-Label "Powered by PDF24 DocTool" 14 34 300 18 -Muted $true -Small $true;$hp.Controls.Add($sl)

$main.Controls.Add((G-Label "Records Root Folder:" 14 72 160 18 -Bold $true))
$rootBox=G-TextBox 14 92 592 26 $true;$main.Controls.Add($rootBox)
$rootBtn=G-MutedBtn "Browse" 612 90 82 28;$main.Controls.Add($rootBtn)

$main.Controls.Add((G-Label "Name Suggestions:" 14 128 160 18 -Bold $true))
$nameSugLbl=G-Label "Checking..." 14 148 450 18 -Muted $true -Small $true;$main.Controls.Add($nameSugLbl)

$main.Controls.Add((G-Label "Default Department:" 554 128 140 18 -Bold $true))
$defDeptDD=G-DropDown 554 146 140 24 @("Medical","PT","Dental");$main.Controls.Add($defDeptDD)

$main.Controls.Add((G-Divider 14 172 680))

$mergeMainBtn=G-AccentBtn "Merge into Existing Record" 14 182 220 40;$mergeMainBtn.Font=$F_BOLD;$main.Controls.Add($mergeMainBtn)
$newMainBtn=G-AccentBtn "Save as New Record" 244 182 220 40;$newMainBtn.Font=$F_BOLD
$newMainBtn.BackColor=[System.Drawing.ColorTranslator]::FromHtml("#1E40AF")
$main.Controls.Add($newMainBtn)
$editMainBtn=G-AccentBtn "Edit Record" 474 182 220 40;$editMainBtn.Font=$F_BOLD
$editMainBtn.BackColor=[System.Drawing.ColorTranslator]::FromHtml("#059669") # distinct color
$main.Controls.Add($editMainBtn)

$mainSb=G-StatusBar $main "Ready"

# Load names
$nameSugLbl.Text=Load-Names $scriptDir

# Load saved settings
$settingsFile=Join-Path $scriptDir "pdftool_settings.json"
if(Test-Path $settingsFile){
    try{
        $sv=Get-Content $settingsFile -Raw|ConvertFrom-Json
        if($sv.RootPath -and (Test-Path $sv.RootPath)){
            $rootBox.Text=$sv.RootPath
            G-SetStatus $mainSb "Settings loaded from last session."
        }
        if($sv.DefaultDept) { $defDeptDD.SelectedItem=$sv.DefaultDept }
    }catch{}
}

function SaveSettings {
    try{@{RootPath=$rootBox.Text;DefaultDept=$defDeptDD.SelectedItem}|ConvertTo-Json|Set-Content $settingsFile -Encoding UTF8}catch{}
}

$defDeptDD.Add_SelectedIndexChanged({ SaveSettings })

$rootBtn.Add_Click({
    $d=New-Object System.Windows.Forms.FolderBrowserDialog
    $d.Description="Select the Records Root Folder"
    if($rootBox.Text -and (Test-Path $rootBox.Text)){$d.SelectedPath=$rootBox.Text}
    if($d.ShowDialog()-eq"OK"){
        $rootBox.Text=$d.SelectedPath
        SaveSettings
        G-SetStatus $mainSb "Root folder set."
    }
})

$mergeMainBtn.Add_Click({
    if(-not $rootBox.Text -or -not(Test-Path $rootBox.Text)){
        [System.Windows.Forms.MessageBox]::Show("Set the Records Root Folder first.","No Root Folder","OK","Warning")|Out-Null;return
    }
    Show-MergeWin $rootBox.Text $defDeptDD.SelectedItem
})

$newMainBtn.Add_Click({
    if(-not $rootBox.Text -or -not(Test-Path $rootBox.Text)){
        [System.Windows.Forms.MessageBox]::Show("Set the Records Root Folder first.","No Root Folder","OK","Warning")|Out-Null;return
    }
    Show-NewRecordWin $rootBox.Text $defDeptDD.SelectedItem
})

$editMainBtn.Add_Click({
    if(-not $rootBox.Text -or -not(Test-Path $rootBox.Text)){
        [System.Windows.Forms.MessageBox]::Show("Set the Records Root Folder first.","No Root Folder","OK","Warning")|Out-Null;return
    }
    Show-EditRecordWin $rootBox.Text $defDeptDD.SelectedItem
})

$main.Add_FormClosing({ SaveSettings })

[void]$main.ShowDialog()
