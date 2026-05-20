# InventoryTool.ps1
# Serial number + storage location tracker with Excel backend

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ─── Paths ────────────────────────────────────────────────────────────────────
$script:ExcelPath = Join-Path $PSScriptRoot "InventoryData.xlsx"
$script:Inventory = [System.Collections.Generic.Dictionary[string,string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

# ─── Excel helpers ────────────────────────────────────────────────────────────
function New-Excel {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible         = $false
    $xl.DisplayAlerts   = $false
    $xl.ScreenUpdating  = $false
    return $xl
}

function Release-Excel ($xl) {
    try { $xl.Quit() } catch {}
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Initialize-DataFile {
    if (Test-Path $script:ExcelPath) { return }
    $xl = New-Excel
    try {
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Sheets(1)
        $ws.Name = "InventoryData"
        $ws.Cells(1,1).Value2 = "SerialNumber"
        $ws.Cells(1,2).Value2 = "Location"
        $ws.Range("A1:B1").Font.Bold = $true
        $ws.Columns("A:B").AutoFit()
        $wb.SaveAs($script:ExcelPath, 51)   # 51 = xlOpenXMLWorkbook
        $wb.Close($false)
    } finally { Release-Excel $xl }
}

function Load-InventoryData {
    $script:Inventory.Clear()
    if (-not (Test-Path $script:ExcelPath)) { return }
    $xl = New-Excel
    try {
        $wb = $xl.Workbooks.Open($script:ExcelPath, $false, $true)   # read-only
        $ws = $wb.Sheets(1)
        $last = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row          # xlUp
        for ($i = 2; $i -le $last; $i++) {
            $sn  = [string]$ws.Cells($i,1).Value2
            $loc = [string]$ws.Cells($i,2).Value2
            if ($sn.Trim() -ne "") { $script:Inventory[$sn.Trim()] = $loc.Trim() }
        }
        $wb.Close($false)
    } finally { Release-Excel $xl }
}

function Save-InventoryData {
    $xl = New-Excel
    try {
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Sheets(1)
        $ws.Name = "InventoryData"
        $ws.Cells(1,1).Value2 = "SerialNumber"
        $ws.Cells(1,2).Value2 = "Location"
        $ws.Range("A1:B1").Font.Bold = $true

        $row = 2
        foreach ($key in ($script:Inventory.Keys | Sort-Object)) {
            $ws.Cells($row,1).Value2 = $key
            $ws.Cells($row,2).Value2 = $script:Inventory[$key]
            $row++
        }
        $ws.Columns("A:B").AutoFit()

        if (Test-Path $script:ExcelPath) { Remove-Item $script:ExcelPath -Force }
        $wb.SaveAs($script:ExcelPath, 51)
        $wb.Close($false)
    } finally { Release-Excel $xl }
}

# ─── Data operations ──────────────────────────────────────────────────────────
function Upsert-Item ([string]$Serial, [string]$Location) {
    $script:Inventory[$Serial.Trim()] = $Location.Trim()
    Save-InventoryData
}

function Find-Item ([string]$Serial) {
    $key = $Serial.Trim()
    if ($script:Inventory.ContainsKey($key)) { return $script:Inventory[$key] }
    return $null
}

# ─── Colours ──────────────────────────────────────────────────────────────────
$Blue  = [System.Drawing.Color]::FromArgb(0, 120, 212)
$Green = [System.Drawing.Color]::FromArgb(16, 124, 16)
$Red   = [System.Drawing.Color]::FromArgb(196, 43, 28)
$Gray  = [System.Drawing.Color]::FromArgb(96, 96, 96)

function New-Label ($text, $x, $y, $w, $h) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $text
    $l.Location  = New-Object System.Drawing.Point($x, $y)
    $l.Size      = New-Object System.Drawing.Size($w, $h)
    $l.TextAlign = "MiddleLeft"
    return $l
}

function New-TextBox ($x, $y, $w) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($x, $y)
    $t.Size     = New-Object System.Drawing.Size($w, 26)
    return $t
}

function New-Button ($text, $x, $y, $w, $h) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $text
    $b.Location  = New-Object System.Drawing.Point($x, $y)
    $b.Size      = New-Object System.Drawing.Size($w, $h)
    $b.BackColor = $Blue
    $b.ForeColor = [System.Drawing.Color]::White
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 0
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

# ─── Main form ────────────────────────────────────────────────────────────────
function Show-MainForm {
    $form              = New-Object System.Windows.Forms.Form
    $form.Text         = "Inventory Tool"
    $form.Size         = New-Object System.Drawing.Size(450, 340)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox  = $false
    $form.BackColor    = [System.Drawing.Color]::White
    $form.Font         = New-Object System.Drawing.Font("Segoe UI", 10)

    # TabControl
    $tabs          = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(10, 10)
    $tabs.Size     = New-Object System.Drawing.Size(415, 280)

    # ── Tab 1: Add Item ───────────────────────────────────────────────────────
    $tabAdd          = New-Object System.Windows.Forms.TabPage
    $tabAdd.Text     = "  Add Item  "
    $tabAdd.BackColor = [System.Drawing.Color]::White
    $tabAdd.Padding  = New-Object System.Windows.Forms.Padding(10)

    $lblSn1   = New-Label "Serial Number:"    10  22 130 24
    $txtSn1   = New-TextBox                  148  20 230
    $lblLoc   = New-Label "Storage Location:" 10  58 130 24
    $txtLoc   = New-TextBox                  148  56 230
    $btnAdd   = New-Button "Add to List"     148  96 120 32
    $lblAddStatus = New-Label "" 10 142 390 72
    $lblAddStatus.AutoSize = $false

    $tabAdd.Controls.AddRange(@($lblSn1, $txtSn1, $lblLoc, $txtLoc, $btnAdd, $lblAddStatus))

    # ── Tab 2: Lookup ─────────────────────────────────────────────────────────
    $tabLookup          = New-Object System.Windows.Forms.TabPage
    $tabLookup.Text     = "  Lookup  "
    $tabLookup.BackColor = [System.Drawing.Color]::White

    $lblSn2   = New-Label "Serial Number:"  10  22 130 24
    $txtSn2   = New-TextBox               148  20 230
    $btnFind  = New-Button "Find Location" 148  58 120 32

    $lblResult          = New-Object System.Windows.Forms.Label
    $lblResult.Location = New-Object System.Drawing.Point(10, 108)
    $lblResult.Size     = New-Object System.Drawing.Size(390, 100)
    $lblResult.Font     = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $lblResult.AutoSize = $false

    $tabLookup.Controls.AddRange(@($lblSn2, $txtSn2, $btnFind, $lblResult))

    $tabs.TabPages.AddRange(@($tabAdd, $tabLookup))
    $form.Controls.Add($tabs)

    # ── Add-Item handler ──────────────────────────────────────────────────────
    $doAdd = {
        $sn  = $txtSn1.Text.Trim()
        $loc = $txtLoc.Text.Trim()

        if ($sn -eq "") {
            $lblAddStatus.ForeColor = $Red
            $lblAddStatus.Text = "Please enter a serial number."
            return
        }
        if ($loc -eq "") {
            $lblAddStatus.ForeColor = $Red
            $lblAddStatus.Text = "Please enter a storage location."
            return
        }

        $existing = Find-Item $sn
        if ($null -ne $existing) {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "Serial '$sn' is already stored at:`n$existing`n`nUpdate to new location '$loc'?",
                "Serial already exists",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        $lblAddStatus.ForeColor = $Gray
        $lblAddStatus.Text = "Saving to Excel..."
        [System.Windows.Forms.Application]::DoEvents()

        Upsert-Item $sn $loc

        $verb = if ($null -ne $existing) { "Updated" } else { "Added" }
        $lblAddStatus.ForeColor = $Green
        $lblAddStatus.Text = "$verb successfully:`n  Serial:   $sn`n  Location: $loc"

        $txtSn1.Clear(); $txtLoc.Clear(); $txtSn1.Focus()
    }

    $btnAdd.Add_Click($doAdd)
    $txtSn1.Add_KeyDown({ if ($_.KeyCode -eq "Return") { & $doAdd } })
    $txtLoc.Add_KeyDown({ if ($_.KeyCode -eq "Return") { & $doAdd } })

    # ── Lookup handler ────────────────────────────────────────────────────────
    $doLookup = {
        $sn = $txtSn2.Text.Trim()
        if ($sn -eq "") {
            $lblResult.ForeColor = $Red
            $lblResult.Text = "Please enter a serial number."
            return
        }
        $loc = Find-Item $sn
        if ($null -eq $loc) {
            $lblResult.ForeColor = $Red
            $lblResult.Text = "Not found"
        } else {
            $lblResult.ForeColor = $Green
            $lblResult.Text = $loc
        }
    }

    $btnFind.Add_Click($doLookup)
    $txtSn2.Add_KeyDown({ if ($_.KeyCode -eq "Return") { & $doLookup } })

    # Clear result when the user starts typing a new serial
    $txtSn2.Add_TextChanged({ $lblResult.Text = "" })
    $tabs.Add_SelectedIndexChanged({ $lblAddStatus.Text = ""; $lblResult.Text = "" })

    [void]$form.ShowDialog()
    $form.Dispose()
}

# ─── Entry point ──────────────────────────────────────────────────────────────
try {
    Initialize-DataFile
    Load-InventoryData
    Show-MainForm
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Error: $_`n`nMake sure Microsoft Excel is installed.",
        "Inventory Tool - Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}
