# Function to enable all add-ins in Word
function Enable-AllAddinsInWord {
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $addIns = $word.COMAddIns
        foreach ($addIn in $addIns) {
            $addIn.Connect = $true  # Enable the add-in
            Write-Output "Enabled Word Add-in: $($addIn.Description)"
        }
    } catch {
        Write-Error "Error enabling add-ins in Word: $_"
    } finally {
        if ($word -is [__ComObject]) { $word.Quit() }
    }
}

# Function to enable all add-ins in Excel
function Enable-AllAddinsInExcel {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $addIns = $excel.COMAddIns
        foreach ($addIn in $addIns) {
            $addIn.Connect = $true  # Enable the add-in
            Write-Output "Enabled Excel Add-in: $($addIn.Description)"
        }
    } catch {
        Write-Error "Error enabling add-ins in Excel: $_"
    } finally {
        if ($excel -is [__ComObject]) { $excel.Quit() }
    }
}

# Function to enable all add-ins in PowerPoint
function Enable-AllAddinsInPowerPoint {
    try {
        $powerpoint = New-Object -ComObject PowerPoint.Application
        $powerpoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $addIns = $powerpoint.COMAddIns
        foreach ($addIn in $addIns) {
            $addIn.Connect = $true  # Enable the add-in
            Write-Output "Enabled PowerPoint Add-in: $($addIn.Description)"
        }
    } catch {
        Write-Error "Error enabling add-ins in PowerPoint: $_"
    } finally {
        if ($powerpoint -is [__ComObject]) { $powerpoint.Quit() }
    }
}

# Activate all add-ins in Word, Excel, and PowerPoint
Enable-AllAddinsInWord
Enable-AllAddinsInExcel
Enable-AllAddinsInPowerPoint
