# Change the target path and working directory of shortcut files in the directory where this script is located

# Set the strings to be changed
Set-Variable -Name "OLD_LINK" -Value "\\old_unc\"
Set-Variable -Name "OLD_LINK_REX" -Value "\\\\old_unc\\"
Set-Variable -Name "NEW_LINK" -Value "\\new_unc\"

# Get the path of the directory where the script is located
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Create a shortcut object
$WScriptShell = New-Object -ComObject WScript.Shell

# Get the list of shortcut files in the directory
$Files = Get-ChildItem -Path $scriptDirectory -Filter *.lnk

# Change the target path of the shortcut files
$cnt = 0
foreach($File in $Files) {
    $changed = 0
    $shortcutPath = $File.FullName
    $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
    $targetPath = $shortcut.TargetPath.ToLower()
    $workingDirectory = $shortcut.WorkingDirectory.ToLower()

    # If the target path is subject to change
    if ($targetPath -like "*$OLD_LINK*") {
        $targetPath = $targetPath -replace $OLD_LINK_REX, $NEW_LINK
        $shortcut.TargetPath = $targetPath 
        $shortcut.Save()
        $changed = 1
    }

    # If the working directory is subject to change
    if ($workingDirectory -like "*$OLD_LINK*") {
        $workingDirectory = $workingDirectory -replace $OLD_LINK_REX, $NEW_LINK
        $shortcut.WorkingDirectory = $workingDirectory 
        $shortcut.Save()
        $changed = 1
    }

    # If there was a change
    if($changed) {
        $cnt = $cnt + 1
    }
}

# Number of processed shortcuts
"Processed shortcuts"
$cnt
