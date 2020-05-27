<#
FormsMenu.ps1

Wayne Lindimore
wlindimore@gmail.com
AdminsCache.Wordpress.com

7-26-14
PowerShell WinForms Menu Demo
#>

# Install .Net Assemblies
[reflection.assembly]::load("System.Windows.Forms") | Out-Null
[reflection.assembly]::load("System.Drawing") | Out-Null

[reflection.assembly]::load("System.Runtime.InteropServices") | Out-Null
 
# Enable Visual Styles
[Windows.Forms.Application]::EnableVisualStyles()

# WinForm Setup
################################################################## Objects
# Main Form .Net Objects
$mainForm         = New-Object System.Windows.Forms.Form
$menuMain         = New-Object System.Windows.Forms.MenuStrip
$menuFile         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuView         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuTools        = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOpen         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuSave         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuSaveAs       = New-Object System.Windows.Forms.ToolStripMenuItem
$menuFullScr      = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOptions      = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOptions1     = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOptions2     = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExit         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelp         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout        = New-Object System.Windows.Forms.ToolStripMenuItem
$mainToolStrip    = New-Object System.Windows.Forms.ToolStrip
$toolStripOpen    = New-Object System.Windows.Forms.ToolStripButton
$toolStripSave    = New-Object System.Windows.Forms.ToolStripButton
$toolStripSaveAs  = New-Object System.Windows.Forms.ToolStripButton
$toolStripFullScr = New-Object System.Windows.Forms.ToolStripButton
$toolStripAbout   = New-Object System.Windows.Forms.ToolStripButton
$toolStripExit    = New-Object System.Windows.Forms.ToolStripButton
$statusStrip      = New-Object System.Windows.Forms.StatusStrip
$statusLabel      = New-Object System.Windows.Forms.ToolStripStatusLabel

################################################################## Icons
# WinForms Icons
# Create Icon Extractor Assembly
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@
Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

# Extract PowerShell Icon from PowerShell Exe
$iconPS   = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)

################################################################## Main Form Setup
# Main Form
$mainForm.Height          = 400
$mainForm.Icon            = $iconPS
$mainForm.MainMenuStrip   = $menuMain
$mainForm.Width           = 800
$mainForm.StartPosition   = "CenterScreen"
$mainForm.Text            = " WinForms Menu Demo"
$mainForm.Controls.Add($menuMain)

################################################################## Main Menu
# Main ToolStrip
[void]$mainForm.Controls.Add($mainToolStrip)

# Main Menu Bar
[void]$mainForm.Controls.Add($menuMain)

# Menu Options - File
$menuFile.Text = "&File"
[void]$menuMain.Items.Add($menuFile)

# Menu Options - File / Open
$menuOpen.Image        = [System.IconExtractor]::Extract("shell32.dll", 4, $true)
$menuOpen.ShortcutKeys = "Control, O"
$menuOpen.Text         = "&Open"
$menuOpen.Add_Click({OpenFile})
[void]$menuFile.DropDownItems.Add($menuOpen)

# Menu Options - File / Save
$menuSave.Image        = [System.IconExtractor]::Extract("shell32.dll", 36, $true)
$menuSave.ShortcutKeys = "F2"
$menuSave.Text         = "&Save"
$menuSave.Add_Click({SaveFile})
[void]$menuFile.DropDownItems.Add($menuSave)

# Menu Options - File / Save As
$menuSaveAs.Image        = [System.IconExtractor]::Extract("shell32.dll", 45, $true)
$menuSaveAs.ShortcutKeys = "Control, S"
$menuSaveAs.Text         = "&Save As"
$menuSaveAs.Add_Click({SaveAs})
[void]$menuFile.DropDownItems.Add($menuSaveAs)

# Menu Options - File / Exit
$menuExit.Image        = [System.IconExtractor]::Extract("shell32.dll", 10, $true)
$menuExit.ShortcutKeys = "Control, X"
$menuExit.Text         = "&Exit"
$menuExit.Add_Click({$mainForm.Close()})
[void]$menuFile.DropDownItems.Add($menuExit)

# Menu Options - View
$menuView.Text      = "&View"
[void]$menuMain.Items.Add($menuView)

# Menu Options - View / Full Screen
$menuFullScr.Image        = [System.IconExtractor]::Extract("shell32.dll",34, $true)
$menuFullScr.ShortcutKeys = "Control, F"
$menuFullScr.Text         = "&Full Screen"
$menuFullScr.Add_Click({FullScreen})
[void]$menuView.DropDownItems.Add($menuFullScr)

# Menu Options - Tools
$menuTools.Text      = "&Tools"
[void]$menuMain.Items.Add($menuTools)

# Menu Options - Tools / Options
$menuOptions.Image     = [System.IconExtractor]::Extract("shell32.dll", 21, $true)
$menuOptions.Text      = "&Options"
[void]$menuTools.DropDownItems.Add($menuOptions)

# Menu Options - Tools / Options / Options 1
$menuOptions1.Image     = [System.IconExtractor]::Extract("shell32.dll", 33, $true)
$menuOptions1.Text      = "&Options 1"
$menuOptions1.Add_Click({Options1})
[void]$menuOptions.DropDownItems.Add($menuOptions1)

# Menu Options - Tools / Options / Options 2
$menuOptions2.Image     = [System.IconExtractor]::Extract("shell32.dll", 35, $true)
$menuOptions2.Text      = "&Options 2"
$menuOptions2.Add_Click({Options2})
[void]$menuOptions.DropDownItems.Add($menuOptions2)

# Menu Options - Help
$menuHelp.Text      = "&Help"
[void]$menuMain.Items.Add($menuHelp)

# Menu Options - Help / About
$menuAbout.Image     = [System.Drawing.SystemIcons]::Information
$menuAbout.Text      = "About MenuStrip"
$menuAbout.Add_Click({About})
[void]$menuHelp.DropDownItems.Add($menuAbout)

################################################################## ToolBar Buttons
# ToolStripButton - Open
$toolStripOpen.ToolTipText  = "Open"
$toolStripOpen.Image = $menuOpen.Image
$toolStripOpen.Add_Click({OpenFile})
[void]$mainToolStrip.Items.Add($toolStripOpen)

# ToolStripButton - Save
$toolStripSave.ToolTipText  = "Save"
$toolStripSave.Image = $menuSave.Image
$toolStripSave.Add_Click({Save})
[void]$mainToolStrip.Items.Add($toolStripSave)

# ToolStripButton - SaveAs
$toolStripSaveAs.ToolTipText  = "SaveAs"
$toolStripSaveAs.Image = $menuSaveAs.Image
$toolStripSaveAs.Add_Click({SaveAs})
[void]$mainToolStrip.Items.Add($toolStripSaveAs)

# ToolStripButton - Full Screen
$toolStripFullScr.ToolTipText  = "Full Screen"
$toolStripFullScr.Image = $menuFullScr.Image
$toolStripFullScr.Add_Click({FullScreen})
[void]$mainToolStrip.Items.Add($toolStripFullScr)

# ToolStripButton - About
$toolStripAbout.ToolTipText  = "About"
$toolStripAbout.Image = $menuAbout.Image
$toolStripAbout.Add_Click({About})
[void]$mainToolStrip.Items.Add($toolStripAbout)

# ToolStripButton - Exit
$toolStripExit.ToolTipText  = "Exit"
$toolStripExit.Image = $menuExit.Image
$toolStripExit.Add_Click({$mainForm.Close()})
[void]$mainToolStrip.Items.Add($toolStripExit)

################################################################## Status Bar
# Status Bar & Label
[void]$statusStrip.Items.Add($statusLabel)
$statusLabel.AutoSize  = $true
$statusLabel.Text      = "Ready"
$mainForm.Controls.Add($statusStrip)

################################################################## Functions
function OpenFile {
    $statusLabel.Text = "Open File"
	$selectOpenForm = New-Object System.Windows.Forms.OpenFileDialog
	$selectOpenForm.Filter = "All Files (*.*)|*.*"
	$selectOpenForm.InitialDirectory = ".\"
	$selectOpenForm.Title = "Select a File to Open"
	$getKey = $selectOpenForm.ShowDialog()
	If ($getKey -eq "OK") {
            $inputFileName = $selectOpenForm.FileName
	}
    $statusLabel.Text = "Ready"
}

function SaveAs {
    $statusLabel.Text = "Save As"
    $selectSaveAsForm = New-Object System.Windows.Forms.SaveFileDialog
	$selectSaveAsForm.Filter = "All Files (*.*)|*.*"
	$selectSaveAsForm.InitialDirectory = ".\"
	$selectSaveAsForm.Title = "Select a File to Save"
	$getKey = $selectSaveAsForm.ShowDialog()
	If ($getKey -eq "OK") {
            $outputFileName = $selectSaveAsForm.FileName
	}
    $statusLabel.Text = "Ready"
}

function SaveFile {
}

function FullScreen {
}

function Options1 {
}

function Options2 {
}

function About {
    $statusLabel.Text = "About"
    # About Form Objects
    $aboutForm          = New-Object System.Windows.Forms.Form
    $aboutFormExit      = New-Object System.Windows.Forms.Button
    $aboutFormImage     = New-Object System.Windows.Forms.PictureBox
    $aboutFormNameLabel = New-Object System.Windows.Forms.Label
    $aboutFormText      = New-Object System.Windows.Forms.Label

    # About Form
    $aboutForm.AcceptButton  = $aboutFormExit
    $aboutForm.CancelButton  = $aboutFormExit
    $aboutForm.ClientSize    = "350, 110"
    $aboutForm.ControlBox    = $false
    $aboutForm.ShowInTaskBar = $false
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.Text          = "About FormsMenu.ps1"
    $aboutForm.Add_Load($aboutForm_Load)

    # About PictureBox
    $aboutFormImage.Image    = $iconPS.ToBitmap()
    $aboutFormImage.Location = "55, 15"
    $aboutFormImage.Size     = "32, 32"
    $aboutFormImage.SizeMode = "StretchImage"
    $aboutForm.Controls.Add($aboutFormImage)

    # About Name Label
    $aboutFormNameLabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $aboutFormNameLabel.Location = "110, 20"
    $aboutFormNameLabel.Size     = "200, 18"
    $aboutFormNameLabel.Text     = "WinForms Menu Demo"
    $aboutForm.Controls.Add($aboutFormNameLabel)

    # About Text Label
    $aboutFormText.Location = "100, 40"
    $aboutFormText.Size     = "300, 30"
    $aboutFormText.Text     = "          Wayne Lindimore `n`r AdminsCache.WordPress.com"
    $aboutForm.Controls.Add($aboutFormText)

    # About Exit Button
    $aboutFormExit.Location = "135, 70"
    $aboutFormExit.Text     = "OK"
    $aboutForm.Controls.Add($aboutFormExit)

    [void]$aboutForm.ShowDialog()
    $statusLabel.Text = "Ready"
} # End About

# Show Main Form
[void] $mainForm.ShowDialog()