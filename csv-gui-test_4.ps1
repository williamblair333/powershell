[reflection.assembly]::load("System.Windows.Forms") | Out-Null
[reflection.assembly]::load("System.Drawing") | Out-Null

# Enable Visual Styles
[Windows.Forms.Application]::EnableVisualStyles()

# This block of code is a file dialog open box
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
	# This uses Environmental Directory Settings such as Desktop
    # InitialDirectory = [Environment]::GetFolderPath('Desktop')

	# Here you can pick the default directory.  If you comment it out, 
	# then OpenFileDialog seems to go to the last directory you chose a file from
	InitialDirectory = "C:\temp\Ricoh"

    # Pick file extensions to chose from.  The first is default.
	Filter = 'Comma Separated Values (*.csv)|*.csv|All Files (*.*)|*.*|SpreadSheet (*.xlsx)|*.xlsx'
}

# The actual dialog open box
$dialogOpen = $FileBrowser.ShowDialog()

# This returns the full path and filename
$pathFileName = $FileBrowser.FileName
# Write-Host $pathFileName 

# This returns the filename only
$fileName = $FileBrowser.SafeFileName
# Write-Host $fileName 

$OnLoadForm_UpdateGrid= {

	# This returns the filename minus extension.
	$baseName = Get-Item $pathFileName | Select-Object -ExpandProperty BaseName
	# Write-Host $baseName
	
	$tmp = $FileBrowser.InitialDirectory + '\' + $baseName + '.tmp'
	Write-Host $tmp

	#Make a copy of the file so we can import it and leave the real file free for exporting to
	Copy-Item $pathFileName -Destination $tmp
	
	# Load the tempfile into memory so we can work
	$tmpFileName = Import-Csv $tmp
	
	#Remove the tempfile now
	Remove-Item $tmp
	
	#Select the datasource so we can prep for the dataGridView
	$dataGridView1.DataSource=[System.Collections.ArrayList]$tmpFileName

    $form.refresh()
}

$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Form Text Goes Here"
$Form.TopMost = $true

$form.KeyPreview = $true
$form.StartPosition = "centerscreen"

$sds_width = 900
$sds_height = 600

$form.MinimumSize= New-Object System.Drawing.Size($sds_width,$sds_height)
$form.Size = New-Object System.Drawing.Size($sds_width,$sds_height)
$form.FormBorderStyle = 'Sizable'
$form.AutoScroll = $true

#dg
$dataGridView1 = New-Object System.Windows.Forms.DataGridView -Property @{
}

$dataGridView1.MinimumSize= New-Object System.Drawing.Size(($sds_width - 25),($sds_height - 100))
$dataGridView1.Size = New-Object System.Drawing.Size(($sds_width - 25),($sds_height - 100))

$dataGridView1.AllowUserToOrderColumns = $true
$dataGridView1AllowUserToResizeColumns = $true
$dataGridView1.AutoResizeColumns()

$dataGridView1.AutoSize = $true

$dataGridView1.ScrollBars = 'Both'

$dataGridView1.Name = $baseName
$dataGridView1.DataMember = ""
$dataGridView1.TabIndex = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 5
$System_Drawing_Point.Y = 5

$dataGridView1.Location = $System_Drawing_Point

#dg

$form.Controls.Add($dataGridView1)
$form.add_Load($OnLoadForm_UpdateGrid)

# This button will save the file
$button1_OnClick= {    
	$dataGridView1.Rows |Select -Expand DataBoundItem | Export-Csv $pathFileName -NoType
}
 
$button = New-Object Windows.Forms.Button
$button.text = "Save"
$button.Location = New-Object Drawing.Point(5,($dataGridView1.height + 25))
$button.Size = New-Object Drawing.Point(125, 25)

$button.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right

$button.TabIndex ="1"
$button.add_Click($button1_OnClick)
$form.controls.add($button)

$form.ShowDialog()
$form.Dispose()