#	Autor: Michal Salinski
#	Wersja: 1.2.3
#
#
#
#
#
Import-Module ActiveDirectory

# CHANGE THIS VARIABLES ACCORDING TO YOUR DOMAIN CONFIGURATION
$dcvariable = "DC=example,DC=com"
$atvariable = "@example.com"
#############################################################


$charSet = @( )
$charSetUpper = @( )
$numberSet = @( )
$specialSet = @( )
$omfg = "Something went wrong"

# Numbers from 0 to 9 is 48..57
50..57 | ForEach-Object {
	$numberSet +=, [char][byte]$_
}

# A .. Z
65..72 + 74..78 + 80..90|ForEach-Object{
    $charSetUpper +=,[char][byte]$_
}

# a .. z
97..107 + 109..117 + 119..122 | ForEach-Object {
	$charSet +=, [char][byte]$_
}


# special chars
33..33 + 35..38 + 64..64 | ForEach-Object {
	$specialSet +=, [char][byte]$_
}


function passwordGenerator
{
	$password = ""
    1..3 | foreach {$password += ($charSetUpper | Get-Random)}
    1..3 | foreach {$password += ($charSet | Get-Random)}
    1..3 | foreach {$password += ($numberSet | Get-Random)}
	return $password
}

function getUserInfoAD([string]$id) {
	try{
        $ret = Get-ADUser -filter "SamAccountName -eq '$id'" | select -ExpandProperty Name
        return $ret
    } catch {
        pconsole($omfg)
    }
}

function getComputerInfoLess([string]$id){
    try 
    {
        [array]$data=Get-ADComputer -Identity $id -Properties * | select -ExpandProperty DistinguishedName
    }
    catch
    {
        $data = $false
    }
    return $data
}

function getUserInfoADMore([string]$id) {
    try {
	    $ret = Get-ADUser -filter "SamAccountName -eq '$id'" -Properties *
        return $ret
    } catch {
        pconsole($omfg)
    }
}


function getUserInfoByNameMore([string]$id) {
    try {
	    $ret = Get-ADUser -f {(sn -like $id) -or (givenname -like $id)} -Properties *
        return $ret
    } catch {
        pconsole($omfg)
    }
}

function getComputerInfo([string]$id){
    try{
        [array]$data=(Get-ADComputer -Identity $id -Properties *)
        return $data
    } catch {
        pconsole($omfg)
    }
}

function resetPassword([string]$login, $condition)
{
	$password = passwordGenerator
	$adPassword = ConvertTo-SecureString -String $password -Force -AsPlainText
	try{
        set-ADAccountPassword $login -NewPassword $adPassword -Reset -Verbose -PassThru
        set-ADUser $login -ChangePasswordAtLogon $condition
	    Write-Host $password
        return @($password)
    } catch{
        pconsole($omfg)
        return $omfg
    }
}

function fillTextBox9(){
    $TextBox9.Text = $TextBox7.Text+" "+$TextBox8.Text
}

function clearModifiedData(){
    $TextBox7.Clear()
    $TextBox8.Clear()
    $TextBox9.Clear()
    $TextBox10.Clear()
    $TextBox11.Clear()
    $TextBox12.Clear()
    $CheckBox7.Checked = $false
    $CheckBox8.Checked = $false
    $CheckBox9.Checked = $false
    $ListBox20.Items.Clear()
    $label34.Text = "..."
}

function fillModifiedData($data){
    clearModifiedData
    $TextBox7.Text = $data.GivenName
    $TextBox8.Text = $data.Surname
    $TextBox9.Text = $data.DisplayName
    $TextBox10.Text = $data.Description
    $TextBox11.Text = $data.mail
    $TextBox12.Text = $data.SamAccountName
    if($data.LockedOut){
        $CheckBox7.Checked = $true
    }
    if($data.PasswordExpired){
        $CheckBox8.Checked = $true
    }
    if($data.Enabled){
        $CheckBox9.Checked = $false
    } else {
        $CheckBox9.Checked = $true
    }
    $ListBox20.Items.AddRange($data.MemberOf)
    $label34.Text = $data.DistinguishedName
    Write-Host "$global:selectedModifiedUser information downloaded!"
}

function pconsole($data){
    $newline = [System.Environment]::NewLine
    $RichTextBox7.Text = $data+$newline+$RichTextBox7.Text
    Out-File ($env:userprofile+"\Simple_Active_Directory_Administration_Console.log") -inputObject ((Get-Date -Format G)+">  "+$data) -encoding "UTF8" -Append -Force
    Write-Host $data
}

function copyToClipboard($data){
    #HERE YOU CAN ADD CUSTOM MESSAGE
    $message = "Passwords for users:`n"
	#This loop adds to message informations from ListBox
    foreach ($item in $data.Items)
	{
		$message += "$item `n"
	}
    #HERE YOU CAN ADD MORE INFORMATION LIKE THIS EXAMPLE:
    $message += "`n" # this adds empty line
    $message += "`n"
    $message += "You are required to change your password at first logon.`n"
    
    try
    {
        [Windows.Forms.Clipboard]::SetText($message)
        pconsole("Copied to clipboard")
    }
    catch 
    {
        pconsole("Field is empty!")
    }

}

#region ScriptForm Designer

#region Constructor

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#endregion

#region Post-Constructor Custom Code

#endregion

#region Form Creation
#Warning: It is recommended that changes inside this region be handled using the ScriptForm Designer.
#When working with the ScriptForm designer this region and any changes within may be overwritten.
#~~< Form1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Form1 = New-Object System.Windows.Forms.Form
$Form1.ClientSize = New-Object System.Drawing.Size(1061, 696)
$Form1.Text = "Simple Active Directory Administration Console"
#~~< Button41 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button41 = New-Object System.Windows.Forms.Button
$Button41.Location = New-Object System.Drawing.Point(661, 669)
$Button41.Size = New-Object System.Drawing.Size(116, 23)
$Button41.TabIndex = 8
$Button41.Text = "Clear"
$Button41.UseVisualStyleBackColor = $true
$Button41.add_Click({Button41Click($Button41)})
#~~< Button40 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button40 = New-Object System.Windows.Forms.Button
$Button40.Location = New-Object System.Drawing.Point(783, 669)
$Button40.Size = New-Object System.Drawing.Size(116, 23)
$Button40.TabIndex = 7
$Button40.Text = "Copy"
$Button40.UseVisualStyleBackColor = $true
$Button40.add_Click({Button40Click($Button40)})
#~~< RichTextBox7 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RichTextBox7 = New-Object System.Windows.Forms.RichTextBox
$RichTextBox7.Font = New-Object System.Drawing.Font("Source Code Pro", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](238)))
$RichTextBox7.Location = New-Object System.Drawing.Point(3, 496)
$RichTextBox7.ReadOnly = $true
$RichTextBox7.Size = New-Object System.Drawing.Size(1052, 167)
$RichTextBox7.TabIndex = 6
$RichTextBox7.Text = ""
$RichTextBox7.BackColor = [System.Drawing.SystemColors]::HotTrack
$RichTextBox7.ForeColor = [System.Drawing.SystemColors]::Window
#~~< Button38 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button38 = New-Object System.Windows.Forms.Button
$Button38.Location = New-Object System.Drawing.Point(3, 668)
$Button38.Size = New-Object System.Drawing.Size(48, 23)
$Button38.TabIndex = 5
$Button38.Text = "info"
$Button38.UseVisualStyleBackColor = $true
$Button38.add_Click({Button38Click($Button38)})
#~~< Label35 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label35 = New-Object System.Windows.Forms.Label
$Label35.Location = New-Object System.Drawing.Point(57, 669)
$Label35.Size = New-Object System.Drawing.Size(122, 23)
$Label35.TabIndex = 4
$Label35.Text = "Author: Michal Salinski"
$Label35.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$Label35.Visible = $false
$Label35.add_Click({Label35Click($Label35)})
#~~< TabControl1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabControl1 = New-Object System.Windows.Forms.TabControl
$TabControl1.Location = New-Object System.Drawing.Point(3, 2)
$TabControl1.Size = New-Object System.Drawing.Size(1056, 492)
$TabControl1.TabIndex = 0
$TabControl1.Text = "Computer management"
#~~< TabPage1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabPage1 = New-Object System.Windows.Forms.TabPage
$TabPage1.Location = New-Object System.Drawing.Point(4, 22)
$TabPage1.Padding = New-Object System.Windows.Forms.Padding(3)
$TabPage1.Size = New-Object System.Drawing.Size(1048, 466)
$TabPage1.TabIndex = 0
$TabPage1.Text = "Passwords"
$TabPage1.UseVisualStyleBackColor = $true
#~~< GroupBox3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox3 = New-Object System.Windows.Forms.GroupBox
$GroupBox3.Location = New-Object System.Drawing.Point(7, 143)
$GroupBox3.Size = New-Object System.Drawing.Size(200, 315)
$GroupBox3.TabIndex = 6
$GroupBox3.TabStop = $false
$GroupBox3.Text = "ID's"
#~~< Button4 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button4 = New-Object System.Windows.Forms.Button
$Button4.Location = New-Object System.Drawing.Point(6, 289)
$Button4.Size = New-Object System.Drawing.Size(188, 23)
$Button4.TabIndex = 1
$Button4.Text = "Load"
$Button4.UseVisualStyleBackColor = $true
$Button4.add_Click({Button4Click($Button4)})
#~~< RichTextBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RichTextBox1 = New-Object System.Windows.Forms.RichTextBox
$RichTextBox1.Location = New-Object System.Drawing.Point(6, 19)
$RichTextBox1.Size = New-Object System.Drawing.Size(188, 263)
$RichTextBox1.TabIndex = 0
$RichTextBox1.Text = "Fill in ID's in separate lines"
$RichTextBox1.add_MouseClick({RichTextBox1MouseClick($RichTextBox1)})
$GroupBox3.Controls.Add($Button4)
$GroupBox3.Controls.Add($RichTextBox1)
#~~< GroupBox2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox2 = New-Object System.Windows.Forms.GroupBox
$GroupBox2.Location = New-Object System.Drawing.Point(213, 7)
$GroupBox2.Size = New-Object System.Drawing.Size(829, 451)
$GroupBox2.TabIndex = 5
$GroupBox2.TabStop = $false
$GroupBox2.Text = "Reset passwords"
#~~< Button10 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button10 = New-Object System.Windows.Forms.Button
$Button10.Location = New-Object System.Drawing.Point(713, 392)
$Button10.Size = New-Object System.Drawing.Size(110, 53)
$Button10.TabIndex = 12
$Button10.Text = "Copy to clipboard"
$Button10.UseVisualStyleBackColor = $true
$Button10.add_Click({Button10Click($Button10)})
#~~< ProgressBar1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ProgressBar1 = New-Object System.Windows.Forms.ProgressBar
$ProgressBar1.Location = New-Object System.Drawing.Point(345, 421)
$ProgressBar1.Size = New-Object System.Drawing.Size(362, 23)
$ProgressBar1.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$ProgressBar1.TabIndex = 11
$ProgressBar1.Text = ""
#~~< ukrytalista >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ukrytalista = New-Object System.Windows.Forms.ListBox
$ukrytalista.FormattingEnabled = $true
$ukrytalista.Location = New-Object System.Drawing.Point(323, 262)
$ukrytalista.SelectedIndex = -1
$ukrytalista.Size = New-Object System.Drawing.Size(17, 17)
$ukrytalista.TabIndex = 7
$ukrytalista.Visible = $false
#~~< CheckBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox1 = New-Object System.Windows.Forms.CheckBox
$CheckBox1.Checked = $true
$CheckBox1.CheckState = [System.Windows.Forms.CheckState]::Checked
$CheckBox1.Location = New-Object System.Drawing.Point(456, 393)
$CheckBox1.Size = New-Object System.Drawing.Size(235, 23)
$CheckBox1.TabIndex = 10
$CheckBox1.Text = "Change password at first logon"
$CheckBox1.UseVisualStyleBackColor = $true
#~~< Button5 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Point(345, 392)
$Button5.Size = New-Object System.Drawing.Size(105, 23)
$Button5.TabIndex = 6
$Button5.Text = "Reset passwords"
$Button5.UseVisualStyleBackColor = $true
$Button5.add_Click({Button5Click($Button5)})
#~~< Label4 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label4 = New-Object System.Windows.Forms.Label
$Label4.Location = New-Object System.Drawing.Point(345, 19)
$Label4.Size = New-Object System.Drawing.Size(430, 15)
$Label4.TabIndex = 5
$Label4.Text = "New passwords"
#~~< ListBox3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox3 = New-Object System.Windows.Forms.ListBox
$ListBox3.FormattingEnabled = $true
$ListBox3.Location = New-Object System.Drawing.Point(346, 37)
$ListBox3.SelectedIndex = -1
$ListBox3.Size = New-Object System.Drawing.Size(477, 342)
$ListBox3.TabIndex = 4
#~~< Label3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label3 = New-Object System.Windows.Forms.Label
$Label3.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](238)))
$Label3.Location = New-Object System.Drawing.Point(7, 263)
$Label3.Size = New-Object System.Drawing.Size(99, 16)
$Label3.TabIndex = 3
$Label3.Text = "Not found in AD"
$Label3.ForeColor = [System.Drawing.Color]::Red
#~~< Label2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location = New-Object System.Drawing.Point(7, 15)
$Label2.Size = New-Object System.Drawing.Size(184, 15)
$Label2.TabIndex = 2
$Label2.Text = "Found"
#~~< ListBox2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox2 = New-Object System.Windows.Forms.ListBox
$ListBox2.FormattingEnabled = $true
$ListBox2.Location = New-Object System.Drawing.Point(6, 285)
$ListBox2.SelectedIndex = -1
$ListBox2.Size = New-Object System.Drawing.Size(334, 160)
$ListBox2.TabIndex = 1
#~~< ListBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox1 = New-Object System.Windows.Forms.ListBox
$ListBox1.FormattingEnabled = $true
$ListBox1.Location = New-Object System.Drawing.Point(7, 33)
$ListBox1.SelectedIndex = -1
$ListBox1.Size = New-Object System.Drawing.Size(333, 225)
$ListBox1.TabIndex = 0
$GroupBox2.Controls.Add($Button10)
$GroupBox2.Controls.Add($ProgressBar1)
$GroupBox2.Controls.Add($ukrytalista)
$GroupBox2.Controls.Add($CheckBox1)
$GroupBox2.Controls.Add($Button5)
$GroupBox2.Controls.Add($Label4)
$GroupBox2.Controls.Add($ListBox3)
$GroupBox2.Controls.Add($Label3)
$GroupBox2.Controls.Add($Label2)
$GroupBox2.Controls.Add($ListBox2)
$GroupBox2.Controls.Add($ListBox1)
#~~< GroupBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox1 = New-Object System.Windows.Forms.GroupBox
$GroupBox1.Location = New-Object System.Drawing.Point(6, 6)
$GroupBox1.Size = New-Object System.Drawing.Size(200, 131)
$GroupBox1.TabIndex = 4
$GroupBox1.TabStop = $false
$GroupBox1.Text = "Password generator"
#~~< Label1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Location = New-Object System.Drawing.Point(7, 16)
$Label1.Size = New-Object System.Drawing.Size(187, 24)
$Label1.TabIndex = 0
$Label1.Text = "Label1"
$Label1.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
#~~< Button1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Point(7, 43)
$Button1.Size = New-Object System.Drawing.Size(187, 26)
$Button1.TabIndex = 1
$Button1.Text = "Refresh"
$Button1.UseVisualStyleBackColor = $true
$Button1.add_Click({Button1OnClick($Button1)})
#~~< Button2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Point(7, 76)
$Button2.Size = New-Object System.Drawing.Size(187, 26)
$Button2.TabIndex = 2
$Button2.Text = "Copy"
$Button2.UseVisualStyleBackColor = $true
$Button2.add_Click({Button2OnClick($Button2)})
$GroupBox1.Controls.Add($Label1)
$GroupBox1.Controls.Add($Button1)
$GroupBox1.Controls.Add($Button2)
$TabPage1.Controls.Add($GroupBox3)
$TabPage1.Controls.Add($GroupBox2)
$TabPage1.Controls.Add($GroupBox1)
#~~< TabPage4 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabPage4 = New-Object System.Windows.Forms.TabPage
$TabPage4.Location = New-Object System.Drawing.Point(4, 22)
$TabPage4.Padding = New-Object System.Windows.Forms.Padding(3)
$TabPage4.Size = New-Object System.Drawing.Size(1048, 466)
$TabPage4.TabIndex = 3
$TabPage4.Text = "Status of user and computer"
$TabPage4.UseVisualStyleBackColor = $true
#~~< Button18 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button18 = New-Object System.Windows.Forms.Button
$Button18.Location = New-Object System.Drawing.Point(757, 39)
$Button18.Size = New-Object System.Drawing.Size(280, 23)
$Button18.TabIndex = 7
$Button18.Text = "Copy to clipboard"
$Button18.UseVisualStyleBackColor = $true
$Button18.add_Click({Button18Click($Button18)})
#~~< Button17 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button17 = New-Object System.Windows.Forms.Button
$Button17.Location = New-Object System.Drawing.Point(467, 39)
$Button17.Size = New-Object System.Drawing.Size(280, 23)
$Button17.TabIndex = 6
$Button17.Text = "Clear results"
$Button17.UseVisualStyleBackColor = $true
$Button17.add_Click({Button17Click($Button17)})
#~~< GroupBox6 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox6 = New-Object System.Windows.Forms.GroupBox
$GroupBox6.Location = New-Object System.Drawing.Point(7, 9)
$GroupBox6.Size = New-Object System.Drawing.Size(211, 53)
$GroupBox6.TabIndex = 5
$GroupBox6.TabStop = $false
$GroupBox6.Text = "Select"
#~~< CheckBox3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox3 = New-Object System.Windows.Forms.CheckBox
$CheckBox3.Location = New-Object System.Drawing.Point(114, 19)
$CheckBox3.Size = New-Object System.Drawing.Size(78, 24)
$CheckBox3.TabIndex = 1
$CheckBox3.Text = "Computer"
$CheckBox3.UseVisualStyleBackColor = $true
$CheckBox3.add_Click({CheckBox3Click($CheckBox3)})
#~~< CheckBox2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox2 = New-Object System.Windows.Forms.CheckBox
$CheckBox2.Location = New-Object System.Drawing.Point(18, 19)
$CheckBox2.Size = New-Object System.Drawing.Size(90, 24)
$CheckBox2.TabIndex = 0
$CheckBox2.Checked = $true
$CheckBox2.CheckState = [System.Windows.Forms.CheckState]::Checked
$CheckBox2.Text = "User"
$CheckBox2.UseVisualStyleBackColor = $true
$CheckBox2.add_Click({CheckBox2Click($CheckBox2)})
$GroupBox6.Controls.Add($CheckBox3)
$GroupBox6.Controls.Add($CheckBox2)
#~~< Button15 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button15 = New-Object System.Windows.Forms.Button
$Button15.Location = New-Object System.Drawing.Point(467, 13)
$Button15.Size = New-Object System.Drawing.Size(280, 23)
$Button15.TabIndex = 3
$Button15.Text = "Check"
$Button15.UseVisualStyleBackColor = $true
$Button15.add_Click({Button15Click($Button15)})
#~~< ListBox8 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox8 = New-Object System.Windows.Forms.ListBox
$ListBox8.FormattingEnabled = $true
$ListBox8.Location = New-Object System.Drawing.Point(7, 68)
$ListBox8.SelectedIndex = -1
$ListBox8.Size = New-Object System.Drawing.Size(1030, 394)
$ListBox8.TabIndex = 2
$ListBox8.add_DoubleClick({ListBox8DoubleClick($ListBox8)})
#~~< Label10 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label10 = New-Object System.Windows.Forms.Label
$Label10.Location = New-Object System.Drawing.Point(224, 18)
$Label10.Size = New-Object System.Drawing.Size(47, 23)
$Label10.TabIndex = 1
$Label10.Text = "Search:"
#~~< TextBox3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox3 = New-Object System.Windows.Forms.TextBox
$TextBox3.Location = New-Object System.Drawing.Point(277, 15)
$TextBox3.Size = New-Object System.Drawing.Size(184, 20)
$TextBox3.TabIndex = 0
$TextBox3.Text = ""
$TabPage4.Controls.Add($Button18)
$TabPage4.Controls.Add($Button17)
$TabPage4.Controls.Add($GroupBox6)
$TabPage4.Controls.Add($Button15)
$TabPage4.Controls.Add($ListBox8)
$TabPage4.Controls.Add($Label10)
$TabPage4.Controls.Add($TextBox3)
#~~< TabPage5 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabPage5 = New-Object System.Windows.Forms.TabPage
$TabPage5.Location = New-Object System.Drawing.Point(4, 22)
$TabPage5.Padding = New-Object System.Windows.Forms.Padding(3)
$TabPage5.Size = New-Object System.Drawing.Size(1048, 466)
$TabPage5.TabIndex = 4
$TabPage5.Text = "Computers"
$TabPage5.UseVisualStyleBackColor = $true
#~~< GroupBox8 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox8 = New-Object System.Windows.Forms.GroupBox
$GroupBox8.Location = New-Object System.Drawing.Point(212, 6)
$GroupBox8.Size = New-Object System.Drawing.Size(830, 454)
$GroupBox8.TabIndex = 8
$GroupBox8.TabStop = $false
$GroupBox8.Text = "Moving"
#~~< GroupBox9 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox9 = New-Object System.Windows.Forms.GroupBox
$GroupBox9.Location = New-Object System.Drawing.Point(253, 334)
$GroupBox9.Size = New-Object System.Drawing.Size(438, 31)
$GroupBox9.TabIndex = 24
$GroupBox9.TabStop = $false
$GroupBox9.Text = "Select mode"
#~~< CheckBox6 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox6 = New-Object System.Windows.Forms.CheckBox
$CheckBox6.Location = New-Object System.Drawing.Point(336, 7)
$CheckBox6.Size = New-Object System.Drawing.Size(91, 24)
$CheckBox6.TabIndex = 2
$CheckBox6.Text = "Delete"
$CheckBox6.UseVisualStyleBackColor = $true
$CheckBox6.add_Click({CheckBox6Click($CheckBox6)})
#~~< CheckBox4 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox4 = New-Object System.Windows.Forms.CheckBox
$CheckBox4.Location = New-Object System.Drawing.Point(228, 7)
$CheckBox4.Size = New-Object System.Drawing.Size(91, 24)
$CheckBox4.TabIndex = 1
$CheckBox4.Text = "Add"
$CheckBox4.UseVisualStyleBackColor = $true
$CheckBox4.add_Click({CheckBox4Click($CheckBox4)})
#~~< CheckBox5 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox5 = New-Object System.Windows.Forms.CheckBox
$CheckBox5.Checked = $true
$CheckBox5.CheckState = [System.Windows.Forms.CheckState]::Checked
$CheckBox5.Location = New-Object System.Drawing.Point(102, 7)
$CheckBox5.Size = New-Object System.Drawing.Size(102, 24)
$CheckBox5.TabIndex = 0
$CheckBox5.Text = "Move"
$CheckBox5.UseVisualStyleBackColor = $true
$CheckBox5.add_Click({CheckBox5Click($CheckBox5)})
$GroupBox9.Controls.Add($CheckBox6)
$GroupBox9.Controls.Add($CheckBox4)
$GroupBox9.Controls.Add($CheckBox5)
#~~< Label13 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label13 = New-Object System.Windows.Forms.Label
$Label13.Location = New-Object System.Drawing.Point(480, 16)
$Label13.Size = New-Object System.Drawing.Size(184, 15)
$Label13.TabIndex = 23
$Label13.Text = "Moved"
#~~< ListBox12 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox12 = New-Object System.Windows.Forms.ListBox
$ListBox12.FormattingEnabled = $true
$ListBox12.Location = New-Object System.Drawing.Point(808, 334)
$ListBox12.SelectedIndex = -1
$ListBox12.Size = New-Object System.Drawing.Size(17, 17)
$ListBox12.TabIndex = 22
$ListBox12.Visible = $false
#~~< Button22 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button22 = New-Object System.Windows.Forms.Button
$Button22.Location = New-Object System.Drawing.Point(715, 369)
$Button22.Size = New-Object System.Drawing.Size(110, 23)
$Button22.TabIndex = 21
$Button22.Text = "Search"
$Button22.UseVisualStyleBackColor = $true
$Button22.add_Click({Button22Click($Button22)})
#~~< TextBox4 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox4 = New-Object System.Windows.Forms.TextBox
$TextBox4.Location = New-Object System.Drawing.Point(253, 371)
$TextBox4.Size = New-Object System.Drawing.Size(456, 20)
$TextBox4.TabIndex = 20
$TextBox4.Text = "Search"
$TextBox4.add_MouseClick({TextBox4MouseClick($TextBox4)})
#~~< ComboBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComboBox1 = New-Object System.Windows.Forms.ComboBox
$ComboBox1.FormattingEnabled = $true
$ComboBox1.Location = New-Object System.Drawing.Point(253, 398)
$ComboBox1.SelectedIndex = -1
$ComboBox1.Size = New-Object System.Drawing.Size(572, 21)
$ComboBox1.TabIndex = 19
$ComboBox1.Text = ""
#~~< ProgressBar2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ProgressBar2 = New-Object System.Windows.Forms.ProgressBar
$ProgressBar2.Location = New-Object System.Drawing.Point(364, 426)
$ProgressBar2.Size = New-Object System.Drawing.Size(461, 23)
$ProgressBar2.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$ProgressBar2.TabIndex = 18
$ProgressBar2.Text = ""
#~~< Button21 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button21 = New-Object System.Windows.Forms.Button
$Button21.Location = New-Object System.Drawing.Point(253, 425)
$Button21.Size = New-Object System.Drawing.Size(105, 24)
$Button21.TabIndex = 17
$Button21.Text = "Move"
$Button21.UseVisualStyleBackColor = $true
$Button21.add_Click({Button21Click($Button21)})
#~~< ListBox9 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox9 = New-Object System.Windows.Forms.ListBox
$ListBox9.FormattingEnabled = $true
$ListBox9.Location = New-Object System.Drawing.Point(480, 34)
$ListBox9.SelectedIndex = -1
$ListBox9.Size = New-Object System.Drawing.Size(344, 290)
$ListBox9.TabIndex = 16
$ListBox9.add_DoubleClick({ListBox9DoubleClick($ListBox9)})
#~~< Label11 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label11 = New-Object System.Windows.Forms.Label
$Label11.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](238)))
$Label11.Location = New-Object System.Drawing.Point(7, 343)
$Label11.Size = New-Object System.Drawing.Size(99, 16)
$Label11.TabIndex = 15
$Label11.Text = "Not found in AD"
$Label11.ForeColor = [System.Drawing.Color]::Red
#~~< Label12 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label12 = New-Object System.Windows.Forms.Label
$Label12.Location = New-Object System.Drawing.Point(6, 16)
$Label12.Size = New-Object System.Drawing.Size(184, 15)
$Label12.TabIndex = 14
$Label12.Text = "Found"
#~~< ListBox10 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox10 = New-Object System.Windows.Forms.ListBox
$ListBox10.FormattingEnabled = $true
$ListBox10.Location = New-Object System.Drawing.Point(6, 366)
$ListBox10.SelectedIndex = -1
$ListBox10.Size = New-Object System.Drawing.Size(241, 82)
$ListBox10.TabIndex = 13
$ListBox10.add_DoubleClick({ListBox10DoubleClick($ListBox10)})
#~~< ListBox11 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox11 = New-Object System.Windows.Forms.ListBox
$ListBox11.FormattingEnabled = $true
$ListBox11.Location = New-Object System.Drawing.Point(6, 34)
$ListBox11.SelectedIndex = -1
$ListBox11.Size = New-Object System.Drawing.Size(468, 290)
$ListBox11.TabIndex = 12
$ListBox11.add_DoubleClick({ListBox11DoubleClick($ListBox11)})
$GroupBox8.Controls.Add($GroupBox9)
$GroupBox8.Controls.Add($Label13)
$GroupBox8.Controls.Add($ListBox12)
$GroupBox8.Controls.Add($Button22)
$GroupBox8.Controls.Add($TextBox4)
$GroupBox8.Controls.Add($ComboBox1)
$GroupBox8.Controls.Add($ProgressBar2)
$GroupBox8.Controls.Add($Button21)
$GroupBox8.Controls.Add($ListBox9)
$GroupBox8.Controls.Add($Label11)
$GroupBox8.Controls.Add($Label12)
$GroupBox8.Controls.Add($ListBox10)
$GroupBox8.Controls.Add($ListBox11)
#~~< GroupBox7 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox7 = New-Object System.Windows.Forms.GroupBox
$GroupBox7.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Center
$GroupBox7.Location = New-Object System.Drawing.Point(6, 6)
$GroupBox7.Size = New-Object System.Drawing.Size(200, 454)
$GroupBox7.TabIndex = 7
$GroupBox7.TabStop = $false
$GroupBox7.Text = "Computers"
#~~< Button19 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button19 = New-Object System.Windows.Forms.Button
$Button19.Location = New-Object System.Drawing.Point(6, 425)
$Button19.Size = New-Object System.Drawing.Size(188, 23)
$Button19.TabIndex = 1
$Button19.Text = "Load"
$Button19.UseVisualStyleBackColor = $true
$Button19.add_Click({Button19Click($Button19)})
#~~< RichTextBox2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RichTextBox2 = New-Object System.Windows.Forms.RichTextBox
$RichTextBox2.Location = New-Object System.Drawing.Point(6, 19)
$RichTextBox2.Size = New-Object System.Drawing.Size(188, 400)
$RichTextBox2.TabIndex = 0
$RichTextBox2.Text = "Fill in computer names in separate lines"
$RichTextBox2.add_MouseClick({RichTextBox2MouseClick($RichTextBox2)})
$GroupBox7.Controls.Add($Button19)
$GroupBox7.Controls.Add($RichTextBox2)
$TabPage5.Controls.Add($GroupBox8)
$TabPage5.Controls.Add($GroupBox7)
#~~< TabPage6 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabPage6 = New-Object System.Windows.Forms.TabPage
$TabPage6.Location = New-Object System.Drawing.Point(4, 22)
$TabPage6.Padding = New-Object System.Windows.Forms.Padding(3)
$TabPage6.Size = New-Object System.Drawing.Size(1048, 466)
$TabPage6.TabIndex = 5
$TabPage6.Text = "Adding new users"
$TabPage6.UseVisualStyleBackColor = $true
#~~< Button42 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button42 = New-Object System.Windows.Forms.Button
$Button42.Location = New-Object System.Drawing.Point(262, 418)
$Button42.Size = New-Object System.Drawing.Size(138, 41)
$Button42.TabIndex = 38
$Button42.Text = "Move to user modification"
$Button42.UseVisualStyleBackColor = $true
$Button42.add_Click({Button42Click($Button42)})
#~~< ListBox17 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox17 = New-Object System.Windows.Forms.ListBox
$ListBox17.FormattingEnabled = $true
$ListBox17.Location = New-Object System.Drawing.Point(221, 354)
$ListBox17.SelectedIndex = -1
$ListBox17.Size = New-Object System.Drawing.Size(19, 17)
$ListBox17.TabIndex = 37
$ListBox17.Visible = $false
#~~< GroupBox11 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox11 = New-Object System.Windows.Forms.GroupBox
$GroupBox11.Location = New-Object System.Drawing.Point(405, 6)
$GroupBox11.Size = New-Object System.Drawing.Size(637, 454)
$GroupBox11.TabIndex = 9
$GroupBox11.TabStop = $false
$GroupBox11.Text = "User data"
#~~< Button27 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button27 = New-Object System.Windows.Forms.Button
$Button27.Location = New-Object System.Drawing.Point(6, 334)
$Button27.Size = New-Object System.Drawing.Size(105, 24)
$Button27.TabIndex = 34
$Button27.Text = "Copy"
$Button27.UseVisualStyleBackColor = $true
$Button27.add_Click({Button27Click($Button27)})
#~~< Button26 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button26 = New-Object System.Windows.Forms.Button
$Button26.Location = New-Object System.Drawing.Point(565, 393)
$Button26.Size = New-Object System.Drawing.Size(66, 22)
$Button26.TabIndex = 33
$Button26.Text = "Search"
$Button26.UseVisualStyleBackColor = $true
$Button26.add_Click({Button26Click($Button26)})
#~~< Label16 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label16 = New-Object System.Windows.Forms.Label
$Label16.Location = New-Object System.Drawing.Point(283, 340)
$Label16.Size = New-Object System.Drawing.Size(78, 20)
$Label16.TabIndex = 32
$Label16.Text = "Search"
$Label16.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< Label15 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label15 = New-Object System.Windows.Forms.Label
$Label15.Location = New-Object System.Drawing.Point(8, 393)
$Label15.Size = New-Object System.Drawing.Size(71, 21)
$Label15.TabIndex = 31
$Label15.Text = "Localization"
$Label15.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< Label14 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label14 = New-Object System.Windows.Forms.Label
$Label14.Location = New-Object System.Drawing.Point(28, 366)
$Label14.Size = New-Object System.Drawing.Size(51, 21)
$Label14.TabIndex = 30
$Label14.Text = "Group"
$Label14.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< ComboBox3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComboBox3 = New-Object System.Windows.Forms.ComboBox
$ComboBox3.FormattingEnabled = $true
$ComboBox3.Location = New-Object System.Drawing.Point(85, 366)
$ComboBox3.SelectedIndex = -1
$ComboBox3.Size = New-Object System.Drawing.Size(474, 21)
$ComboBox3.TabIndex = 29
$ComboBox3.Text = ""
#~~< ListBox14 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox14 = New-Object System.Windows.Forms.ListBox
$ListBox14.FormattingEnabled = $true
$ListBox14.Location = New-Object System.Drawing.Point(6, 181)
$ListBox14.SelectedIndex = -1
$ListBox14.Size = New-Object System.Drawing.Size(625, 147)
$ListBox14.TabIndex = 28
#~~< ListBox13 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox13 = New-Object System.Windows.Forms.ListBox
$ListBox13.FormattingEnabled = $true
$ListBox13.Location = New-Object System.Drawing.Point(6, 19)
$ListBox13.SelectedIndex = -1
$ListBox13.Size = New-Object System.Drawing.Size(625, 160)
$ListBox13.TabIndex = 27
#~~< Button24 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button24 = New-Object System.Windows.Forms.Button
$Button24.Location = New-Object System.Drawing.Point(565, 366)
$Button24.Size = New-Object System.Drawing.Size(66, 21)
$Button24.TabIndex = 26
$Button24.Text = "Search"
$Button24.UseVisualStyleBackColor = $true
$Button24.add_Click({Button24Click($Button24)})
#~~< TextBox5 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox5 = New-Object System.Windows.Forms.TextBox
$TextBox5.Location = New-Object System.Drawing.Point(367, 340)
$TextBox5.Size = New-Object System.Drawing.Size(264, 20)
$TextBox5.TabIndex = 25
$TextBox5.Text = ""
#~~< ComboBox2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComboBox2 = New-Object System.Windows.Forms.ComboBox
$ComboBox2.FormattingEnabled = $true
$ComboBox2.Location = New-Object System.Drawing.Point(85, 393)
$ComboBox2.SelectedIndex = -1
$ComboBox2.Size = New-Object System.Drawing.Size(474, 21)
$ComboBox2.TabIndex = 24
$ComboBox2.Text = ""
#~~< ProgressBar3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ProgressBar3 = New-Object System.Windows.Forms.ProgressBar
$ProgressBar3.Location = New-Object System.Drawing.Point(117, 421)
$ProgressBar3.Size = New-Object System.Drawing.Size(516, 23)
$ProgressBar3.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$ProgressBar3.TabIndex = 23
$ProgressBar3.Text = ""
#~~< Button25 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button25 = New-Object System.Windows.Forms.Button
$Button25.Location = New-Object System.Drawing.Point(6, 420)
$Button25.Size = New-Object System.Drawing.Size(105, 24)
$Button25.TabIndex = 22
$Button25.Text = "Add"
$Button25.UseVisualStyleBackColor = $true
$Button25.add_Click({Button25Click($Button25)})
$GroupBox11.Controls.Add($Button27)
$GroupBox11.Controls.Add($Button26)
$GroupBox11.Controls.Add($Label16)
$GroupBox11.Controls.Add($Label15)
$GroupBox11.Controls.Add($Label14)
$GroupBox11.Controls.Add($ComboBox3)
$GroupBox11.Controls.Add($ListBox14)
$GroupBox11.Controls.Add($ListBox13)
$GroupBox11.Controls.Add($Button24)
$GroupBox11.Controls.Add($TextBox5)
$GroupBox11.Controls.Add($ComboBox2)
$GroupBox11.Controls.Add($ProgressBar3)
$GroupBox11.Controls.Add($Button25)
#~~< ListBox16 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox16 = New-Object System.Windows.Forms.ListBox
$ListBox16.FormattingEnabled = $true
$ListBox16.Location = New-Object System.Drawing.Point(196, 354)
$ListBox16.SelectedIndex = -1
$ListBox16.Size = New-Object System.Drawing.Size(19, 17)
$ListBox16.TabIndex = 36
$ListBox16.Visible = $false
#~~< Button28 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button28 = New-Object System.Windows.Forms.Button
$Button28.Location = New-Object System.Drawing.Point(262, 374)
$Button28.Size = New-Object System.Drawing.Size(138, 43)
$Button28.TabIndex = 35
$Button28.Text = "Move to password reset"
$Button28.UseVisualStyleBackColor = $true
$Button28.add_Click({Button28Click($Button28)})
#~~< Label17 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label17 = New-Object System.Windows.Forms.Label
$Label17.Location = New-Object System.Drawing.Point(6, 6)
$Label17.Size = New-Object System.Drawing.Size(122, 20)
$Label17.TabIndex = 4
$Label17.Text = "ID"
#~~< Label20 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label20 = New-Object System.Windows.Forms.Label
$Label20.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](238)))
$Label20.Location = New-Object System.Drawing.Point(5, 358)
$Label20.Size = New-Object System.Drawing.Size(169, 16)
$Label20.TabIndex = 35
$Label20.Text = "Found"
$Label20.ForeColor = [System.Drawing.Color]::Red
#~~< RichTextBox3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RichTextBox3 = New-Object System.Windows.Forms.RichTextBox
$RichTextBox3.Location = New-Object System.Drawing.Point(4, 29)
$RichTextBox3.Size = New-Object System.Drawing.Size(123, 316)
$RichTextBox3.TabIndex = 0
$RichTextBox3.Text = "Fill in ID's in separate lines"
$RichTextBox3.add_MouseClick({RichTextBox3MouseClick($RichTextBox3)})
#~~< ListBox15 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox15 = New-Object System.Windows.Forms.ListBox
$ListBox15.FormattingEnabled = $true
$ListBox15.Location = New-Object System.Drawing.Point(5, 377)
$ListBox15.SelectedIndex = -1
$ListBox15.Size = New-Object System.Drawing.Size(251, 82)
$ListBox15.TabIndex = 34
#~~< Button23 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button23 = New-Object System.Windows.Forms.Button
$Button23.Location = New-Object System.Drawing.Point(262, 349)
$Button23.Size = New-Object System.Drawing.Size(138, 23)
$Button23.TabIndex = 1
$Button23.Text = "Load"
$Button23.UseVisualStyleBackColor = $true
$Button23.add_Click({Button23Click($Button23)})
#~~< Label19 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label19 = New-Object System.Windows.Forms.Label
$Label19.Location = New-Object System.Drawing.Point(262, 6)
$Label19.Size = New-Object System.Drawing.Size(137, 20)
$Label19.TabIndex = 6
$Label19.Text = "Surname"
#~~< RichTextBox4 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RichTextBox4 = New-Object System.Windows.Forms.RichTextBox
$RichTextBox4.Location = New-Object System.Drawing.Point(133, 29)
$RichTextBox4.Size = New-Object System.Drawing.Size(122, 316)
$RichTextBox4.TabIndex = 2
$RichTextBox4.Text = "Fill in names in separate lines"
$RichTextBox4.add_MouseClick({RichTextBox4MouseClick($RichTextBox4)})
#~~< Label18 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label18 = New-Object System.Windows.Forms.Label
$Label18.Location = New-Object System.Drawing.Point(133, 6)
$Label18.Size = New-Object System.Drawing.Size(121, 20)
$Label18.TabIndex = 5
$Label18.Text = "Name"
#~~< RichTextBox5 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RichTextBox5 = New-Object System.Windows.Forms.RichTextBox
$RichTextBox5.Location = New-Object System.Drawing.Point(261, 29)
$RichTextBox5.Size = New-Object System.Drawing.Size(138, 316)
$RichTextBox5.TabIndex = 3
$RichTextBox5.Text = "Fill in surnames in separate lines"
$RichTextBox5.add_MouseClick({RichTextBox5MouseClick($RichTextBox5)})
$TabPage6.Controls.Add($Button42)
$TabPage6.Controls.Add($ListBox17)
$TabPage6.Controls.Add($GroupBox11)
$TabPage6.Controls.Add($ListBox16)
$TabPage6.Controls.Add($Button28)
$TabPage6.Controls.Add($Label17)
$TabPage6.Controls.Add($Label20)
$TabPage6.Controls.Add($RichTextBox3)
$TabPage6.Controls.Add($ListBox15)
$TabPage6.Controls.Add($Button23)
$TabPage6.Controls.Add($Label19)
$TabPage6.Controls.Add($RichTextBox4)
$TabPage6.Controls.Add($Label18)
$TabPage6.Controls.Add($RichTextBox5)
$TabPage6.add_Click({TabPage6Click($TabPage6)})
#~~< TabPage7 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabPage7 = New-Object System.Windows.Forms.TabPage
$TabPage7.Location = New-Object System.Drawing.Point(4, 22)
$TabPage7.Padding = New-Object System.Windows.Forms.Padding(3)
$TabPage7.Size = New-Object System.Drawing.Size(1048, 466)
$TabPage7.TabIndex = 6
$TabPage7.Text = "User modifications"
$TabPage7.UseVisualStyleBackColor = $true
#~~< CheckBox9 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox9 = New-Object System.Windows.Forms.CheckBox
$CheckBox9.CheckAlign = [System.Drawing.ContentAlignment]::MiddleRight
$CheckBox9.Enabled = $true
$CheckBox9.Location = New-Object System.Drawing.Point(469, 115)
$CheckBox9.Size = New-Object System.Drawing.Size(143, 24)
$CheckBox9.TabIndex = 67
$CheckBox9.Text = "Enabled"
$CheckBox9.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$CheckBox9.UseVisualStyleBackColor = $true
#~~< Button39 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button39 = New-Object System.Windows.Forms.Button
$Button39.Location = New-Object System.Drawing.Point(464, 263)
$Button39.Size = New-Object System.Drawing.Size(120, 43)
$Button39.TabIndex = 45
$Button39.Text = "Delete from selected group"
$Button39.UseVisualStyleBackColor = $true
$Button39.add_Click({Button39Click($Button39)})
#~~< GroupBox13 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox13 = New-Object System.Windows.Forms.GroupBox
$GroupBox13.Location = New-Object System.Drawing.Point(458, 312)
$GroupBox13.Size = New-Object System.Drawing.Size(584, 148)
$GroupBox13.TabIndex = 66
$GroupBox13.TabStop = $false
$GroupBox13.Text = "Moving"
#~~< Label25 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label25 = New-Object System.Windows.Forms.Label
$Label25.Location = New-Object System.Drawing.Point(233, 10)
$Label25.Size = New-Object System.Drawing.Size(78, 20)
$Label25.TabIndex = 41
$Label25.Text = "Search"
$Label25.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< ComboBox5 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComboBox5 = New-Object System.Windows.Forms.ComboBox
$ComboBox5.FormattingEnabled = $true
$ComboBox5.Location = New-Object System.Drawing.Point(74, 92)
$ComboBox5.SelectedIndex = -1
$ComboBox5.Size = New-Object System.Drawing.Size(435, 21)
$ComboBox5.TabIndex = 34
$ComboBox5.Text = ""
#~~< Button31 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button31 = New-Object System.Windows.Forms.Button
$Button31.Location = New-Object System.Drawing.Point(515, 36)
$Button31.Size = New-Object System.Drawing.Size(66, 21)
$Button31.TabIndex = 35
$Button31.Text = "Search"
$Button31.UseVisualStyleBackColor = $true
$Button31.add_Click({Button31Click($Button31)})
#~~< ComboBox4 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComboBox4 = New-Object System.Windows.Forms.ComboBox
$ComboBox4.FormattingEnabled = $true
$ComboBox4.Location = New-Object System.Drawing.Point(74, 36)
$ComboBox4.SelectedIndex = -1
$ComboBox4.Size = New-Object System.Drawing.Size(435, 21)
$ComboBox4.TabIndex = 36
$ComboBox4.Text = ""
#~~< Label24 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label24 = New-Object System.Windows.Forms.Label
$Label24.Location = New-Object System.Drawing.Point(17, 36)
$Label24.Size = New-Object System.Drawing.Size(51, 21)
$Label24.TabIndex = 37
$Label24.Text = "Group"
$Label24.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< Label23 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label23 = New-Object System.Windows.Forms.Label
$Label23.Location = New-Object System.Drawing.Point(-3, 93)
$Label23.Size = New-Object System.Drawing.Size(71, 21)
$Label23.TabIndex = 38
$Label23.Text = "Localization"
$Label23.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< Button30 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button30 = New-Object System.Windows.Forms.Button
$Button30.Location = New-Object System.Drawing.Point(515, 92)
$Button30.Size = New-Object System.Drawing.Size(66, 22)
$Button30.TabIndex = 39
$Button30.Text = "Search"
$Button30.UseVisualStyleBackColor = $true
$Button30.add_Click({Button30Click($Button30)})
#~~< TextBox6 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox6 = New-Object System.Windows.Forms.TextBox
$TextBox6.Location = New-Object System.Drawing.Point(317, 10)
$TextBox6.Size = New-Object System.Drawing.Size(264, 20)
$TextBox6.TabIndex = 40
$TextBox6.Text = ""
#~~< Button32 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button32 = New-Object System.Windows.Forms.Button
$Button32.Location = New-Object System.Drawing.Point(374, 119)
$Button32.Size = New-Object System.Drawing.Size(207, 23)
$Button32.TabIndex = 42
$Button32.Text = "Move to selected localization"
$Button32.UseVisualStyleBackColor = $true
$Button32.add_Click({Button32Click($Button32)})
#~~< Button33 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button33 = New-Object System.Windows.Forms.Button
$Button33.Location = New-Object System.Drawing.Point(161, 63)
$Button33.Size = New-Object System.Drawing.Size(207, 23)
$Button33.TabIndex = 43
$Button33.Text = "Add to selected group"
$Button33.UseVisualStyleBackColor = $true
$Button33.add_Click({Button33Click($Button33)})
#~~< Button34 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button34 = New-Object System.Windows.Forms.Button
$Button34.Location = New-Object System.Drawing.Point(374, 63)
$Button34.Size = New-Object System.Drawing.Size(207, 23)
$Button34.TabIndex = 44
$Button34.Text = "Remove from selected group"
$Button34.UseVisualStyleBackColor = $true
$Button34.add_Click({Button34Click($Button34)})
$GroupBox13.Controls.Add($Label25)
$GroupBox13.Controls.Add($ComboBox5)
$GroupBox13.Controls.Add($Button31)
$GroupBox13.Controls.Add($ComboBox4)
$GroupBox13.Controls.Add($Label24)
$GroupBox13.Controls.Add($Label23)
$GroupBox13.Controls.Add($Button30)
$GroupBox13.Controls.Add($TextBox6)
$GroupBox13.Controls.Add($Button32)
$GroupBox13.Controls.Add($Button33)
$GroupBox13.Controls.Add($Button34)
#~~< Label34 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label34 = New-Object System.Windows.Forms.Label
$Label34.Location = New-Object System.Drawing.Point(591, 159)
$Label34.Size = New-Object System.Drawing.Size(448, 75)
$Label34.TabIndex = 65
$Label34.Text = "..."
#~~< Label33 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label33 = New-Object System.Windows.Forms.Label
$Label33.Location = New-Object System.Drawing.Point(464, 159)
$Label33.Size = New-Object System.Drawing.Size(121, 20)
$Label33.TabIndex = 64
$Label33.Text = "Localization"
$Label33.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< Button37 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button37 = New-Object System.Windows.Forms.Button
$Button37.Location = New-Object System.Drawing.Point(888, 7)
$Button37.Size = New-Object System.Drawing.Size(154, 46)
$Button37.TabIndex = 63
$Button37.Text = "Move to password reset"
$Button37.UseVisualStyleBackColor = $true
$Button37.add_Click({Button37Click($Button37)})
#~~< Button36 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button36 = New-Object System.Windows.Forms.Button
$Button36.Location = New-Object System.Drawing.Point(888, 59)
$Button36.Size = New-Object System.Drawing.Size(154, 26)
$Button36.TabIndex = 62
$Button36.Text = "Refresh"
$Button36.UseVisualStyleBackColor = $true
$Button36.add_Click({Button36Click($Button36)})
#~~< Button35 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button35 = New-Object System.Windows.Forms.Button
$Button35.Location = New-Object System.Drawing.Point(888, 91)
$Button35.Size = New-Object System.Drawing.Size(154, 65)
$Button35.TabIndex = 61
$Button35.Text = "Save"
$Button35.UseVisualStyleBackColor = $true
$Button35.add_Click({Button35Click($Button35)})
#~~< Label32 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label32 = New-Object System.Windows.Forms.Label
$Label32.Location = New-Object System.Drawing.Point(464, 237)
$Label32.Size = New-Object System.Drawing.Size(120, 20)
$Label32.TabIndex = 60
$Label32.Text = "Group member"
$Label32.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< ListBox20 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox20 = New-Object System.Windows.Forms.ListBox
$ListBox20.FormattingEnabled = $true
$ListBox20.Location = New-Object System.Drawing.Point(591, 237)
$ListBox20.SelectedIndex = -1
$ListBox20.Size = New-Object System.Drawing.Size(448, 69)
$ListBox20.TabIndex = 59
#~~< CheckBox8 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox8 = New-Object System.Windows.Forms.CheckBox
$CheckBox8.CheckAlign = [System.Drawing.ContentAlignment]::MiddleRight
$CheckBox8.Location = New-Object System.Drawing.Point(746, 115)
$CheckBox8.Size = New-Object System.Drawing.Size(136, 24)
$CheckBox8.TabIndex = 58
$CheckBox8.Text = "Force password change"
$CheckBox8.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$CheckBox8.UseVisualStyleBackColor = $true
#~~< CheckBox7 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox7 = New-Object System.Windows.Forms.CheckBox
$CheckBox7.CheckAlign = [System.Drawing.ContentAlignment]::MiddleRight
$CheckBox7.Enabled = $false
$CheckBox7.Location = New-Object System.Drawing.Point(619, 115)
$CheckBox7.Size = New-Object System.Drawing.Size(121, 24)
$CheckBox7.TabIndex = 57
$CheckBox7.Text = "Unlock"
$CheckBox7.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$CheckBox7.UseVisualStyleBackColor = $true
#~~< Label31 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label31 = New-Object System.Windows.Forms.Label
$Label31.Location = New-Object System.Drawing.Point(469, 33)
$Label31.Size = New-Object System.Drawing.Size(57, 20)
$Label31.TabIndex = 56
$Label31.Text = "Login"
$Label31.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< TextBox12 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox12 = New-Object System.Windows.Forms.TextBox
$TextBox12.Location = New-Object System.Drawing.Point(534, 34)
$TextBox12.Size = New-Object System.Drawing.Size(116, 20)
$TextBox12.TabIndex = 55
$TextBox12.Text = ""
#~~< Label30 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label30 = New-Object System.Windows.Forms.Label
$Label30.Location = New-Object System.Drawing.Point(461, 62)
$Label30.Size = New-Object System.Drawing.Size(67, 20)
$Label30.TabIndex = 54
$Label30.Text = "E-mail"
$Label30.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< TextBox11 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox11 = New-Object System.Windows.Forms.TextBox
$TextBox11.Location = New-Object System.Drawing.Point(534, 63)
$TextBox11.Size = New-Object System.Drawing.Size(348, 20)
$TextBox11.TabIndex = 53
$TextBox11.Text = ""
#~~< Label29 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label29 = New-Object System.Windows.Forms.Label
$Label29.Location = New-Object System.Drawing.Point(667, 34)
$Label29.Size = New-Object System.Drawing.Size(55, 20)
$Label29.TabIndex = 52
$Label29.Text = "Description"
$Label29.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< TextBox10 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox10 = New-Object System.Windows.Forms.TextBox
$TextBox10.Location = New-Object System.Drawing.Point(728, 34)
$TextBox10.Size = New-Object System.Drawing.Size(154, 20)
$TextBox10.TabIndex = 51
$TextBox10.Text = ""
#~~< Label28 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label28 = New-Object System.Windows.Forms.Label
$Label28.Location = New-Object System.Drawing.Point(455, 89)
$Label28.Size = New-Object System.Drawing.Size(110, 20)
$Label28.TabIndex = 50
$Label28.Text = "Display name"
$Label28.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< TextBox9 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox9 = New-Object System.Windows.Forms.TextBox
$TextBox9.Location = New-Object System.Drawing.Point(571, 89)
$TextBox9.Size = New-Object System.Drawing.Size(311, 20)
$TextBox9.TabIndex = 49
$TextBox9.Text = ""
#~~< Label27 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label27 = New-Object System.Windows.Forms.Label
$Label27.Location = New-Object System.Drawing.Point(656, 8)
$Label27.Size = New-Object System.Drawing.Size(66, 20)
$Label27.TabIndex = 48
$Label27.Text = "Surname"
$Label27.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< TextBox8 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox8 = New-Object System.Windows.Forms.TextBox
$TextBox8.Location = New-Object System.Drawing.Point(728, 7)
$TextBox8.Size = New-Object System.Drawing.Size(154, 20)
$TextBox8.TabIndex = 47
$TextBox8.Text = ""
$TextBox8.add_KeyUp({TextBox8KeyUp($TextBox8)})
#~~< Label26 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label26 = New-Object System.Windows.Forms.Label
$Label26.Location = New-Object System.Drawing.Point(471, 7)
$Label26.Size = New-Object System.Drawing.Size(55, 20)
$Label26.TabIndex = 46
$Label26.Text = "Name"
$Label26.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
#~~< TextBox7 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox7 = New-Object System.Windows.Forms.TextBox
$TextBox7.Location = New-Object System.Drawing.Point(534, 8)
$TextBox7.Size = New-Object System.Drawing.Size(116, 20)
$TextBox7.TabIndex = 45
$TextBox7.Text = ""
$TextBox7.add_KeyUp({TextBox7KeyUp($TextBox7)})
#~~< Label22 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label22 = New-Object System.Windows.Forms.Label
$Label22.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](238)))
$Label22.Location = New-Object System.Drawing.Point(212, 307)
$Label22.Size = New-Object System.Drawing.Size(99, 16)
$Label22.TabIndex = 11
$Label22.Text = "Not found in AD"
$Label22.ForeColor = [System.Drawing.Color]::Red
#~~< ListBox19 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox19 = New-Object System.Windows.Forms.ListBox
$ListBox19.FormattingEnabled = $true
$ListBox19.Location = New-Object System.Drawing.Point(211, 326)
$ListBox19.SelectedIndex = -1
$ListBox19.Size = New-Object System.Drawing.Size(241, 134)
$ListBox19.TabIndex = 10
#~~< Label21 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label21 = New-Object System.Windows.Forms.Label
$Label21.Location = New-Object System.Drawing.Point(212, 6)
$Label21.Size = New-Object System.Drawing.Size(184, 15)
$Label21.TabIndex = 9
$Label21.Text = "Found"
#~~< ListBox18 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListBox18 = New-Object System.Windows.Forms.ListBox
$ListBox18.FormattingEnabled = $true
$ListBox18.Location = New-Object System.Drawing.Point(212, 24)
$ListBox18.SelectedIndex = -1
$ListBox18.Size = New-Object System.Drawing.Size(240, 277)
$ListBox18.TabIndex = 8
$ListBox18.add_Click({ListBox18Click($ListBox18)})
#~~< GroupBox12 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GroupBox12 = New-Object System.Windows.Forms.GroupBox
$GroupBox12.Location = New-Object System.Drawing.Point(6, 6)
$GroupBox12.Size = New-Object System.Drawing.Size(200, 454)
$GroupBox12.TabIndex = 7
$GroupBox12.TabStop = $false
$GroupBox12.Text = "ID's"
#~~< Button29 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button29 = New-Object System.Windows.Forms.Button
$Button29.Location = New-Object System.Drawing.Point(6, 424)
$Button29.Size = New-Object System.Drawing.Size(188, 23)
$Button29.TabIndex = 1
$Button29.Text = "Load"
$Button29.UseVisualStyleBackColor = $true
$Button29.add_Click({Button29Click($Button29)})
#~~< RichTextBox6 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RichTextBox6 = New-Object System.Windows.Forms.RichTextBox
$RichTextBox6.Location = New-Object System.Drawing.Point(6, 19)
$RichTextBox6.Size = New-Object System.Drawing.Size(188, 399)
$RichTextBox6.TabIndex = 0
$RichTextBox6.Text = "Fill in ID's in separate lines"
$RichTextBox6.add_MouseClick({RichTextBox6MouseClick($RichTextBox6)})
$GroupBox12.Controls.Add($Button29)
$GroupBox12.Controls.Add($RichTextBox6)
$TabPage7.Controls.Add($CheckBox9)
$TabPage7.Controls.Add($Button39)
$TabPage7.Controls.Add($GroupBox13)
$TabPage7.Controls.Add($Label34)
$TabPage7.Controls.Add($Label33)
$TabPage7.Controls.Add($Button37)
$TabPage7.Controls.Add($Button36)
$TabPage7.Controls.Add($Button35)
$TabPage7.Controls.Add($Label32)
$TabPage7.Controls.Add($ListBox20)
$TabPage7.Controls.Add($CheckBox8)
$TabPage7.Controls.Add($CheckBox7)
$TabPage7.Controls.Add($Label31)
$TabPage7.Controls.Add($TextBox12)
$TabPage7.Controls.Add($Label30)
$TabPage7.Controls.Add($TextBox11)
$TabPage7.Controls.Add($Label29)
$TabPage7.Controls.Add($TextBox10)
$TabPage7.Controls.Add($Label28)
$TabPage7.Controls.Add($TextBox9)
$TabPage7.Controls.Add($Label27)
$TabPage7.Controls.Add($TextBox8)
$TabPage7.Controls.Add($Label26)
$TabPage7.Controls.Add($TextBox7)
$TabPage7.Controls.Add($Label22)
$TabPage7.Controls.Add($ListBox19)
$TabPage7.Controls.Add($Label21)
$TabPage7.Controls.Add($ListBox18)
$TabPage7.Controls.Add($GroupBox12)
$TabPage7.add_Click({TabPage7Click($TabPage7)})
$TabControl1.Controls.Add($TabPage1)
$TabControl1.Controls.Add($TabPage4)
$TabControl1.Controls.Add($TabPage5)
$TabControl1.Controls.Add($TabPage6)
$TabControl1.Controls.Add($TabPage7)
$TabControl1.SelectedIndex = 0
#~~< Button3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Point(939, 669)
$Button3.Size = New-Object System.Drawing.Size(116, 23)
$Button3.TabIndex = 3
$Button3.Text = "Close"
$Button3.UseVisualStyleBackColor = $true
$Button3.add_Click({Button3OnClick($Button3)})
$Form1.Controls.Add($Button41)
$Form1.Controls.Add($Button40)
$Form1.Controls.Add($RichTextBox7)
$Form1.Controls.Add($Button38)
$Form1.Controls.Add($Label35)
$Form1.Controls.Add($TabControl1)
$Form1.Controls.Add($Button3)

#endregion

#region Custom Code
$TextBox3.add_KeyDown({if($_.KeyCode -eq "Enter") {Button15Click($Button15)} })
$TextBox4.add_KeyDown({if($_.KeyCode -eq "Enter") {Button22Click($Button22)} })

pconsole(">>>>>>>>>>>>>>>>New sesion of console<<<<<<<<<<<<<<<<")
pconsole("Initialization of AD Organization Units")
$listOfOUs = Get-ADOrganizationalUnit -filter "Name -like '*'" | select -ExpandProperty DistinguishedName
    $ComboBox2.Items.AddRange($listOfOUs)
    $ComboBox1.Items.AddRange($listOfOUs)
    $ComboBox5.Items.AddRange($listOfOUs)
pconsole("Initialization of AD Groups")
$listOfGroups = Get-ADGroup -filter "Name -like '*'" | select -ExpandProperty Name
    $ComboBox3.Items.AddRange($listOfGroups)
    $ComboBox4.Items.AddRange($listOfGroups)

$listOfUsersFoundInAD=@()
$listOfAddedUsers = @()
$listOfPasswordResetUsers = @()
$listOfModifiedUsers = @()
$selectedModifiedUser = ""


#endregion

#region Event Loop

function Main{
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($Form1)
}

#endregion

#endregion





#region Event Handlers
$Label1.Text = passwordGenerator
pconsole("New password generated "+$Label1.Text)
function Button3OnClick( $object ){
	$Form1.Close()
}

function Button2OnClick( $object ){
	try{
        [Windows.Forms.Clipboard]::SetText($Label1.Text)
        pconsole("New password copied to clipboard - "+$Label1.Text)
    } catch {
        pconsole($omfg)
    }
}

function Button1OnClick( $object ){
	$pass = passwordGenerator
	pconsole("New password generated "+$pass)
	$Label1.Text = $pass
}

function Button4Click( $object ){
    #Load
	$ListBox1.Items.Clear()
	$ListBox2.Items.Clear()
	$global:listOfPasswordResetUsers = @()
    try{
	    [array]$loadedUserIds = $RichTextBox1.Text.Split("`n") | % { $_.trim() }
    } catch {
        pconsole($omfg)
    }
    $uniqueIdis = New-Object System.Collections.ArrayList
    $loadedUserIds | Select-Object -Unique | foreach {$uniqueIdis.Add($_) }

	foreach ($val in $loadedUserIds) { 
		pconsole($val)
        if($val -ne ""){
            if($uniqueIdis -contains $val){
                if($nameUser = getUserInfoAD($val)){
			        $ListBox1.Items.Add("ID: " + $val + ", Name: " + $nameUser)
			        $global:listOfPasswordResetUsers += $val
                } else {
        	        $ListBox2.Items.Add($val)
                }
                $uniqueIdis.Remove($val)
            }
        }
		
	}
    if($global:listOfPasswordResetUsers.Count -gt 0 -or $ListBox2.Items.Count -gt 0){
        pconsole("Loaded data")
    }
}


function Button5Click($object)
{
    #password reset
	$ListBox3.Items.Clear()
	$ProgressBar1.Value = 0
	$ProgressBar1.Maximum = $global:listOfPasswordResetUsers.Count
    $i = 0
	foreach ($item in $ListBox1.Items)
	{ 
        $login = $global:listOfPasswordResetUsers[$i]
		$passReset = resetPassword $login $($CheckBox1.Checked)
        if($passReset -eq $omfg){
            $ListBox3.Items.Add("Error, no permissions to object " + $item)
            pconsole("Error, no permissions to object " + $item)
        } else {
            $ListBox3.Items.Add($item + ", Password: " + $passReset[1])
            pconsole("New password for "+$item+", Password: " + $passReset[1])
        }
		$ProgressBar1.PerformStep()
        $i++
	}
	
}

function RichTextBox1MouseClick( $object ){
	
	if ($RichTextBox1.Text.Contains("Fill in")){
	    $RichTextBox1.Clear() 
    }
}

function Button10Click( $object ){
    copyToClipboard($ListBox3)
}

function Button15Click( $object ){
    $ListBox8.Items.Add("Results for phrase: "+ $TextBox3.Text)
    
    if ($CheckBox3.Checked){
        try{
            [array]$data=getComputerInfo($TextBox3.Text)
            pconsole("Loaded data")
        } catch {
            pconsole($omfg)
        }
        if($data){
            [string]$CanonicalName=$data.CanonicalName
            [string]$Created=$data.Created
            [string]$IPv4Address=$data.IPv4Address
            [string]$LastLogonDate=$data.LastLogonDate
            [string]$dns=$data.DNSHostName
            [string]$system=$data.OperatingSystem
            [string]$systemv=$data.OperatingSystemVersion
            $ListBox8.Items.Add("Computer: "+ $CanonicalName)
            $ListBox8.Items.Add("Last contact with AD: "+ $LastLogonDate)
            $ListBox8.Items.Add("Added: "+ $Created)
            $ListBox8.Items.Add("IP: "+ $IPv4Address)
            $ListBox8.Items.Add("DNS: "+ $dns)
            $ListBox8.Items.Add("System: "+ $system + ", version: "+$systemv)
            $ListBox8.Items.Add("======================================================")
            $ListBox8.Items.Add("")
         } else {
            $ListBox8.Items.Add("Not found")
            $ListBox8.Items.Add("")
         }
    } elseif($CheckBox2.Checked){
        try{
            [array]$data=getUserInfoADMore($TextBox3.Text)
            [array]$data2=getUserInfoByNameMore($TextBox3.Text)
            pconsole("Loaded data")
        } catch {
            pconsole($omfg)
        }
        if($data){
            [string]$CanonicalName=$data.CanonicalName
            [string]$nameSurname=$data.CN
            [string]$Created=$data.Created
            [string]$mail=$data.EmailAddress
            [string]$BadPassword=$data.LastBadPasswordAttempt
            [string]$LastLogon=$data.LastLogonDate
            $MemberOf=($data.MemberOf -split $dcvariable+'\s+')
            $passLastSet=$data.PasswordLastSet
            $ListBox8.Items.Add("User: "+ $nameSurname)
            $ListBox8.Items.Add("User location: "+ $CanonicalName)
            $ListBox8.Items.Add("Created: "+ $Created)
            $ListBox8.Items.Add("E-mail: "+ $mail)
            $ListBox8.Items.Add("Last password set: "+$passLastSet)
            $ListBox8.Items.Add("Last bad password: "+ $BadPassword)
            $ListBox8.Items.Add("Last logon: "+ $LastLogon)
            $ListBox8.Items.Add("Group member:")
            foreach ($val in $MemberOf){
                $ListBox8.Items.Add($val)
            }
            $ListBox8.Items.Add("======================================================")
            $ListBox8.Items.Add("")
        } else {
            $ListBox8.Items.Add("Not found by ID")
            $ListBox8.Items.Add("")
        }

        if($data2){
            foreach ($user in $data2)
            {
            [string]$identyfikator=$user.SamAccountName
            [string]$CanonicalName=$user.CanonicalName
            [string]$nameSurname=$user.CN
            [string]$Created=$user.Created
            [string]$mail=$user.EmailAddress
            [string]$BadPassword=$user.LastBadPasswordAttempt
            [string]$LastLogon=$user.LastLogonDate
            $MemberOf=($user.MemberOf -split $dcvariable+'\s+')
            $passLastSet=$user.PasswordLastSet
            $ListBox8.Items.Add("User: "+ $nameSurname)
            $ListBox8.Items.Add("ID: "+ $identyfikator)
            $ListBox8.Items.Add("User location: "+ $CanonicalName)
            $ListBox8.Items.Add("Created: "+ $Created)
            $ListBox8.Items.Add("E-mail: "+ $mail)
            $ListBox8.Items.Add("Last password set: "+$passLastSet)
            $ListBox8.Items.Add("Last bad password: "+ $BadPassword)
            $ListBox8.Items.Add("Last logon: "+ $LastLogon)
            $ListBox8.Items.Add("Group member:")
            foreach ($val in $MemberOf){
                $ListBox8.Items.Add($val)
            }
            $ListBox8.Items.Add("======================================================")
            $ListBox8.Items.Add("")
            }
        } else {
            $ListBox8.Items.Add("Not found by name or surname")
            $ListBox8.Items.Add("")
        }
    }

    $ListBox8.Items.Add('$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$')
    $ListBox8.Items.Add("")
}

function ListBox8DoubleClick( $object ){
    
    $message=""
    $message += $ListBox8.SelectedItem
    try
    {
        [Windows.Forms.Clipboard]::SetText($message)
        pconsole("Copied")
    } 
    catch 
    {
        pconsole("Field is empty!")
    }
    
}

function CheckBox3Click( $object ){
    $CheckBox2.Checked=$false
    $CheckBox3.Checked=$true
}

function CheckBox2Click( $object ){
    $CheckBox3.Checked=$false
    $CheckBox2.Checked=$true
}

function Button17Click( $object ){
    #Clear
    $ListBox8.Items.Clear()
}

function Button18Click( $object ){
    #copy
    $message=""
    foreach ($item in $ListBox8.Items){
		$message += "$item `n"
	}
    try {
        [Windows.Forms.Clipboard]::SetText($message)
        pconsole("Copied to clipboard")
    } 
    catch {
        pconsole($omfg)
    }
}


function Button21Click( $object ){
    #move
    $ListBox9.Items.Clear()
    $ProgressBar2.Value = 0
    if ($CheckBox5.Checked)
    {
        #move
	    $ProgressBar2.Maximum = $ListBox12.Items.Count
        if ($listOfOUs -contains $ComboBox1.SelectedItem) {
            foreach ($item in $ListBox12.Items){
                try{
                    move-ADObject -Identity (Get-ADComputer $item).objectguid -TargetPath $ComboBox1.SelectedItem
                    pconsole("Computer "+$item+" moved")
                    pconsole( "Change will take up to 30 seconds")
                } catch {
                    pconsole($omfg)
                }
                $ListBox9.Items.Add($item)
                $ProgressBar2.PerformStep()
            }
        } else {
            $msg = "Choose correct OU Unit"
            $ListBox9.Items.Add($msg)
            pconsole($msg)
        }
    } elseif ($CheckBox4.Checked)
    {
        #Add
        $ProgressBar2.Maximum = $ListBox10.Items.Count
        if ($listOfOUs -contains $ComboBox1.SelectedItem) {
            foreach ($item in $ListBox10.Items){
                try{
                    New-ADComputer -Name $item -Path $ComboBox1.SelectedItem
                    pconsole("Computer "+$item+" added")
                    pconsole( "Change will take up to 30 seconds")
                } catch {
                    pconsole($omfg)
                }
                $ListBox9.Items.Add($item)
                $ProgressBar2.PerformStep()
            }
        } else {
            $msg = "Choose correct OU Unit"
            $ListBox9.Items.Add($msg)
            pconsole($msg)
        }
    } elseif ($CheckBox6.Checked)
    {
        #Delete
        $ProgressBar2.Maximum = $ListBox12.Items.Count
        foreach ($item in $ListBox12.Items){
            try {
                Remove-ADComputer -Identity $item -Confirm:$false -ErrorAction continue
                pconsole("Computer "+$item+" deleted")
                pconsole( "Change will take up to 30 seconds")
            } catch {
                pconsole($omfg)
            }
            $ListBox9.Items.Add($item)
            $ProgressBar2.PerformStep()
        }
    }
    
}

function Button19Click( $object ){
    #Load
    $ListBox11.Items.Clear()
	$ListBox10.Items.Clear()
    $ListBox12.Items.Clear()
    $ListBox9.Items.Clear()
	[array]$loadedComps = $RichTextBox2.Text.Split("`n") | % { $_.trim() }
	foreach ($val in $loadedComps) {
        if($val -ne ""){
		    pconsole($val)
            if($compOU = getComputerInfoLess($val)){
			    $ListBox11.Items.Add($compOU)
                $ListBox12.Items.Add($val) #hidden list
            } else {
        	    $ListBox10.Items.Add($val)
            }
        }
	}
    if($ListBox12.Items.Count -gt 0 -or $ListBox10.Items.Count -gt 0){
        pconsole("Data loaded")
    }
}

function Button22Click( $object ){
    #search
    if($TextBox4.Text){
        $ComboBox1.Items.Clear()
        foreach($item in $listOfOUs){
            if($item.ToLower().Contains($TextBox4.Text.ToLower())){
                $ComboBox1.Items.Add($item)
            }
        }
        pconsole("List updated")
    }
}

function RichTextBox2MouseClick( $object ){
    if ($RichTextBox2.Text.Contains("Fill in"))
	{
	$RichTextBox2.Clear() }
}

function CheckBox4Click( $object ){
    $CheckBox5.Checked=$false
    $CheckBox6.Checked=$false
    $CheckBox4.Checked=$true
    $GroupBox8.Text = $Label13.Text = "Adding"
    $Button21.Text = "Add"
}

function CheckBox5Click( $object ){
    $CheckBox4.Checked=$false
    $CheckBox6.Checked=$false
    $CheckBox5.Checked=$true
    $GroupBox8.Text = $Label13.Text = "Moving"
    $Button21.Text = "Move"
}

function CheckBox6Click( $object ){
    $CheckBox4.Checked=$false
    $CheckBox6.Checked=$true
    $CheckBox5.Checked=$false
    $GroupBox8.Text = $Label13.Text = "Deleting"
    $Button21.Text = "Delete"
}

function ListBox9DoubleClick( $object ){
    $message=""
    foreach ($item in $ListBox9.Items)
	{
		$message += "$item `n"
	}
    try
    {
        [Windows.Forms.Clipboard]::SetText($message)
        pconsole("Copied")
    } 
    catch 
    {
        pconsole("Field empty!")
    }
    
}

function ListBox11DoubleClick( $object ){
    $message=""
    foreach ($item in $ListBox11.Items)
	{
		$message += "$item `n"
	}
    try
    {
        [Windows.Forms.Clipboard]::SetText($message)
        pconsole("Copied")
    } 
    catch 
    {
        pconsole("Field empty!")
    }
    
}

function ListBox10DoubleClick( $object ){
    $message=""
    foreach ($item in $ListBox10.Items)
	{
		$message += "$item `n"
	}
    try
    {
        [Windows.Forms.Clipboard]::SetText($message)
        pconsole("Copied")
    } 
    catch 
    {
        pconsole("Field empty!")
    }
    
}

function TextBox4MouseClick( $object ){
    if ($TextBox4.Text.Contains("Search"))
	{
	$TextBox4.Clear() }
}

function Button24Click( $object ){
    # search
    if($TextBox5.Text){
        $ComboBox3.Items.Clear()
        foreach($item in $listOfGroups){
            if($item.ToLower().Contains($TextBox5.Text.ToLower())){
                $ComboBox3.Items.Add($item)
            }
        }
        pconsole("List updated")
    }
}

function Button25Click( $object ){
    # add user
    $ListBox14.Items.Clear()
	$ProgressBar3.Value = 0
	$ProgressBar3.Maximum = $global:listOfAddedUsers.Length
    if(-not ($ComboBox3.SelectedItem -eq $null -and $ComboBox2.SelectedItem -eq $null)) {
        if (-not ($listOfOUs -contains $ComboBox2.SelectedItem)){
            $ListBox14.Items.Add("Choose correct OU Unit!")
        } elseif (-not ($listOfGroups -contains $ComboBox3.SelectedItem)){
            $ListBox14.Items.Add("Choose correct group!")
        } else {
            foreach ($item in $global:listOfAddedUsers)
	        {
                pconsole("Processing: "+$item)
                $password = passwordGenerator
	            $adPassword = ConvertTo-SecureString -String $password -Force -AsPlainText
                try{
                    New-ADUser ($item.Name+" "+$item.Surname) -GivenName $item.Name -Surname $item.Surname -DisplayName ($item.Name+" "+$item.Surname) -Description $item.Id -Country "PL" -AccountPassword $adPassword -UserPrincipalName ($($item.Id)+$atvariable) -SamAccountName $item.Id -Enabled $true -Path $ComboBox2.SelectedItem -ChangePasswordAtLogon $true -ErrorAction Stop
                } catch {
                    pconsole($omfg)
                    $errorAddingUser=1
                }
                if($errorAddingUser -ne 1){
                    try{
                        Add-ADGroupMember -Identity $ComboBox3.SelectedItem -Member $item.Id -ErrorAction Stop
                    } catch {
                        pconsole($omfg)
                    }
		            $ListBox14.Items.Add("Id: "+$item.Id + " Name: "+$item.Name + " Surname: "+$item.Surname+" Password: "+$password)
                    pconsole("User added ID: "+$item.Id + " Name: "+$item.Name + " Surname: "+$item.Surname+" Password: "+$password)
                } else {
                    $ListBox14.Items.Add("Error adding user Id:"+$item.Id+" probably user with tah name and surname is already in AD!")
                    pconsole("Error adding user Id:"+$item.Id+" probably user with tah name and surname is already in AD!")
                }
                $ProgressBar3.PerformStep()
                $errorAddingUser=0
	        }
        $global:listOfAddedUsers = @()
        $ListBox13.Items.Clear()
        }
    } elseif($global:listOfAddedUsers.Length -eq 0){
        $msg="No data to add!"
        $ListBox14.Items.Add($msg)
        pconsole($msg)
    } else {
        $msg = "Choose group and OU!"
        $ListBox14.Items.Add($msg)
        pconsole($msg)
    }
}

function RichTextBox5MouseClick( $object ){
    if ($RichTextBox5.Text.Contains("Fill in"))
	{
	$RichTextBox5.Clear() }
}

function RichTextBox4MouseClick( $object ){
    if ($RichTextBox4.Text.Contains("Fill in"))
	{
	$RichTextBox4.Clear() }
}

function Button23Click( $object ){
    # Load
    $ListBox13.Items.Clear()
    $ListBox14.Items.Clear()
    $ListBox15.Items.Clear()
    $global:listOfAddedUsers = @()
    $listOfAddedUsersTemp = @()
    $global:listOfUsersFoundInAD=@()
	[array]$loadedIdis = $RichTextBox3.Text.Split("`n") | % { $_.trim() }
    $uniqueIdis = New-Object System.Collections.ArrayList
    $loadedIdis | Select-Object -Unique | foreach {$uniqueIdis.Add($_) }
    [array]$loadedNames = $RichTextBox4.Text.Split("`n") | % { $_.trim() }
    [array]$loadedSurnames = $RichTextBox5.Text.Split("`n") | % { $_.trim() }
    if(-not ($loadedIdis[0].Contains("Wprowad")))
    {
        if(($loadedIdis.Length -eq $loadedNames.Length -and $loadedIdis.Length -eq $loadedSurnames.Length -and $loadedNames.Length -eq $loadedSurnames.Length))
        {
	        for($i=0; $i -lt $loadedIdis.Length; $i++){
                if($loadedIdis[$i] -ne "")
                {
                    $item = New-Object PSObject
                    $item | Add-Member -type NoteProperty -Name 'Id' -Value $loadedIdis[$i]
                    $item | Add-Member -type NoteProperty -Name 'Name' -Value $loadedNames[$i]
                    $item | Add-Member -type NoteProperty -Name 'Surname' -Value $loadedSurnames[$i]
                    if($uniqueIdis -contains $loadedIdis[$i]){
                        $listOfAddedUsersTemp += $item
                        $uniqueIdis.Remove($loadedIdis[$i])
                    }
                }
            }
            foreach($item in $listOfAddedUsersTemp)
            {
                if($nameUser = getUserInfoAD($item.Id)){
                    $ListBox15.Items.Add("Id: "+$item.Id + " Name: "+$item.Name + " Surname: "+$item.Surname+" || In AD ("+$nameUser+")")
                    $global:listOfUsersFoundInAD+= $item.Id
                } else {
                    $global:listOfAddedUsers += $item
                    $ListBox13.Items.Add("Id: "+$item.Id + " Name: "+$item.Name + " Surname: "+$item.Surname)
                }
            }
        } else 
        {
            $ListBox13.Items.Add("Error of filled in data! Number of ID's: "+$loadedIdis.Length +", names: "+$loadedNames.Length+", surnames: "+$loadedSurnames.Length)
        }
    }
    else
    {
        $ListBox13.Items.Add("Fill in data.")
    }
    $global:listOfAddedUsers | foreach {pconsole($_.Id)}
}

function RichTextBox3MouseClick( $object ){
    if ($RichTextBox3.Text.Contains("Fill in"))
	{
	$RichTextBox3.Clear() }
}

function Button26Click( $object ){
    # search
    if($TextBox5.Text){
        $ComboBox2.Items.Clear()
        foreach($item in $listOfOUs){
            if($item.ToLower().Contains($TextBox5.Text.ToLower())){
                $ComboBox2.Items.Add($item)
            }
        }
        pconsole("List updated")
    }
}

function Button27Click( $object ){
    #copy
    copyToClipboard($ListBox14)
}

function Button28Click( $object ){
    $RichTextBox1.Clear()
    foreach($item in $global:listOfUsersFoundInAD){
        $RichTextBox1.Text += $item+"`n"
    }
    pconsole("ID's moved to password reset")
    $TabControl1.SelectedIndex=0;
    Button4Click($Button4)
}

function Button42Click( $object ){
    $RichTextBox6.Clear()
    foreach($item in $global:listOfUsersFoundInAD){
        $RichTextBox6.Text += $item+"`n"
    }
    pconsole("ID's moved to user modification")
    $TabControl1.SelectedIndex=4;
    Button29Click($Button29)
}

function ListBox18Click( $object ){
    if($global:listOfModifiedUsers[0]){
        $global:selectedModifiedUser = $global:listOfModifiedUsers[$ListBox18.SelectedIndex]
        try{
            $global:daneselectedModifiedUser = getUserInfoADMore($global:selectedModifiedUser)
            fillModifiedData($global:daneselectedModifiedUser)
            pconsole("Loaded data")
        } catch {
            pconsole($omfg)
        }
    }
}

function Button29Click( $object ){
    #load
    $ListBox18.Items.Clear()
	$ListBox19.Items.Clear()
    clearModifiedData
    $global:listOfModifiedUsers = @()
	[array]$loadedUserIds = $RichTextBox6.Text.Split("`n") | % { $_.trim() }
	foreach ($val in $loadedUserIds) { 
		pconsole($val)
        if($val -ne ""){
            if($nameUser = getUserInfoAD($val)){
			    $ListBox18.Items.Add("ID: " + $val + ", Name: " + $nameUser)
			    $global:listOfModifiedUsers += $val
            } else {
        	    $ListBox19.Items.Add($val)
            }
        }
		
	}
    if($ListBox18.Items.Count -gt 0 -or $ListBox19.Items.Count -gt 0){
        pconsole("Loaded data")
    }
}

function RichTextBox6MouseClick( $object ){
    if ($RichTextBox6.Text.Contains("Fill in")){
	    $RichTextBox6.Clear() 
    }
}

function Button30Click( $object ){
    # search
    if($TextBox6.Text){
        $ComboBox5.Items.Clear()
        foreach($item in $listOfOUs){
            if($item.ToLower().Contains($TextBox6.Text.ToLower())){
                $ComboBox5.Items.Add($item)
            }
        }
        pconsole("OU Unit list updated")
    }
}

function Button31Click( $object ){
    # search
    if($TextBox6.Text){
        $ComboBox4.Items.Clear()
        foreach($item in $listOfGroups){
            if($item.ToLower().Contains($TextBox6.Text.ToLower())){
                $ComboBox4.Items.Add($item)
            }
        }
        pconsole("Group list updated")
    }
}

function Button36Click( $object ){
    #update
    if($global:selectedModifiedUser -ne ""){
        try {
            fillModifiedData(getUserInfoADMore($global:selectedModifiedUser))
            pconsole("Data updated")
        } catch {
            pconsole($omfg)
        }
    }
}

function Button35Click( $object ){
    if($global:selectedModifiedUser -ne ""){
        if($global:daneselectedModifiedUser.mail -ne $TextBox11.Text -and $TextBox11.Text -ne ""){
            try{
                set-ADUser -Identity $global:selectedModifiedUser -EmailAddress $TextBox11.Text
                pconsole("E-mail changed to $($TextBox11.Text)")
            } catch {
                pconsole($omfg)
            }
        }
        if($global:daneselectedModifiedUser.Description -ne $TextBox10.Text){
            try {
                set-ADUser -Identity $global:selectedModifiedUser -Description $TextBox10.Text
                pconsole("Description changed to $($TextBox10.Text)")
            } catch {
                pconsole($omfg)
            }
        }
        if($global:daneselectedModifiedUser.DisplayName -ne $TextBox9.Text -and $global:daneselectedModifiedUser.Name -ne $TextBox9.Text){
            try {
                set-ADUser -Identity $global:selectedModifiedUser -DisplayName $TextBox9.Text
                Rename-ADObject -Identity $Label34.Text -NewName $TextBox9.Text
                pconsole("Display name changed to $($TextBox9.Text)")
            } catch {
                pconsole($omfg)
            }
        }
        if($global:daneselectedModifiedUser.Surname -ne $TextBox8.Text){
            try {
                set-ADUser -Identity $global:selectedModifiedUser -Surname $TextBox8.Text
                pconsole( "Surname changed to $($TextBox8.Text)")
            } catch {
                pconsole($omfg)
            }
        }
        if($global:daneselectedModifiedUser.GivenName -ne $TextBox7.Text){
            try {
                set-ADUser -Identity $global:selectedModifiedUser -GivenName $TextBox7.Text
                pconsole( "Name changed to $($TextBox7.Text)")
            } catch {
                pconsole($omfg)
            }
        }
        if($global:daneselectedModifiedUser.LockedOut -eq $true){
            try {
                Unlock-ADAccount -Identity $global:selectedModifiedUser
                pconsole( "Account unlocked $global:selectedModifiedUser")
            } catch {
                pconsole($omfg)
            }
        }
        if($global:daneselectedModifiedUser.PasswordExpired -ne $CheckBox8.Checked){
            try {
                set-ADUser -Identity $global:selectedModifiedUser -ChangePasswordAtLogon $CheckBox8.Checked
                if($CheckBox8.Checked){
                    pconsole("Forced password change on next logon")
                } else {
                    pconsole("Released password change on next logon")
                }
            } catch {
                pconsole($omfg)
            }
        }
        if($global:daneselectedModifiedUser.Enabled -ne (-Not ($CheckBox9.Checked))){
            if($CheckBox9.Checked){
                try {
                    Disable-ADAccount -Identity $global:selectedModifiedUser
                    pconsole("Account disabled")
                } catch {
                    pconsole($omfg)
                }

            } else {
                try {
                    Enable-ADAccount -Identity $global:selectedModifiedUser
                    pconsole("Account enabled")
                } catch {
                    pconsole($omfg)
                }
            }
        }
        if($global:daneselectedModifiedUser.SamAccountName -ne $TextBox12.Text){
            try {
                set-ADUser -Identity $global:selectedModifiedUser -UserPrincipalName ($TextBox12.Text+$atvariable) -SamAccountName $TextBox12.Text
                pconsole( "Login changed to $($TextBox12.Text)")
                Button29Click($Button29)
            } catch {
                pconsole($omfg)
            }
        }
    }
}

function Button34Click( $object ){
        if ($listOfGroups -contains $ComboBox4.SelectedItem){
            try {
                $group = Get-ADGroup -filter "Name -eq '$($ComboBox4.SelectedItem)'" | select -ExpandProperty DistinguishedName
                Remove-ADGroupMember -Identity $group -Members $global:selectedModifiedUser -Confirm:$false
                pconsole( "Deleted from "+$group)
                pconsole( "Change will take up to 30 seconds")
            } catch {
                pconsole( "Already deleted from selected group, or you have no permission to this group.")
            }
        } else {
            pconsole( "Choose correct group!")
        }
}

function Button33Click( $object ){
        if ($listOfGroups -contains $ComboBox4.SelectedItem){
            try {
                $group = Get-ADGroup -filter "Name -eq '$($ComboBox4.SelectedItem)'" | select -ExpandProperty DistinguishedName
                Add-ADGroupMember -Identity $group -Members $global:selectedModifiedUser
                pconsole( "Added to "+$group)
                pconsole( "Change will take up to 30 seconds")
            } catch {
                pconsole( "Already added to selected group, or you have no permission to this group.")
            }
        } else {
            pconsole( "Choose correct group!")
        }
}

function Button32Click( $object ){
    if ($listOfOUs -contains $ComboBox5.SelectedItem){
        try{
            move-ADObject -Identity $Label34.Text -TargetPath $ComboBox5.SelectedItem 
            pconsole( "Moved to "+$($ComboBox5.SelectedItem))
            pconsole( "Change will take up to 30 seconds")
            if($global:selectedModifiedUser -ne ""){
                fillModifiedData(getUserInfoADMore($global:selectedModifiedUser))
            }
        } catch {
            pconsole( "Error!")
        }
    }
}

function Button37Click( $object ){
    $RichTextBox1.Clear()
    $RichTextBox1.Text += $TextBox12.Text
    $TabControl1.SelectedIndex=0;
    Button4Click($Button4)
}

function Button38Click( $object ){
    $Label35.Visible = $true
}

function Label35Click( $object ){
    $Label35.Visible = $false
}

function TextBox8KeyUp( $object ){
    fillTextBox9
}

function TextBox7KeyUp( $object ){
    fillTextBox9
}

function Button39Click( $object ){
    if($ListBox20.SelectedItem -match $dcvariable){
        try {
            Remove-ADGroupMember -Identity $ListBox20.SelectedItem -Members $global:selectedModifiedUser -Confirm:$false
            pconsole( "Deleted from "+$($ListBox20.SelectedItem) )
            pconsole( "Change will take up to 30 seconds")
        } catch {
            pconsole( "Already deleted from selected group, or you have no permission to this group.")
        }
    }
}

function Button41Click( $object ){
    $RichTextBox7.Clear()
    pconsole("Console")
}

function Button40Click( $object ){
    try{
        [Windows.Forms.Clipboard]::SetText($RichTextBox7.Text)
    } catch {
        pconsole( $omfg)
    }
}

function TabPage6Click( $object ){

}

function TabPage7Click( $object ){

}

Main # This call must remain below all other event functions

#endregion
