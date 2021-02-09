Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing



#----------------------------------------------------------------------#

#Versteckt die Console falls man das PS öffnet im Exe nicht nötig

function Show-Console
{
    param ([Switch]$Show,[Switch]$Hide)
    if (-not ("Console.Window" -as [type])) { 

        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        '
    }

    if ($Show)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()

        # Hide = 0,
        # ShowNormal = 1,
        # ShowMinimized = 2,
        # ShowMaximized = 3,
        # Maximize = 3,
        # ShowNormalNoActivate = 4,
        # Show = 5,
        # Minimize = 6,
        # ShowMinNoActivate = 7,
        # ShowNoActivate = 8,
        # Restore = 9,
        # ShowDefault = 10,
        # ForceMinimized = 11

        $null = [Console.Window]::ShowWindow($consolePtr, 5)
    }

    if ($Hide)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        #0 hide
        $null = [Console.Window]::ShowWindow($consolePtr, 0)
    } 
}

show-console -hide
#----------------------------------------------------------------------#


#Die Grundform
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Quick Support'
$form.Size = New-Object System.Drawing.Size(450,350)
$form.StartPosition = 'CenterScreen'
$icon = New-Object system.drawing.icon ("C:\.\.ico")
$form.Icon = $icon


#----------------------------------------------------------------------#


#OK und Cancel Button -> Grundform
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(40,220)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'Ok'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::ok
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(140,220)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

#----------------------------------------------------------------------#

#INFO (Rechte Seite)
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(270,20)
$label.Text = 'Bitte Wählen Sie, den Support den Sie benötigen:'
$form.Controls.Add($label)

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(260,50)
$label1.Size = New-Object System.Drawing.Size(300,20)
$label1.Text = "Hostname: $env:computername"
$form.Controls.Add($label1)

$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(260,70)
$label5.Size = New-Object System.Drawing.Size(300,20)
$label5.Text = "Benutzername: $env:USERNAME"
$form.Controls.Add($label5)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(260,90)
$label2.Size = New-Object System.Drawing.Size(300,35)
$label2.Text = "Helpdesk:
https://helpdesk.ch"
$form.Controls.Add($label2)

$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(260,125)
$label6.Size = New-Object System.Drawing.Size(300,20)
$label6.Text = 'Tel.Nr.: ...'
$form.Controls.Add($label6)


$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(260,150)
$label3.Size = New-Object System.Drawing.Size(300,40)
$label3.Text = 'Support-Zeiten:
7:40 - 12:00 und 13:00 - 16:30'
$form.Controls.Add($label3)

$label7 = New-Object System.Windows.Forms.Label
$label7.Location = New-Object System.Drawing.Point(30,260)
$label7.Size = New-Object System.Drawing.Size(280,50)
$label7.Text = 'Für die meisten Anwedungen ist eine Verbindung zum "Domain" notwendig!'
$form.Controls.Add($label7)




#----------------------------------------------------------------------#

#Hier sind die verschiedene Buttons und was sie tun sollten

#GPUPDATE Button
$SkriptButton = New-Object System.Windows.Forms.Button
$SkriptButton.Location = New-Object System.Drawing.Size(30,50)
$SkriptButton.Size = New-Object System.Drawing.Size(200,35)
$SkriptButton.Text = "Skripte"
$SkriptButton.Add_Click(
{


$form1 = New-Object System.Windows.Forms.Form
$form1.Text = 'Skripte'
$form1.Size = New-Object System.Drawing.Size(350,400)
$form1.StartPosition = 'CenterScreen'
$form1.FormBorderStyle = "fixeddialog"
$form1.Icon = $icon
$form1.AutoScroll = 'True'



$OLButton = New-Object System.Windows.Forms.Button 
$OLButton.Location = New-Object System.Drawing.Size(40,30)
$olButton.Size = New-Object System.Drawing.Size(250,30)
$olButton.Text = "Outlook Profil erneuern"
$OLButton.add_click(
 
{
#Warning Fenster Outlook Profil erstellen
$form2 = New-Object System.Windows.Forms.Form
$form2.Text = 'Achtung!'
$form2.Size = New-Object System.Drawing.Size(300,150)
$form2.StartPosition = 'CenterScreen'
$form2.FormBorderStyle = "fixeddialog"
$icon1 = New-Object system.drawing.icon ("C:\.\warning.ico")
$form2.Icon = $icon1

$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(10,20)
$label5.Size = New-Object System.Drawing.Size(300,40)
$label5.Text = 'Sind Sie sicher, dass ein neues Outlook Profil erstellt werden sollte? '
$form2.Controls.Add($label5)

#OK und Cancel Button -> Warning Fenster Outlook
$OKButton1 = New-Object System.Windows.Forms.Button
$OKButton1.Location = New-Object System.Drawing.Point(30,70)
$OKButton1.Size = New-Object System.Drawing.Size(75,23)
$OKButton1.Text = 'ja'
$OKButton1.DialogResult = [System.Windows.Forms.DialogResult]::ok
$form2.AcceptButton = $OKButton1
$form2.Controls.Add($OKButton1)

$CancelButton1 = New-Object System.Windows.Forms.Button
$CancelButton1.Location = New-Object System.Drawing.Point(120,70)
$CancelButton1.Size = New-Object System.Drawing.Size(75,23)
$CancelButton1.Text = 'Nein'
$form2.CancelButton = $CancelButton1
$form2.Controls.Add($CancelButton1)

$OKButton1.add_click(

#Hier ist der Script für ein neues Outlook Profil
{
$Path1 = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
$Path2 = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles"
$DefaultPath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook"
 
If ( -not ( Test-Path -Path $Path1 ) )
	{
	$Path1 = $null
	}
If ( -not ( Test-Path -Path $Path2 ) )
	{
	$Path2 = $null
	}
 
if ( -not $Path1 -and -not $Path2 )
	{
	write-Error "Reg-Path does not exist"
	exit
	}
 
$Val = (Get-ItemProperty -Path $DefaultPath -Name DefaultProfile).DefaultProfile
 
#Check if Outlook is runing end end the proccess
$Outlook = Get-Process -Name "OUTLOOK"
if ( $Outlook )
	{
	$OLPath= $Outlook.Path
	$Outlook | Stop-Process -Force
	}
 
if ( $Val -like "Outlook*" )
	{
	#Profile "Outlook" already exist. Use Timestamp to be unique
	$Profile = "Outlook-" + $(Get-Date -UFormat "%y%m%d%H%M")
	}
else
	{
	$Profile = "Outlook"
	}
 
# Create empty Profile on new Position
if ( $Path1 )
	{
	New-Item -Path $( Join-Path -Path $Path1 -ChildPath $Profile ) -Force | Out-Null
	}
 
# Create empty Profile on new legacy Position
if ( $Path2 )
	{
	New-Item -Path $( Join-Path -Path $Path2 -ChildPath $Profile ) -Force | Out-Null
	}
 
# Set new Profile as default
New-ItemProperty -Path $DefaultPath -Name DefaultProfile -Value $Profile -PropertyType String -Force | Out-Null	
 
# Start Outlook if it was running...
if ( $OLPath )
	{
	Start-Process -FilePath $OLPath -WindowStyle Minimized
	}
}


)
[void]$form2.ShowDialog()

})

$Form1.Controls.Add($olButton)

#----------------------------------------------------------------------#


#Laufwerk Button führt zum Script
$LFButton = New-Object System.Windows.Forms.Button
$LFButton.Location = New-Object System.Drawing.Size(40,80)
$LFButton.Size = New-Object System.Drawing.Size(250,30)
$LFButton.Text = "Laufwerke aktualisieren"
$LFButton.Add_Click(
{         
              
    
    Start-Process "\\.\Laufwerke.bat"            

}
)
$Form1.Controls.Add($LFButton)

#----------------------------------------------------------------------#


#Printer Button führt ein Script druch welcher alle Drucker Treiber löscht und wieder einrichtet
$PRButton = New-Object System.Windows.Forms.Button
$PRButton.Location = New-Object System.Drawing.Size(40,130)
$pRButton.Size = New-Object System.Drawing.Size(250,30)
$prButton.Text = "Drucker Treiber erneuern"
$prButton.Add_Click(
{

Start-Process "\\.\Delete_Printer.bat"



}
)
$Form1.Controls.Add($PRButton)



#----------------------------------------------------------------------#

#PDF Drucker einrichten
$PRTREIBER = New-Object System.Windows.Forms.Button
$PRTREIBER.Location = New-Object System.Drawing.Size(40,180)
$PRTREIBER.Size = New-Object System.Drawing.Size(250,30)
$PRTREIBER.Text = "PDF Drucker einrichten"
$PRTREIBER.Add_Click(
{

Start-Process "\\.\Add_PDF_Printer.bat"

}
)

$form1.controls.Add($PRTREIBER)
#----------------------------------------------------------------------#

#GPUPDATE Button
$GPButton = New-Object System.Windows.Forms.Button
$GPButton.Location = New-Object System.Drawing.Size(40,230)
$GPButton.Size = New-Object System.Drawing.Size(250,30)
$GPButton.Text = "Gruppenrichtlinien aktualisieren"
$GPButton.Add_DoubleClick(
{

Start-Process "\\.\gpupdate.bat"

}
)
$Form1.Controls.Add($GPButton)

#----------------------------------------------------------------------#

#%Temp% Ordnerinhalt löschen
$temp = New-Object System.Windows.Forms.Button
$temp.Location = New-Object System.Drawing.Size(40,280)
$temp.Size = New-Object System.Drawing.Size(250,30)
$temp.Text = "%TEMP% löschen"
$temp.Add_Click(
{

Start-Process "\\.\tempdel.bat - Shortcut.lnk"

}
)
$form1.controls.Add($temp)

#----------------------------------------------------------------------#



#----------------------------------------------------------------------#




#----------------------------------------------------------------------#

[void]$form1.ShowDialog()

}
)
$Form.Controls.Add($SkriptButton)



#--------------------------------------------------------------------------------------------------------------------------------------------#


#Laufwerk Button führt zum Script
$Weiterbutton = New-Object System.Windows.Forms.Button
$Weiterbutton.Location = New-Object System.Drawing.Size(30,100)
$Weiterbutton.Size = New-Object System.Drawing.Size(200,35)
$Weiterbutton.Text = "Weiterleitung"
$Weiterbutton.Add_Click(
{
              
$form3 = New-Object System.Windows.Forms.Form
$form3.Text = 'Weiterleitung'
$form3.Size = New-Object System.Drawing.Size(350,400)
$form3.StartPosition = 'CenterScreen'
$form3.FormBorderStyle = "fixeddialog"
$form3.Icon = $icon
$form3.AutoScroll = 'True' 


#----------------------------------------------------------------------#



#Passwort zurücksetzen Button
$Passwort = New-Object System.Windows.Forms.Button
$Passwort.Location = New-Object System.Drawing.Size(40,30)
$Passwort.Size = New-Object System.Drawing.Size(250,30)
$Passwort.Text = "Passwort zurücksetzen"
$Passwort.Add_Click(
{  

Start-Process #Weiterleitung

}
)

$Form3.Controls.Add($Passwort)

#----------------------------------------------------------------------#

#Programme und Features öffnen
$appwiz = New-Object System.Windows.Forms.Button
$appwiz.Location = New-Object System.Drawing.Size(40,80)
$appwiz.Size = New-Object System.Drawing.Size(250,30)
$appwiz.Text = "Programme und Features"
$appwiz.Add_Click(
{

appwiz.cpl

}
)
$form3.controls.Add($appwiz)
#----------------------------------------------------------------------#

#nach Windows Updates suchen
$Updates = New-Object System.Windows.Forms.Button
$Updates.Location = New-Object System.Drawing.Size(40,130)
$Updates.Size = New-Object System.Drawing.Size(250,30)
$Updates.Text = "Nach Updates suchen"
$Updates.Add_Click(
{

(New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow()


}
)

$form3.controls.Add($Updates)

#----------------------------------------------------------------------#




#----------------------------------------------------------------------#

           
[void]$form3.ShowDialog()             



}
)
$Form.Controls.Add($Weiterbutton)




#----------------------------------------------------------------------##----------------------------------------------------------------------#





#Printer Button führt ein Script druch welcher alle Drucker Treiber löscht und wieder einrichtet
$ADBUTTON = New-Object System.Windows.Forms.Button
$ADBUTTON.Location = New-Object System.Drawing.Size(30,150)
$ADBUTTON.Size = New-Object System.Drawing.Size(200,35)
$ADBUTTON.Text = "AD QuickSearch"
$ADBUTTON.Add_Click(
{
Import-Module ActiveDirectory

$form4 = New-Object System.Windows.Forms.Form
$form4.Text = 'AD QuickSearch'
$form4.Size = New-Object System.Drawing.Size(350,600)
$form4.StartPosition = 'CenterScreen'
$form4.FormBorderStyle = "fixeddialog"
$form4.Icon = $icon
$form4.AutoScroll = 'True' 

#----------------------------------------------------------------------#

#Gruppe suchen und Mitglieder herausfinden.

$label8 = New-Object System.Windows.Forms.Label
$label8.Location = New-Object System.Drawing.Point(10,20)
$label8.Size = New-Object System.Drawing.Size(140,30)
$label8.Text = 'Gruppenname eingeben:'
$form4.Controls.Add($label8)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,50)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form4.Controls.Add($textBox)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(10,80)
$textBox1.Size = New-Object System.Drawing.Size(260,150)
$textBox1.Multiline = $true
$textBox1.Text = "Gruppenmitglieder:


"
$form4.Controls.Add($textBox1)





$Suchbutton = New-Object System.Windows.Forms.Button
$Suchbutton.Location = New-Object System.Drawing.Size(270,45)
$Suchbutton.Size = New-Object System.Drawing.Size(60,30)
$Suchbutton.Text = "Suchen"
$Suchbutton.Add_Click(
{

  
$gruppenname = $textBox.Text
$gruppenname
$gruppenmitglieder = ((Get-ADGroup $gruppenname  -Properties Members).Members -replace '\\,', "" -replace '^CN=([^,]+),OU=.+$','$1')


$textBox1.Text = "Gruppenmitglieder:

$gruppenmitglieder
"
 




})


$Form4.Controls.Add($Suchbutton);


#----------------------------------------------------------------------#

#User suchen und Gruppen herausfinden.

$label9 = New-Object System.Windows.Forms.Label
$label9.Location = New-Object System.Drawing.Point(10,300)
$label9.Size = New-Object System.Drawing.Size(140,30)
$label9.Text = 'Username eingeben:'
$form4.Controls.Add($label9)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,330)
$textBox2.Size = New-Object System.Drawing.Size(260,20)
$form4.Controls.Add($textBox2)

$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(10,360)
$textBox3.Size = New-Object System.Drawing.Size(260,150)
$textBox3.Multiline = $true
$textBox3.Text = "Mitglied von:

"
$form4.Controls.Add($textBox3)




$Suchbutton1 = New-Object System.Windows.Forms.Button
$Suchbutton1.Location = New-Object System.Drawing.Size(270,325)
$Suchbutton1.Size = New-Object System.Drawing.Size(60,30)
$Suchbutton1.Text = "Suchen"
$Suchbutton1.Add_Click(
{

  
$User = $textBox2.Text


$Usermemberships = ((Get-ADPrincipalGroupMembership -Identity $User).name)


$textBox3.Text = "Mitglied von:

$Usermemberships
"

    })


$Form4.Controls.Add($Suchbutton1);




#----------------------------------------------------------------------#

[void]$form4.ShowDialog() 
}
)
$Form.Controls.Add($ADBUTTON)


#----------------------------------------------------------------------##----------------------------------------------------------------------#











#Diesen Befehl bitte nicht löschen
[void]$form.ShowDialog()

