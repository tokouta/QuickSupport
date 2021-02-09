

<###############################################################################################
#
# Quick Support GUI
#
#
# Dieses Script stellt ein Quick Support Tool dar.
#
#
# 19.10.2020   Shipinyuan Su : 1.4
#
################################################################################################>





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
#----------------------------------------------------------------------##----------------------------------------------------------------------#



#Die Grundform
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Quick Support'
$form.Size = New-Object System.Drawing.Size(580,550)
$form.StartPosition = 'CenterScreen'
$icon = New-Object system.drawing.icon (".\icon.ico")
$form.Icon = $icon
$form.AutoScroll = "true"



#----------------------------------------------------------------------#



#OK und Cancel Button -> Grundform
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(260,400)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'Ok'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::ok
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(360,400)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

#----------------------------------------------------------------------#

#TFBERN INFO (Rechte Seite)
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
$label6.Text = 'Tel.Nr.: .....'
$form.Controls.Add($label6)


$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(260,150)
$label3.Size = New-Object System.Drawing.Size(300,40)
$label3.Text = 'Support-Zeiten:
7:40 - 12:00 und 13:00 - 16:30'
$form.Controls.Add($label3)

$label7 = New-Object System.Windows.Forms.Label
$label7.Location = New-Object System.Drawing.Point(260,250)
$label7.Size = New-Object System.Drawing.Size(280,50)
$label7.Text = 'Für die meisten Anwedungen ist eine Verbindung zum "Domain" notwendig!'
$form.Controls.Add($label7)



#----------------------------------------------------------------------##----------------------------------------------------------------------#


#Hier sind die verschiedene Buttons und was sie tun sollten


########################################################################
#Beispiel#
<#

$GPButton = New-Object System.Windows.Forms.Button
$GPButton.Location = New-Object System.Drawing.Size(00,00)
$GPButton.Size = New-Object System.Drawing.Size(00,00)
$GPButton.Text = "Gruppenrichtlinien aktualisieren"
$GPButton.Add_Click(
{


Hier die Anweisung was der Button macht

}
)
$Form.Controls.Add($GPButton)


#Infotexte falls man fragen hat, bezüglich der Funktionen
$itext1 = New-Object System.Windows.Forms.label
$itext1.Location = New-Object System.Drawing.Point(260,300)
$itext1.Size = New-Object System.Drawing.Size(150,100)
$itext1.Text = 'Ein GPUPDATE /FORCE wird im CMD durchgeführt.
Dabei werden alle Richtlinien erneut angewandt'
$itext1.ForeColor = "blue"



#Info Buttons
$img = [System.Drawing.Image]::fromfile("Pfad")
$info1 = New-Object System.Windows.Forms.Button
$info1.Location = New-Object System.Drawing.Size(230,50)
$info1.Size = New-Object System.Drawing.Size(23,23)
$info1.Image = $img
$info1.Add_Click(
{

#Die neu erstelle Funktion
remove-text -additext $itext1


})
$Form.Controls.Add($info1)

#>
########################################################################
#
#Hier unten die Eigenschaften der Elemente eingeben
#
#----------------------------------------------------------------------#

#Funktion um den blauen Infotext zu entfernen und den richtigen hinzuzufügen 
#Funktion --> remove-text -additext $itext* <-- den Text welches hinzugefügt werden soll

function remove-text{

param( $additext )

#Hier $itext* erweitern
$allitexts = $itext1, $itext2, $itext3, $itext4, $itext5, $itext6, $itext7, $itext8, $itext9

foreach($allitext in $allitexts){

$form.Controls.Remove($allitext)

}

$form.controls.Add($additext)}

#----------------------------------------------------------------------#


#GPUPDATE Button
$GPButton = New-Object System.Windows.Forms.Button
$GPButton.Location = New-Object System.Drawing.Size(30,50)
$GPButton.Size = New-Object System.Drawing.Size(200,23)
$GPButton.Text = "Gruppenrichtlinien aktualisieren"
$GPButton.Add_Click(
{

Start-Process "\\.\gpupdate.bat" 

}
)
$Form.Controls.Add($GPButton)

#Infotexte falls man fragen hat, bezüglich der Funktionen
$itext1 = New-Object System.Windows.Forms.label
$itext1.Location = New-Object System.Drawing.Point(260,300)
$itext1.Size = New-Object System.Drawing.Size(150,100)
$itext1.Text = 'Ein GPUPDATE /FORCE wird im CMD durchgeführt.
Dabei werden alle Richtlinien erneut angewandt'
$itext1.ForeColor = "blue"



#Info Buttons
$img = [System.Drawing.Image]::fromfile(".\info1.png")
$info1 = New-Object System.Windows.Forms.Button
$info1.Location = New-Object System.Drawing.Size(230,50)
$info1.Size = New-Object System.Drawing.Size(23,23)
$info1.Image = $img
$info1.Add_Click(
{

remove-text -additext $itext1


})
$Form.Controls.Add($info1)

#----------------------------------------------------------------------#

#Laufwerk Button führt zum Script
$LFButton = New-Object System.Windows.Forms.Button
$LFButton.Location = New-Object System.Drawing.Size(30,90)
$LFButton.Size = New-Object System.Drawing.Size(200,23)
$LFButton.Text = "Laufwerke aktualisieren"
$LFButton.Add_Click(
{

                
              
    
    Start-Process "\\.\Laufwerke.bat"            



}
)
$Form.Controls.Add($LFButton)

$itext2 = New-Object System.Windows.Forms.label
$itext2.Location = New-Object System.Drawing.Point(260,300)
$itext2.Size = New-Object System.Drawing.Size(150,100)
$itext2.Text = 'Ein Script wird druchgeführt, welches die Berechtigungen kontrolliert und die passenden Laufwerke hinzufügt.'
$itext2.ForeColor = "blue"


$info2 = New-Object System.Windows.Forms.Button
$info2.Location = New-Object System.Drawing.Size(230,90)
$info2.Size = New-Object System.Drawing.Size(23,23)
$info2.Image = $img
$info2.Add_Click(
{

remove-text -additext $itext2

}
)
$Form.Controls.Add($info2)

#----------------------------------------------------------------------#

#Printer Button führt ein Script druch welcher alle Drucker Treiber löscht und wieder einrichtet
$PRButton = New-Object System.Windows.Forms.Button
$PRButton.Location = New-Object System.Drawing.Size(30,130)
$pRButton.Size = New-Object System.Drawing.Size(200,23)
$prButton.Text = "Drucker Treiber erneuern"
$prButton.Add_Click(
{

Start-Process "\\.\Delete_Printer.bat"



}
)
$Form.Controls.Add($PRButton)

$itext3 = New-Object System.Windows.Forms.label
$itext3.Location = New-Object System.Drawing.Point(260,300)
$itext3.Size = New-Object System.Drawing.Size(150,100)
$itext3.Text = 'Ein Skript wird durchgeführt, bei welchem alle Druckertreiber erneuert werden.'
$itext3.ForeColor = "blue"


$info3 = New-Object System.Windows.Forms.Button
$info3.Location = New-Object System.Drawing.Size(230,130)
$info3.Size = New-Object System.Drawing.Size(23,23)
$info3.Image = $img
$info3.Add_Click(
{

remove-text -additext $itext3

}
)
$Form.Controls.Add($info3)

#----------------------------------------------------------------------#


<#Script für das erstellen eines neues Outlook Profiles

Das script hat noch ein zwischen Fenster und ist deshalb so lang#>

$OLButton = New-Object System.Windows.Forms.Button 
$OLButton.Location = New-Object System.Drawing.Size(30,170)
$olButton.Size = New-Object System.Drawing.Size(200,23)
$olButton.Text = "Outlook Profil erneuern"
$OLButton.add_click(
 
{
#Warning Fenster Outlook Profil erstellen
$form1 = New-Object System.Windows.Forms.Form
$form1.Text = 'Achtung!'
$form1.Size = New-Object System.Drawing.Size(300,150)
$form1.StartPosition = 'CenterScreen'
$form1.FormBorderStyle = "fixeddialog"
$icon1 = New-Object system.drawing.icon ("\\.\warning.ico")
$form1.Icon = $icon1

$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(10,20)
$label5.Size = New-Object System.Drawing.Size(300,40)
$label5.Text = 'Sind Sie sicher, dass ein neues Outlook Profil erstellt werden sollte? '
$form1.Controls.Add($label5)

#OK und Cancel Button -> Warning Fenster Outlook
$OKButton1 = New-Object System.Windows.Forms.Button
$OKButton1.Location = New-Object System.Drawing.Point(30,70)
$OKButton1.Size = New-Object System.Drawing.Size(75,23)
$OKButton1.Text = 'ja'
$OKButton1.DialogResult = [System.Windows.Forms.DialogResult]::ok
$form1.AcceptButton = $OKButton1
$form1.Controls.Add($OKButton1)

$CancelButton1 = New-Object System.Windows.Forms.Button
$CancelButton1.Location = New-Object System.Drawing.Point(120,70)
$CancelButton1.Size = New-Object System.Drawing.Size(75,23)
$CancelButton1.Text = 'Nein'
$form1.CancelButton = $CancelButton1
$form1.Controls.Add($CancelButton1)

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
[void]$form1.ShowDialog()

})
$Form.Controls.Add($olButton)


$itext4 = New-Object System.Windows.Forms.label
$itext4.Location = New-Object System.Drawing.Point(260,300)
$itext4.Size = New-Object System.Drawing.Size(150,100)
$itext4.Text = 'Es wird ein neuer Outlook Profil erstellt. Das alte Profil wird archiviert.'
$itext4.ForeColor = "blue"


$info4 = New-Object System.Windows.Forms.Button
$info4.Location = New-Object System.Drawing.Size(230,170)
$info4.Size = New-Object System.Drawing.Size(23,23)
$info4.Image = $img
$info4.Add_Click(
{

remove-text -additext $itext4

}
)
$Form.Controls.Add($info4)

#----------------------------------------------------------------------#


#Passwort zurücksetzen Button
$Passwort = New-Object System.Windows.Forms.Button
$Passwort.Location = New-Object System.Drawing.Size(30,210)
$Passwort.Size = New-Object System.Drawing.Size(200,23)
$Passwort.Text = "Passwort zurücksetzen"
$Passwort.Add_Click(
{  

Start-Process #verweist auf PW Seite der Firma

}
)

$Form.Controls.Add($Passwort)


#Infotexte falls man fragen hat, bezüglich der Funktionen
$itext5 = New-Object System.Windows.Forms.label
$itext5.Location = New-Object System.Drawing.Point(260,300)
$itext5.Size = New-Object System.Drawing.Size(150,100)
$itext5.Text = 'Sie werden weitergeleitet auf einer Webseite, wo man das Kennwort zurücksetzten kann.'
$itext5.ForeColor = "blue"



#Info Buttons
$img = [System.Drawing.Image]::fromfile("\\.\info1.png")
$info5 = New-Object System.Windows.Forms.Button
$info5.Location = New-Object System.Drawing.Size(230,210)
$info5.Size = New-Object System.Drawing.Size(23,23)
$info5.Image = $img
$info5.Add_Click(
{

remove-text -additext $itext5


})
$Form.Controls.Add($info5)
#----------------------------------------------------------------------#

#PDF Drucker einrichten
$PRTREIBER = New-Object System.Windows.Forms.Button
$PRTREIBER.Location = New-Object System.Drawing.Size(30,250)
$PRTREIBER.Size = New-Object System.Drawing.Size(200,23)
$PRTREIBER.Text = "PDF Drucker einrichten"
$PRTREIBER.Add_Click(
{

Start-Process "\\.\Add_PDF_Printer.bat"

}
)
$form.controls.Add($PRTREIBER)


#Infotexte falls man fragen hat, bezüglich der Funktionen
$itext6 = New-Object System.Windows.Forms.label
$itext6.Location = New-Object System.Drawing.Point(260,300)
$itext6.Size = New-Object System.Drawing.Size(150,100)
$itext6.Text = 'Der PDF Drucker wird hinzugefügt.'
$itext6.ForeColor = "blue"



#Info Buttons
$img = [System.Drawing.Image]::fromfile("\\.\info1.png")
$info6 = New-Object System.Windows.Forms.Button
$info6.Location = New-Object System.Drawing.Size(230,250)
$info6.Size = New-Object System.Drawing.Size(23,23)
$info6.Image = $img
$info6.Add_Click(
{

remove-text -additext $itext6


})
$Form.Controls.Add($info6)
#----------------------------------------------------------------------#

#Programme und Features öffnen
$appwiz = New-Object System.Windows.Forms.Button
$appwiz.Location = New-Object System.Drawing.Size(30,290)
$appwiz.Size = New-Object System.Drawing.Size(200,23)
$appwiz.Text = "Programme und Features"
$appwiz.Add_Click(
{

appwiz.cpl

}
)


$form.controls.Add($appwiz)


#Infotexte falls man fragen hat, bezüglich der Funktionen
$itext7 = New-Object System.Windows.Forms.label
$itext7.Location = New-Object System.Drawing.Point(260,300)
$itext7.Size = New-Object System.Drawing.Size(150,100)
$itext7.Text = 'Das Fenster für Programme zu deinstallieren oder zu ändern öffnet sich.'
$itext7.ForeColor = "blue"



#Info Buttons
$img = [System.Drawing.Image]::fromfile("\\.\info1.png")
$info7 = New-Object System.Windows.Forms.Button
$info7.Location = New-Object System.Drawing.Size(230,290)
$info7.Size = New-Object System.Drawing.Size(23,23)
$info7.Image = $img
$info7.Add_Click(
{

remove-text -additext $itext7


})
$Form.Controls.Add($info7)

#----------------------------------------------------------------------#


#%Temp% Ordnerinhalt löschen
$temp = New-Object System.Windows.Forms.Button
$temp.Location = New-Object System.Drawing.Size(30,330)
$temp.Size = New-Object System.Drawing.Size(200,23)
$temp.Text = "User Temp Ordnerinhalt löschen"
$temp.Add_Click(
{

Start-Process "\\.\tempdel.bat - Shortcut.lnk"

}
)
$form.controls.Add($temp)

#Infotexte falls man fragen hat, bezüglich der Funktionen
$itext8 = New-Object System.Windows.Forms.label
$itext8.Location = New-Object System.Drawing.Point(260,300)
$itext8.Size = New-Object System.Drawing.Size(150,100)
$itext8.Text = 'Der %TEMP% Ordner wird bestmöglichst gelöscht.'
$itext8.ForeColor = "blue"



#Info Buttons
$img = [System.Drawing.Image]::fromfile("\\.\info1.png")
$info8 = New-Object System.Windows.Forms.Button
$info8.Location = New-Object System.Drawing.Size(230,330)
$info8.Size = New-Object System.Drawing.Size(23,23)
$info8.Image = $img
$info8.Add_Click(
{

remove-text -additext $itext8


})
$Form.Controls.Add($info8)


#----------------------------------------------------------------------#





#----------------------------------------------------------------------#




#----------------------------------------------------------------------##----------------------------------------------------------------------#


#Button für den AD QuickSearch
$ADBUTTON = New-Object System.Windows.Forms.Button
$ADBUTTON.Location = New-Object System.Drawing.Size(260,190)
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
$textBox1.ScrollBars = "Vertical"
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
$gruppenmitglieder = ((Get-ADGroup $gruppenname  -Properties Members).Members -replace '\\,', "" -replace '^CN=([^,]+),OU=.+$','$1') -join "`r`n"


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
$textBox3.ScrollBars = "Vertical"
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


$Usermemberships = ((Get-ADPrincipalGroupMembership -Identity $User).name) -join "`r`n"


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





#Diesen Befehl bitte nicht löschen
[void]$form.ShowDialog()


