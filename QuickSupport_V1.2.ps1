
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
$form.Size = New-Object System.Drawing.Size(460,400)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = "fixeddialog"
$icon = New-Object system.drawing.icon ("\\.\.ico")
$form.Icon = $icon
$form.MaximizeBox = $false
$form.AutoScale = "Dpi"

#----------------------------------------------------------------------#


#OK und Cancel Button -> Grundform
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(30,250)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'Ok'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::ok
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(130,250)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

#OK und Cancel Button -> Nach Mehr Anzeigen
$OKButton2 = New-Object System.Windows.Forms.Button
$OKButton2.Location = New-Object System.Drawing.Point(30,460)
$OKButton2.Size = New-Object System.Drawing.Size(75,23)
$OKButton2.Text = 'Ok'
$OKButton2.DialogResult = [System.Windows.Forms.DialogResult]::ok
$form.AcceptButton = $OKButton2


$CancelButton2 = New-Object System.Windows.Forms.Button
$CancelButton2.Location = New-Object System.Drawing.Point(130,460)
$CancelButton2.Size = New-Object System.Drawing.Size(75,23)
$CancelButton2.Text = 'Cancel'
$CancelButton2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton2

#----------------------------------------------------------------------#

# INFO (Rechte Seite)
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

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(260,70)
$label2.Size = New-Object System.Drawing.Size(300,20)
$label2.Text = "Benutzername: $env:USERNAME"
$form.Controls.Add($label2)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(260,90)
$label3.Size = New-Object System.Drawing.Size(300,35)
$label3.Text = "Helpdesk:
https://helpdesk.ch"
$form.Controls.Add($label3)

$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(260,125)
$label4.Size = New-Object System.Drawing.Size(300,20)
$label4.Text = 'Tel.Nr.: ...'
$form.Controls.Add($label4)


$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(260,150)
$label5.Size = New-Object System.Drawing.Size(300,40)
$label5.Text = 'Support-Zeiten:
7:40 - 12:00 und 13:00 - 16:30'
$form.Controls.Add($label5)

$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(10,290)
$label6.Size = New-Object System.Drawing.Size(280,50)
$label6.Text = 'Für die meisten Anwedungen ist eine Verbindung zum "Domain" notwendig!'
$form.Controls.Add($label6)

if(Test-Connection -Count 1 -ComputerName vmdc10 -Quiet){

$img1 = [System.Drawing.Image]::Fromfile("\\.\Grünerhäckchen.png")
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.location = New-Object System.Drawing.Point(300,285)
$pictureBox.Width = $img1.Size.Width
$pictureBox.Height = $img1.Size.Height
$pictureBox.Image = $img1
$form.controls.add($pictureBox)

} else {

$img2 = [System.Drawing.Image]::Fromfile("\\.\Roteskreuz.png")
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.location = New-Object System.Drawing.Point(300,285)
$pictureBox.Width = $img2.Size.Width
$pictureBox.Height = $img2.Size.Height
$pictureBox.Image = $img2
$form.controls.add($pictureBox)

}

#----------------------------------------------------------------------#

#Hier sind die verschiedene Buttons und was sie tun sollten

#GPUPDATE Button
$GPButton = New-Object System.Windows.Forms.Button
$GPButton.Location = New-Object System.Drawing.Size(30,50)
$GPButton.Size = New-Object System.Drawing.Size(200,23)
$GPButton.Text = "Gruppenrichtlinien aktualisieren"
$GPButton.Add_DoubleClick(
{

Start-Process "\\.\gpupdate.bat"

}
)
$Form.Controls.Add($GPButton)

#Laufwerk Button führt zum Script
$LFButton = New-Object System.Windows.Forms.Button
$LFButton.Location = New-Object System.Drawing.Size(30,90)
$LFButton.Size = New-Object System.Drawing.Size(200,23)
$LFButton.Text = "Laufwerke aktualisieren"
$LFButton.Add_Click(
{

                
              
    cd \\vmdienst1\DCSWRepository\RI
    Start-Process "\\.\Laufwerke.bat"            



}
)
$Form.Controls.Add($LFButton)

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

$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(10,20)
$label6.Size = New-Object System.Drawing.Size(300,40)
$label6.Text = 'Sind Sie sicher, dass ein neues Outlook Profil erstellt werden sollte? '
$form1.Controls.Add($label6)

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

#----------------------------------------------------------------------#


#Infotexte falls man fragen hat, bezüglich der Funktionen
$itext1 = New-Object System.Windows.Forms.label
$itext1.Location = New-Object System.Drawing.Point(260,200)
$itext1.Size = New-Object System.Drawing.Size(150,100)
$itext1.Text = 'Ein GPUPDATE /FORCE wird im CMD durchgeführt.
Dabei werden alle Richtlinien erneut angewandt'
$itext1.ForeColor = "blue"


$itext2 = New-Object System.Windows.Forms.label
$itext2.Location = New-Object System.Drawing.Point(260,200)
$itext2.Size = New-Object System.Drawing.Size(150,100)
$itext2.Text = 'Ein Script wird druchgeführt, welches die Berechtigungen kontrolliert und die passenden Laufwerke hinzufügt.'
$itext2.ForeColor = "blue"


$itext3 = New-Object System.Windows.Forms.label
$itext3.Location = New-Object System.Drawing.Point(260,200)
$itext3.Size = New-Object System.Drawing.Size(150,100)
$itext3.Text = 'Die Richtlinien werden erneut angewandt. Sie können per Skript versuchen den Drucker manuel hinzuzufügen.'
$itext3.ForeColor = "blue"

$itext4 = New-Object System.Windows.Forms.label
$itext4.Location = New-Object System.Drawing.Point(260,200)
$itext4.Size = New-Object System.Drawing.Size(150,100)
$itext4.Text = 'Es wird ein neuer Outlook Profil erstellt. Das alte Profil wird archiviert.'
$itext4.ForeColor = "blue"


#----------------------------------------------------------------------#
#Info Buttons
$img = [System.Drawing.Image]::fromfile("\\.\info1.png")
$info1 = New-Object System.Windows.Forms.Button
$info1.Location = New-Object System.Drawing.Size(230,50)
$info1.Size = New-Object System.Drawing.Size(23,23)
$info1.Image = $img
$info1.Add_Click(
{

$form.Controls.Remove($itext4)
$form.Controls.Remove($itext3)
$form.Controls.Remove($itext2)
$form.Controls.Add($itext1)


}
)
$Form.Controls.Add($info1)


$info2 = New-Object System.Windows.Forms.Button
$info2.Location = New-Object System.Drawing.Size(230,90)
$info2.Size = New-Object System.Drawing.Size(23,23)
$info2.Image = $img
$info2.Add_Click(
{

$form.Controls.Remove($itext4)
$form.Controls.Remove($itext3)
$form.Controls.Remove($itext1)
$form.Controls.Add($itext2)


}
)
$Form.Controls.Add($info2)



$info3 = New-Object System.Windows.Forms.Button
$info3.Location = New-Object System.Drawing.Size(230,130)
$info3.Size = New-Object System.Drawing.Size(23,23)
$info3.Image = $img
$info3.Add_Click(
{

$form.Controls.Remove($itext4)
$form.Controls.Remove($itext2)
$form.Controls.Remove($itext1)
$form.Controls.Add($itext3)

}
)
$Form.Controls.Add($info3)


$info4 = New-Object System.Windows.Forms.Button
$info4.Location = New-Object System.Drawing.Size(230,170)
$info4.Size = New-Object System.Drawing.Size(23,23)
$info4.Image = $img
$info4.Add_Click(
{


$form.Controls.Remove($itext3)
$form.Controls.Remove($itext2)
$form.Controls.Remove($itext1)
$form.Controls.Add($itext4)


}
)
$Form.Controls.Add($info4)






#----------------------------------------------------------------------#
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

$itext5 = New-Object System.Windows.Forms.label
$itext5.Location = New-Object System.Drawing.Point(220,50)
$itext5.Size = New-Object System.Drawing.Size(200,50)
$itext5.Text = 'Auf der Webseite können Sie, Ihr AD Kennwort zurücksetzen.'
$itext5.ForeColor = "blue"


#>
########################################################################
#
#Hier unten die Eigenschaften der Elemente eingeben
#
#----------------------------------------------------------------------#
#Passwort zurücksetzen Button
$Passwort = New-Object System.Windows.Forms.Button
$Passwort.Location = New-Object System.Drawing.Size(30,210)
$Passwort.Size = New-Object System.Drawing.Size(200,23)
$Passwort.Text = "Passwort zurücksetzen"
$Passwort.Add_Click(
{  

Start-Process #weiterleitung 

}
)

$itext5 = New-Object System.Windows.Forms.label
$itext5.Location = New-Object System.Drawing.Point(240,210)
$itext5.Size = New-Object System.Drawing.Size(200,50)
$itext5.Text = 'Auf der Webseite können Sie, Ihr AD Kennwort zurücksetzen.'
$itext5.ForeColor = "blue"


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

$itext6 = New-Object System.Windows.Forms.label
$itext6.Location = New-Object System.Drawing.Point(240,250)
$itext6.Size = New-Object System.Drawing.Size(200,50)
$itext6.Text = 'PDF Drucker wird eingerichtet.'
$itext6.ForeColor = "blue"



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

$itext7 = New-Object System.Windows.Forms.label
$itext7.Location = New-Object System.Drawing.Point(240,290)
$itext7.Size = New-Object System.Drawing.Size(200,50)
$itext7.Text = 'Öffnet Programme und Features'
$itext7.ForeColor = "blue"



#----------------------------------------------------------------------#


#%Temp% Ordnerinhalt löschen
$temp = New-Object System.Windows.Forms.Button
$temp.Location = New-Object System.Drawing.Size(30,330)
$temp.Size = New-Object System.Drawing.Size(200,23)
$temp.Text = "%TEMP% löschen"
$temp.Add_Click(
{

Start-Process "\\.\tempdel.bat - Shortcut.lnk"

}
)

$itext8 = New-Object System.Windows.Forms.label
$itext8.Location = New-Object System.Drawing.Point(240,330)
$itext8.Size = New-Object System.Drawing.Size(200,50)
$itext8.Text = 'Löscht den Ordnerinhalt des %Temp% Ordners.'
$itext8.ForeColor = "blue"



#----------------------------------------------------------------------#

#nach Windows Updates suchen
$Updates = New-Object System.Windows.Forms.Button
$Updates.Location = New-Object System.Drawing.Size(30,370)
$Updates.Size = New-Object System.Drawing.Size(200,23)
$Updates.Text = "Nach Updates suchen"
$Updates.Add_Click(
{

(New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow()


}
)

$itext9 = New-Object System.Windows.Forms.label
$itext9.Location = New-Object System.Drawing.Point(240,370)
$itext9.Size = New-Object System.Drawing.Size(200,50)
$itext9.Text = 'Sucht nach Windows Updates.'
$itext9.ForeColor = "blue"



#----------------------------------------------------------------------#


#----------------------------------------------------------------------#
<#
Der untere Abschnitt ist vorallem zum erweitern.
Es geht um den "Mehr Anzeigen" Funktion.


  #>




#Mehr anzeigen Button
$Anzeigen = New-Object System.Windows.Forms.Button 
$Anzeigen.Location = New-Object System.Drawing.Size(80,200)
$Anzeigen.Size = New-Object System.Drawing.Size(100,20)
$Anzeigen.Text = "Mehr anzeigen"
$form.Refresh()
$Anzeigen.add_click(

{

$form.Controls.Remove($OKButton)
$form.Controls.Remove($CancelButton)
$form.Controls.Remove($label6)
$form.Controls.Remove($pictureBox)
$form.controls.Remove($Anzeigen)
$form.Controls.Remove($itext1)
$form.Controls.Remove($itext2)
$form.Controls.Remove($itext3)
$form.Controls.Remove($itext4)
$form.AutoSize = $true
$form.Anchor = "top, left"



$form.Controls.Add($CancelButton2)
$form.Controls.Add($OKButton2)
$Form.Controls.Add($weniger)

#----------------------------------------------------------------------#
#Hier unten die Elemente hinzufügen
#
$Form.Controls.Add($Passwort)
#
#
$form.controls.Add($PRTREIBER)
#
#
$form.controls.Add($appwiz)
#
#
$form.controls.Add($temp)
#
#
$form.controls.Add($Updates)
#










$form.Refresh();



})


$Form.Controls.Add($Anzeigen)


#Weniger anzeigen Button
$weniger = New-Object System.Windows.Forms.Button 
$weniger.Location = New-Object System.Drawing.Size(70,410)
$weniger.Size = New-Object System.Drawing.Size(110,20)
$weniger.Text = "Weniger Anzeigen"
$weniger.add_click(

{



#----------------------------------------------------------------------#
#Hier unten die Elemente hinzufügen
#
$Form.Controls.Remove($Passwort)
#
#
$form.controls.Remove($PRTREIBER)
#
#
$form.controls.Remove($appwiz)
#
#
$form.controls.Remove($temp)
#
#
$form.controls.Remove($Updates)
#



$form.Controls.Add($OKButton)
$form.Controls.Add($CancelButton)
$form.Controls.Add($label1)
$form.Controls.Add($label2)
$form.Controls.Add($label3)
$form.Controls.Add($label4)
$form.Controls.add($label5)
$form.Controls.Add($label6)
$form.Controls.Add($pictureBox)
$form.controls.Add($Anzeigen)
$form.Controls.Remove($itext1)
$form.Controls.Remove($itext2)
$form.Controls.Remove($itext3)
$form.Controls.Remove($itext4)
$form.Controls.Remove($OKButton2)
$form.Controls.Remove($CancelButton2)
$form.Controls.Remove($weniger)
$form.AutoSize = $false


})








#Diesen Befehl bitte nicht löschen
#############################
### [void]$form.ShowDialog()#
[void]$form.ShowDialog()#####
### [void]$form.ShowDialog()#
#############################


