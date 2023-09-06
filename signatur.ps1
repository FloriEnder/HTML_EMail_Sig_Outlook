import-module ActiveDirectory
#Start Vars:
$HTMLSignaturPath_org = ""

#AD Abfarge nach Werten:
$AD_User = Get-ADUser -Identity $Env:UserName -Properties GivenName,sn,title,department,telephoneNumber,mail,persLinkedInLink,akademischerTitel

#Alle Werte erstzten:
$HTML = Get-Content -Path $HTMLSignaturPath_org -Encoding utf8
$HTML = $HTML.replace('@givenName',$AD_User.GivenName)
$HTML = $HTML.replace('@sn',$AD_User.sn)
$HTML = $HTML.replace('@telephoneNumber',$AD_User.telephoneNumber)
$HTML = $HTML.replace('@mail',$AD_User.mail)


###Einzelne Prüfung
##Prüfen ob ein pers. LinkedIN Adresse verfügbar ist
if ($AD_User.persLinkedInLink -eq $null) {
    $HTML = $HTML.replace('@LinkedIN','https://linkedin.com/firma')
}else {
    $HTML = $HTML.replace('@LinkedIN', $AD_User.persLinkedInLink)
}

#Datei im Sigantur Ordner hinterlegen
$HTML | Out-File -FilePath "$env:APPDATA\Microsoft\Signatures\signatur.htm" -Encoding utf8

$MSWord = New-Object -COMObject word.application 
$MSWord.EmailOptions.EmailSignature.NewMessageSignature="signatur"
$MSWord.EmailOptions.EmailSignature.ReplyMessageSignature="signatur"
$MSWord.Quit()
