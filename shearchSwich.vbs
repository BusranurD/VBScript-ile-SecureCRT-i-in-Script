# $language = "VBScript"
# $interface = "1.0"
' istediğimiz Sunucu ve Port bilgilerine ulaşıp içerisinde bulunan ip değerinin bir üstüne dileğimiz
' sessiona giriş yaparak açma işlemi

Sub Main

'nkdysession="session sekmesinden istenilen sekme - Sunucu"

Dim ip, ipAd
Dim temp
temp = InputBox("Degeri girin:")


Dim regex
Set regex = New RegExp
regex.Pattern = "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"


crt.Screen.Send "show router " & temp & " interface" &chr(10)


Dim result
result = crt.Screen.ReadString("#")
Set matches = regex.Execute(result)



regex.Global = True
Dim write
write = ""
Set matches = regex.Execute(result)
Dim match,i,liste(99)
i = 1
liste(0) = ""

Dim cihaz,cihazAd
cihaz = Split(result,"vprn-me|")
For Each match In matches
	cihazAd = Split(cihaz(i)," ")
	temp = Split(match,".")
	If temp(0) = "255" Then
	else
	liste(i) = match
	write = write & CStr(i) & "---" & cihazAd(0) & "------>" &match & vbCr
	i = i+1
	end If
Next

index = InputBox(write & vbcr & vbcr& "Secimi girin:")
	If index > 0 And index < UBound(liste) Then
	Set YeniTab = crt.Session.ConnectInTab("/S """&nkdysession&"""")
	YeniTab.Screen.Synchronous = True
	'yeni tab açıldığında istediğimiz ifadeyi görene kadar Scripti bekletiriz
	YeniTab.Screen.WaitForString "Type to search or select one:"
	Dim ipParts, newIP
	   ipParts = Split(liste(index), ".")
	   newIP = CInt(ipParts(3)) + 1
	   newIP = ipParts(0) & "." & ipParts(1) & "." & ipParts(2) & "." & CStr(newIP)
	YeniTab.Screen.Send newIp & vbCr
	else
	MsgBox "secim hatali !!!"
	End If

End Sub
