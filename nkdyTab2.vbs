# $language = "VBScript"
# $interface = "1.0"
' aradığımız vlan değerinin altında bulunan ipnin bir üstünü istedğimiz session üzerinden açan Script

Sub Main

	strVersionPart = Split(crt.Version, " ")(0)
	vVersionElements = Split(strVersionPart, ".")

	
	On Error Resume Next
	crt.Screen.Synchronous = True
	Dim result
	Dim vLines
	Dim lc
	'nkdysession="session sekmesinden istenilen sekme - Sunucu"
	lc = 0
	
	Dim ip, ipAd
	Dim vlanId
	vlanId = InputBox("Vlan değerini girin:", "Show Vlan")
	crt.Screen.Send "show running-config interface vlan " & vlanId & chr(10)

	result = crt.Screen.ReadString("#" & result)
	ip= Split(result, "address")
	ipAd= Split(Trim(ip(1))," ")


	Dim ipParts, newIP
	   ipParts = Split(ipAd(0), ".")
	   newIP = CInt(ipParts(3)) + 1
	   newIP = ipParts(0) & "." & ipParts(1) & "." & ipParts(2) & "." & CStr(newIP)
		

	Set YeniTab = crt.Session.ConnectInTab("/S """nkdysession"""")
	YeniTab.Screen.Synchronous = True
	'yeni tab açıldığında istediğimiz ifadeyi görene kadar Scripti bekletiriz
	YeniTab.Screen.WaitForString "Type to search or select one:"
	


	YeniTab.Screen.Send newIP & vbCr

End Sub