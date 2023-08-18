Sub Main

'NACKarantina ile yakalanmış olan vlan bilgilerine ulaşıp bizim için çalıştıran Script
	crt.Screen.Send "show vlan" & chr(10)
	Dim result,int,port,ports,temp,newports
	result = crt.Screen.ReadString("#" & result)
	int= Split(result, "NACKarantina")
	port = split(int(1), vbCr)

	InputBox(int(1))
	ports = Split(port(0) , "active")
	
	InputBox(ports(1))
	temp = split(ports(1), ",")

	'for i 0 to UBound(temp)+1
	'Show(Trim(temp(i)))
	Dim i
	For i = 1 to UBound(temp)
		If temp(i) = chr(0) Then Exit For
		InputBox(temp(i))
		Show(Trim(temp(i)))
	Next
	End Sub

	Function Show(metin)

	crt.Screen.Send "show running-config interface " &  metin 

End Function