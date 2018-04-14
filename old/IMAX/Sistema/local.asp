<HTML>
<BODY>
<%
Dim ArrayIPLocalStart(2)
Dim ArrayIPLocalEnd(2)
Dim ArrayIPClient
Dim IPClient 
Dim blnLocal
Ipclient = Request.ServerVariables("REMOTE_ADDR")
IpLocal = Request.ServerVariables("Local_ADDR")
If IpLocal <> Ipclient then
	blnLocal = False
	' These IP address are constant IP of a 
	' LAN
	ArrayIPLocalStart(0) = "010.000.000.000" '"10.0.0.0"
	ArrayIPLocalEnd(0) = "010.255.255.255" '"10.255.255.255"
	ArrayIPLocalStart(1) = "172.016.000.000" '"172.16.0.0"
	ArrayIPLocalEnd(1) = "172.031.000.000" '"172.31.0.0"
	ArrayIPLocalStart(2) = "192.168.000.000" '"192.168.0.0"
	ArrayIPLocalEnd(2) = "192.168.255.000" '"192.168.255.0"
	ArrayIPClient = Split(Ipclient,".")
	' format the remote IP like constant LAN
	' IP
	For i = LBound(ArrayIPClient) To UBound(ArrayIPClient)
	    ArrayIPClient(i) = String(3-Len(ArrayIPClient(i) ),"0") & ArrayIPClient(i)
	Next
	IPClient = Join(ArrayIPClient,"")
	If Trim(ipclient) <> "" Then
	    For i = LBound(ArrayIPLocalStart) To UBound(ArrayIPLocalStart)
	        ArrayIPLocalStart(i) = Replace(ArrayIPLocalStart(i),".","")
	        ArrayIPLocalEnd(i) = Replace(ArrayIPLocalEnd(i),".","")
	        If IPClient >= ArrayIPLocalStart(i) And IPClient =< ArrayIPLocalEnd(i) Then
	            blnLocal = True
	            Exit For
	        End If
	    Next
	End If
End If
	If blnLocal Or IpLocal = Ipclient Then
	    Response.Write " And your computer has a local IP"
	Else
	    Response.Write " And your computer has a internet IP"
	End If
	%>
</BODY>
</HTML>