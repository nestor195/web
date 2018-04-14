<%@ Language=VBScript %>
<%
Usuario = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "ACYP"
Sel = "SELECT Id FROM Usuarios Where Usuario = '" & Request.form ("Usuario") & "' And Contrasena = '" & Request.form ("Contrasena") & "'"
ObRs.Open Sel,ObConn
If ObRs.EOF = false Then
Usuario = ObRs ("Id")
End If
ObRs.Close
ObConn.Close

If Usuario <> 0 Then
Session("loginokay") = Usuario
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "ACYP"
Sel = "Select * From Log"
ObRs.Open Sel, ObConn, 3, 3
ObRs.Addnew
ObRs ("Fecha") = now()

ObRs ("Usuario") = Usuario
ObRs ("Evento") = "Login de Usuario"
ObRs.Update
ObRs.Close
ObConn.Close
End If


%>
<HTML>
<HEAD>
<meta content="es-ar" http-equiv="Content-Language" />
<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type" />

</HEAD>
<BODY>
<%if Session("loginokay") = "" then %>
<form action="login.asp" name=login method=post>
	Usuario:
	<input name="Usuario" type="text"><br>Contraseña:<input id="password" name="Contrasena" type="password"><br>
	&nbsp;<INPUT type="submit" value="Login" id=submit1 name=Enviar style="height: 26px; width: 47px">
</form>
<%else
Response.redirect "default.asp"
end if%>
</BODY>
</HTML>