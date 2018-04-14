<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "ACYP"
SQL = "SELECT * FROM Usuarios WHERE Usuario = '" & Request.Form ("Usuario") & "' AND Contrasena = '"
SQL = SQL & Request.Form ("Contrasena") & "'"
ObRs.Open SQL,ObConn
If ObRs.EOF = False Then
Session("username") = ObRs ("Id")
End If
ObRs.Close
ObConn.Close
 
If Session("username") <> "" Then
Response.Redirect ("default.asp")
End If%>
<html>
<head>

<meta content="es-ar" http-equiv="Content-Language">
<title>Inicio</title>

<meta content="text/html; charset=windows-1252" http-equiv="Content-Type">

<link href="estilo.css" rel="stylesheet" type="text/css">

</head>
<%

%>
<body>
<form method="post" action=""><p>Usuario<input name="Usuario" type="text"></p>
<p>Contraseña<input name="Contrasena" type="password"></p>
	<p><input name="Entrar" type="submit" value="Entrar"></p>
</form>
</body>
</html>