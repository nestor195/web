<html>

<head>
<meta http-equiv="Content-Language" content="es-ar">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Usuario</title>
</head>

<body>
<%
if Request.QueryString("Logout") = 1 then
Session ("usuario") = ""
end if
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios where Nick = '" & Request.Form("User") &"' and Password = '" & Request.Form("Password") & "'"
ObRs.Open SQL,ObConn
If not ObRs.Eof then
Session ("Session") = ObRs("Id")
Usuario = ObRs("Nick")
end if
ObRs.Close
ObConn.Close

IF Session("Usuario") = "" THEN
%>

<form method="POST" action="Login.asp" webbot-action="--WEBBOT-SELF--">
  <p>
  Usuario: <select size="1" name="D1"></select></p>
  <p>Contraseña: <input type="text" name="T1" size="20"></p>
  <p><input type="submit" value="Enviar" name="B1"></p>
</form>
<%
ELSE
response.redirect("default.asp")
END IF
%>

</body>

</html>