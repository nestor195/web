<html>

<head>
<meta http-equiv="Content-Language" content="es-ar">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Ingreso comentario</title>
</head>
<%
IF Request.Form = "" THEN
%>

<body text="#FFFFFF" bgcolor="#000000">

<p><b><font face="Arial">Ingresar comentario</font></b></p>
<form method="POST" action="ingresocomentario.asp" webbot-action="--WEBBOT-SELF--">
  <p>
  <textarea rows="2" name="S1" cols="28"></textarea></p>
  <p><input type="submit" value="Enviar" name="B1"></p>
</form>

</body>
<%
ELSE
END IF
%>

</html>