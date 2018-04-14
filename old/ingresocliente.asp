<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="es-ar" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Nuevo Cliente&nbsp; Nombre</title>
<style type="text/css">
</style>
<link href="estilo.css" rel="stylesheet" type="text/css" />
</head>
<%
IF Request.Form ("enviado") = "" THEN
%>

<body>

<form action="ingresocliente.asp" method="post" webbot-action="--WEBBOT-SELF--">
	<span class="Tilulo">Nuevo Cliente</span><br class="auto-style1" />
	<br class="auto-style1" />
	<span class="auto-style1">Nombre: </span>
	<input class="auto-style1" name="Cliente" type="text" /><br class="auto-style1" />
	<span class="auto-style1">Habilitado: </span>
	<input class="auto-style1" name="Habilitado" type="checkbox" checked="checked" value="true" /><input class="Oculto" name="Enviado" type="text" value="true" /><input class="Oculto" name="pagina" type="text" value='<%response.write Request.Querystring ("pagina")%>' /><input class="Oculto" name="Equipo" type="text" value='<%response.write Request.Querystring ("Equipo")%>' /><strong><br class="auto-style1" />
	<br />
	</strong><input name="Ingresar" type="submit" value="Ingresar" /></form>


</body>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Ordenes"
ObRs.Open "Clientes",ObConn, 3, 3

ObRs.AddNew
ObRs ("Cliente") = Request.Form ("cliente")
ObRs ("Habilitado") = Request.Form ("Habilitado")
ObRs.Update

ObRs.Close
ObConn.Close
Response.redirect Request.Form("pagina") & "?a=1" & "&Modelo=" & Request.Form ("modelo") & "&Cliente=" & Request.Form ("cliente")

END IF
%>

</html>
