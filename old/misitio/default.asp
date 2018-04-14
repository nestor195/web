<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//ES" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="es-ar" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<link href="estilos.css" rel="stylesheet" type="text/css" />
<title>Inicio Gestion</title>
</head>

<body>

<p><span class="Titulo">GESTION</span></p>
<form method="post" action="principal.asp">
<p>Seleccion de Empresa: <select name="Empresa">
	<option selected="selected" value="0">Elija una Opción</option>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Gestion"
SQL = "Select Id, NombreCorto From Empresas"
ObRs.Open SQL,ObConn

DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("NombreCorto")%></option>
<%
ObRs.MoveNext
LOOP

ObRs.Close
ObConn.Close
%>
</select></p>
	<p><input name="Boton" type="submit" value="Ingresar a Empresa" /></p>
</form>
</body>

</html>
