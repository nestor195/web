<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="es-ar" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Sin título 1</title>
<style type="text/css">
</style>
<link href="estilo.css" rel="stylesheet" type="text/css" />
</head>
<%
IF Request.Form ("enviado") = "" THEN

SET ObConn = Server.CreateObject ("ADODB.Connection")
ObConn.Open "Ordenes"

SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObRs.Open "Select Id from Clientes where Cliente = '" & Request.Querystring ("Cliente") & "'",ObConn

iF ObRs.EOF = false then
IdCliente = ObRs ("Id")
End If

ObRs.Close

SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObRs.Open "Select Id from Equipos where Modelo = '" & Request.Querystring ("Modelo") & "'",ObConn

iF ObRs.EOF = false then
IdModelo = ObRs ("Id")
End If
ObRs.Close

ObConn.Close

%>

<body>

<p><span class="Tilulo">Nueva Orden</span></p>
<form action="" method="post">
	<span class="Titulo2">Cliente:</span> 

	<select name="Cliente">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Ordenes"
ObRs.Open "Select Id, Cliente from Clientes order by Cliente asc",ObConn
Do While ObRs.EOF = false
If IdCliente = ObRs ("Id") then
%>
	<option selected="selected" value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Cliente")%></option>
<%
Else
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Cliente")%></option>
<%
End If
ObRs.MoveNext
Loop
ObRs.Close
ObConn.Close
%>
	</select>
	<a href="ingresocliente.asp?pagina=ingresoorden.asp">Nuevo</a><br />
	<span class="Titulo2">Equipo:</span> 
	<select name="Equipo">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Ordenes"
ObRs.Open "Select * from Acciones order by Accion asc",ObConn
Do While ObRs.EOF = false
If IdModelo = ObRs ("Id") then
%>
	<option selected="selected" value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Modelo")%></option>
<%
Else
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Modelo")%></option>
<%
End If
ObRs.MoveNext
Loop
ObRs.Close
ObConn.Close
%>
	</select>
	<a href="ingresoequipo.asp?pagina=ingresoorden.asp">Nuevo</a><br />
	<span class="Titulo2">Serie: 
	</span> 
	<input name="Serie" type="text" /><br />
	<span class="Titulo2">Adicionales:</span> 
	<input name="Adicionales" type="text" /><br />
	<span class="Titulo2">Observacion de Ingreso:</span> 
	<input name="Inconveniente" type="text" /><input class="Oculto" name="Enviado" type="text" value="true" /><br />
	<input name="Ingresar" type="submit" value="Ingresar" /></form>
<p>&nbsp;</p>

</body>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Ordenes"
ObRs.Open "Select * From Ordenes Order by Id desc",ObConn, 3, 3

numeroorden = ObRs ("Id") + 1

ObRs.AddNew
ObRs ("Cliente") = Request.Form ("cliente")
ObRs ("Equipo") = Request.Form ("Equipo")
ObRs ("FechaIngreso") = DATE
ObRs ("FechaEstado") = DATE
ObRs ("Serie") = Request.Form ("Serie")
ObRs ("Adicionales") = Request.Form ("Adicionales")
ObRs ("Observacion") = Request.Form ("Inconveniente")
ObRs ("Estado") = 1

ObRs.Update

ObRs.Close
ObConn.Close
Response.redirect "consultaorden.asp" & "?a=1" & "&Orden=" & numeroorden

END IF
%>

</html>
