<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="es-ar" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Listado de ordenes</title>
<link href="estilo.css" rel="stylesheet" type="text/css" />
</head>

<body>

<p><span class="Tilulo">Listado de ordenes</span></p>
<p>
<form method="post">
	<input name="valor" type="text" /><select name="Campo">
	<option value="Id">Orden</option>
	<option value="Equipo">Equipo</option>
	<option value="Empresa">Empresa</option>
	<option value="Estado">Estado</option>
	<option value="Fecha">Fecha</option>
	</select><input name="Button1" type="submit" value="Ir" /></form>
</p>
<table class="tablas" style="width: 100%">
	<tr>
		<td><span class="Titulo2">Orden</span></td>
		<td><span class="Titulo2">Equipo</span></td>
		<td><span class="Titulo2">Serie</span></td>
		<td><span class="Titulo2">Empresa</span></td>
		<td><span class="Titulo2">Estado</span></td>
		<td><span class="Titulo2">Fecha</span></td>
	</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Ordenes"
SQL = "SELECT Ordenes.Id, Clientes.Cliente, Equipos.Equipo, Equipos.Modelo, Estados.Estado, Ordenes.Serie, Ordenes.FechaIngreso"
SQL = SQL & " FROM ((Ordenes LEFT JOIN Estados ON Ordenes.Estado = Estados.Id) LEFT JOIN Clientes ON Ordenes.Cliente = Clientes.Id) LEFT JOIN Equipos ON Ordenes.Equipo = Equipos.Id"

if request.form ("Valor") = "" then
SQL = SQL & " Where " & Request.Form ("Campo") & " = '" & Request.Form ("Valor") & "'"
end if

SQL =SQL & " Order by Ordenes.Id Desc"
ObRs.Open SQL,ObConn
Do While ObRs.EOF = false
%>
	<tr>
		<td><a href="consultaorden.asp?Orden=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a></td>
		<td><%Response.Write ObRs ("Equipo")%></td>
		<td><%Response.Write ObRs ("Serie")%></td>
		<td><%Response.Write ObRs ("Cliente")%></td>
		<td><%Response.Write ObRs ("Estado")%></td>
		<td><%Response.Write ObRs ("FechaIngreso")%></td>
	</tr>
<%
ObRs.MoveNext
Loop
ObRs.Close
ObConn.Close
%>
</table>

</body>

</html>
