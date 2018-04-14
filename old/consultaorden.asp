<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="es-ar" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Orden</title>
<style type="text/css">
</style>
<link href="estilo.css" rel="stylesheet" type="text/css" />
</head>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Ordenes"
Sel = "SELECT Ordenes.Id, Clientes.Cliente, Equipos.Equipo, Estados.Estado, Ordenes.Serie, Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Adicionales, Ordenes.Inconveniente, Ordenes.Observacion"
Sel = Sel & " FROM ((Ordenes LEFT JOIN Clientes ON Ordenes.Cliente = Clientes.Id) LEFT JOIN Equipos ON Ordenes.Equipo = Equipos.Id) LEFT JOIN Estados ON Ordenes.Estado = Estados.Id Where Ordenes.Id = " & Request.Querystring ("Orden")
ObRs.Open Sel,ObConn
numeroorden = ObRs ("Id")
fechaingreso = ObRs ("FechaIngreso")
fechaestado = ObRs ("FechaEstado")
cliente = ObRs ("Cliente")
equipo = ObRs ("Equipo")
serie = ObRs ("Serie")
adicionales = ObRs ("Adicionales")
inconveniente = ObRs ("Inconveniente")
observacion = ObRs ("Observacion")
estado = ObRs ("Estado")

ObRs.Close
ObConn.Close
%>

<body>

<p class="auto-style1"><strong><span class="Titulo2">Orden:</span> <%Response.Write numeroorden%> 
<span class="Titulo2">Fecha ingreso: </span> <%Response.Write fechaingreso%></strong></p>
<p class="auto-style1"><strong class="Titulo2">Cliente: <%Response.Write cliente%></strong></p>
<p class="auto-style1"><strong class="Titulo2">Equipo: <%Response.Write equipo%></strong></p>
<p class="auto-style1"><strong class="Titulo2">Serie: <%Response.Write serie%></strong></p>
<p class="auto-style1"><strong class="Titulo2">Adicionales: <%Response.Write adicionales%></strong></p>
<p class="auto-style1"><strong class="Titulo2">Inconveniente declarado: <%Response.Write inconveniente%></strong></p>
<p class="auto-style1"><strong class="Titulo2">Observaciones Tecnicas: <%Response.Write observacion%></strong></p>
<p class="auto-style1"><strong><span class="Titulo2">Estado:</span> <%Response.Write estado%> 
<span class="Titulo2">Fecha:</span> <%Response.Write fechaestado%></strong></p>
<p>&nbsp;</p>

</body>

</html>
