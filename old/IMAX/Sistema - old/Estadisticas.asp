<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Estadisticas</title>
</head>

<body>

<p><font face="Arial" size="1">Estadísticas</font></p>

<p><font face="Arial" size="1">Ingresos Mensuales</font></p>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#09/30/2008# and FechaEstado <#11/01/2008# And (Estado = 9 or Estado = 19 or Estado = 20)",ObConn
numer=0
DO WHILE NOT ObRs.Eof

SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext

LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Octubre : $<%Response.Write(FormatNumber(numer))%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#08/31/2008# and FechaEstado <#10/01/2008# And (Estado = 9 or Estado = 19 or Estado = 20)",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Septiembre : $<%Response.Write(FormatNumber(numer))%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#07/31/2008# and FechaEstado <#09/01/2008# And (Estado = 9 or Estado = 19 or Estado = 20)",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Agosto : $<%Response.Write(FormatNumber(numer))%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#06/30/2008# and FechaEstado <#08/01/2008# And (Estado = 9 or Estado = 19 or Estado = 20)",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Julio : $<%Response.Write(FormatNumber(numer))%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#05/31/2008# and FechaEstado <#07/01/2008# And (Estado = 9 or Estado = 19 or Estado = 20)",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Junio : $<%Response.Write(FormatNumber(numer))%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#04/30/2008# and FechaEstado <#06/01/2008# And (Estado = 9 or Estado = 19 or Estado = 20)",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Mayo : $<%Response.Write(FormatNumber(numer))%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#03/31/2008# and FechaEstado <#05/01/2008# And (Estado = 9 or Estado = 19 or Estado = 20)",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Abril : $<%Response.Write(FormatNumber(numer))%> </font> <br>


<p><font face="Arial" size="1">Ingresos en Cuenta Mensuales</font></p>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#09/30/2008# and FechaEstado <#11/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof

SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext

LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Octubre : $<%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#08/31/2008# and FechaEstado <#10/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Septiembre : $<%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#07/31/2008# and FechaEstado <#09/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Agosto : $<%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#06/30/2008# and FechaEstado <#08/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Julio : $<%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#05/31/2008# and FechaEstado <#07/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Junio : $<%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#04/30/2008# and FechaEstado <#06/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Mayo : $<%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#03/31/2008# and FechaEstado <#05/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Abril : $<%Response.Write(numer)%> </font> <br>

<p><font face="Arial" size="1">Ingresos de Equipos mensuales</font></p>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#09/30/2008# and FechaIngreso <#11/01/2008# and Estado <> 'Anulado' and Estado <> 'Vendido'",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Octubre : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#08/31/2008# and FechaIngreso <#10/01/2008# and Estado <> 'Anulado' and Estado <> 'Vendido'",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Septiembre : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#07/31/2008# and FechaIngreso <#09/01/2008# and Estado <> 'Anulado' and Estado <> 'Vendido'",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Agosto : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#06/30/2008# and FechaIngreso <#08/01/2008# and Estado <> 'Anulado' and Estado <> 'Vendido'",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Julio : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#05/31/2008# and FechaIngreso <#07/01/2008# and Estado <> 'Anulado' and Estado <> 'Vendido'",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Junio : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#04/30/2008# and FechaIngreso <#06/01/2008# and Estado <> 'Anulado' and Estado <> 'Vendido'",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Mayo : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#03/31/2008# and FechaIngreso <#05/01/2008# and Estado <> 'Anulado' and Estado <> 'Vendido'",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Abril : <%Response.Write(numer)%> </font> <br>
<p><font face="Arial" size="1">Ingresos de Equipos mensuales del gremio</font></p>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#09/30/2008# and FechaIngreso <#11/01/2008# And TipoCliente = 1",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Octubre : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#08/31/2008# and FechaIngreso <#10/01/2008# And TipoCliente = 1",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Septiembre : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#07/31/2008# and FechaIngreso <#09/01/2008# And TipoCliente = 1",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Agosto : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#06/30/2008# and FechaIngreso <#08/01/2008# And TipoCliente = 1",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Julio : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#05/31/2008# and FechaIngreso <#07/01/2008# And TipoCliente = 1",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Junio : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#04/30/2008# and FechaIngreso <#06/01/2008# And TipoCliente = 1",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Mayo : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From ConsultaOrdenes Where FechaIngreso >#03/31/2008# and FechaIngreso <#05/01/2008# And TipoCliente = 1",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Abril : <%Response.Write(numer)%> </font> <br>

<p><font face="Arial" size="1">Equipos Entregados y Cobrados</font></p>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#09/30/2008# and FechaEstado <#11/01/2008# And Estado = 9",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Octubre : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#08/31/2008# and FechaEstado <#10/01/2008# And Estado = 9",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Septiembre : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#07/31/2008# and FechaEstado <#09/01/2008# And Estado = 9",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Agosto : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#06/30/2008# and FechaEstado <#08/01/2008# And Estado = 9",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Julio : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#05/31/2008# and FechaEstado <#07/01/2008# And Estado = 9",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Junio : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#04/30/2008# and FechaEstado <#06/01/2008# And Estado = 9",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Mayo : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#03/31/2008# and FechaEstado <#05/01/2008# And Estado = 9",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Abril : <%Response.Write(numer)%> </font> <br>

<p><font face="Arial" size="1">Equipos En Cuenta</font></p>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#09/30/2008# and FechaEstado <#11/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Octubre : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#08/31/2008# and FechaEstado <#10/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Septiembre : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#07/31/2008# and FechaEstado <#09/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Agosto : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#06/30/2008# and FechaEstado <#08/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Julio : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#05/31/2008# and FechaEstado <#07/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Junio : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#04/30/2008# and FechaEstado <#06/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Mayo : <%Response.Write(numer)%> </font> <br>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >#03/31/2008# and FechaEstado <#05/01/2008# And Estado = 11",ObConn
numer=0
DO WHILE NOT ObRs.Eof
numer = numer + 1
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<font face="Arial" size="1">Abril : <%Response.Write(numer)%> </font> <br>


</body>

</html>