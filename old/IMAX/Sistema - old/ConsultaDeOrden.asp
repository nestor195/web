<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Consulta de Orden</title>
</head>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaOrdenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL, ObConn
Select Case ObRs ("TipoCliente")
Case 0
%>
<body bgcolor="#FFFFFF">
<%
Case 1
%>
<body bgcolor="#FFFFAA">
<%
End Select
ObRs.Close
ObConn.Close
%>

<form method="GET" action="ConsultaDeOrden.asp" webbot-action="--WEBBOT-SELF--">
	<p><input type="text" name="Id" size="20"><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>
<form method="POST" action="ConsultaDeOrden.asp" webbot-action="--WEBBOT-SELF--">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaOrdenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL, ObConn
%>
	<p>Orden:<select size="1" name="Id">
	<option><%Response.Write ObRs ("Id")%></option>
	</select> <a href="fpdf/Orden.asp?Id=<%Response.Write ObRs ("Id")%>">Imprimir</a></p>
	<p><b>Datos del Cliente: </b>
	<a href="ModificarCliente.asp?IdCliente=<%Response.Write ObRs ("Cliente")%>">Modificar Cliente</a><br>
	<b>Nombre:</b> <%Response.Write ObRs ("Nombre")%>&nbsp;&nbsp;
	<b>&nbsp;&nbsp;&nbsp;&nbsp; Dirección:</b> <%Response.Write ObRs ("Direccion")%><br>
	<b>Teléfono:</b> <%Response.Write ObRs ("Telefono")%>
	<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Email:</b> <%Response.Write ObRs ("Email")%></p>
	<p><b>Datos del Equipo:</b><br>
	<b>Tipo:</b> <%Response.Write ObRs ("Tipo")%>&nbsp;&nbsp;
	<b>Marca: </b> <%Response.Write ObRs ("Marca")%>&nbsp;&nbsp;&nbsp;
	<b>Modelo:</b> <%Response.Write ObRs ("Modelo")%><br>
	<b>Serie:</b> <%Response.Write ObRs ("Serie")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="ListaOrdenSerie.asp?Serie=<%Response.Write ObRs ("Serie")%>">Otros 
    Ingresos del Equipo</a></p>
	<p><b>Fecha de Ingreso:</b> <%Response.Write ObRs ("FechaIngreso")%>&nbsp;&nbsp;&nbsp; 
	<b>Usuario de Ingreso:</b> <%Response.Write ObRs (4)%><br>
	<b>Estado:</b> <%Response.Write ObRs ("Estado")%>&nbsp;&nbsp;&nbsp;
	<b>Fecha de Estado:</b> <%Response.Write ObRs ("FechaEstado")%>
	<a target="_parent" href="ModificarEstadoOrden.asp?Id=<%Response.Write ObRs ("Id")%>">
	Modificar</a></p>
	<p><b>Técnico:</b> <%Response.Write ObRs (5)%><b><Br>
	Accesorios:</b> <%Response.Write ObRs ("Accesorios")%><br>
	<b>Observaciones de Ingreso:</b><br>
	<%Response.Write ObRs ("ObservacionIngreso")%></p>
	<p><b>Observaciones del Técnico:</b>
	<a target="_parent" href="ModificarObservacioTecnico.asp?Id=<%Response.Write ObRs ("Id")%>">
	Modificar</a><br>
	<%Response.Write ObRs ("ObservacionTecnico")%></p>
<%
ObRs.Close
ObConn.Close
%>
</form>

<form method="GET" action="IngresoOrdenItem.asp">
	<table border="1" width="100%" id="table1">
		<tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL, ObConn
%>

			<td width="51"><a href="ModificarEquipo.asp?IdEquipo=<%Response.Write ObRs ("Equipo")%>">Listado</a></td>
<%
ObRs.Close
ObConn.Close
%>
			<td width="51"><b>Código</b></td>
			<td width="446"><b>Descripción</b></td>
			<td width="55"><b>Cantidad</b></td>
			<td width="91"><b>Precio Unitario</b></td>
			<td><b>Total</b></td>
		</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = " & Request.QueryString("Id")
ObRs.Open SQL,ObConn
Total = 0
DO WHILE NOT ObRs.Eof
%>
		<tr>
			<td width="51">
			<a target="_parent" href="EliminarOrdenItem.asp?Id=<%Response.Write ObRs ("Id")%>">
			Eliminar</a></td>
			<td width="51"><%Response.Write ObRs ("Codigo")%>&nbsp;</td>
			<td width="446"><%Response.Write ObRs ("Descripcion")%>&nbsp;</td>
			<td width="55"><%Response.Write ObRs ("Cantidad")%>&nbsp;</td>
			<td width="91"><%Response.Write ObRs ("PrecioUnitario")%>&nbsp;</td>
			<td><%Response.Write ObRs ("Cantidad") * ObRs ("PrecioUnitario")%>&nbsp;</td>
		</tr>
<%
Total = Total + ObRs ("Cantidad") * ObRs ("PrecioUnitario")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
		<tr>
			<td width="643" colspan="5">
			<p align="right"><b>Total</b></td>
			<td><%Response.Write Total%>&nbsp;</td>
		</tr>
	</table>
	<p><select size="1" name="Id">
	<option value="<%Response.Write Request.QueryString("Id")%>">
	<%Response.Write Request.QueryString("Id")%></option>
	</select><input type="submit" value="Submit" name="B5">
	<input type="reset" value="Reset" name="B6"></p>
</form>

</body>

</html>