<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Consulta de Orden</title>
</head>

<form method="GET" action="IngresoVentaItem.asp">
	<table border="1" width="100%" id="table1">
		<tr>

			<td width="51">&nbsp;</td>
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
SQL = "Select * From ConsultaOrdenItem Where Orden = 1"
ObRs.Open SQL,ObConn
Total = 0
DO WHILE NOT ObRs.Eof
%>
		<tr>
			<td width="51">
			<a target="_parent" href="EliminarVentaItem.asp?Id=<%Response.Write ObRs ("Id")%>">
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
	<p><input type="submit" value="Agregar Item" name="B5">
	<input type="reset" value="Reset" name="B6"></p>
</form>

</body>

<p><span lang="es"><a href="IngresoOrdenVenta.asp">Ingreso Venta</a></span></p>


</html>