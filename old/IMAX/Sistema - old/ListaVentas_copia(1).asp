<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<form method="Get" action="ListaVentas_copia(1).asp" webbot-action="--WEBBOT-SELF--">
<span lang="es">Fecha</span>:
	<select size="1" name="Fecha">
	<option value="">HOY</option>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select Distinct FechaEstado From Ordenes Where Estado = 19 or Estado = 9 or Estado = 20",ObConn
DO WHILE NOT ObRs.Eof
If FormatDateTime(Request.QueryString("Fecha")) = FormatDateTime(ObRs ("FechaEstado")) THEN
%>
	<option selected value="<%Response.Write FormatDateTime(ObRs ("FechaEstado"))%>"><%Response.Write FormatDateTime(ObRs ("FechaEstado"))%></option>
<%
SQLFecha = " FechaEstado = '" & FormatDateTime(ObRs ("FechaEstado")) & "'"

ELSE
%>
	<option value="<%Response.Write FormatDateTime(ObRs ("FechaEstado"))%>"><%Response.Write FormatDateTime(ObRs ("FechaEstado"))%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select></p>
	<p><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>
<table border="1" width="787" id="table1">
	<tr>
		<td width="63"><span lang="es">Nº Orden</span></td>
		<td width="177"><span lang="es">Cliente</span></td>
		<td width="113"><span lang="es">Estado</span></td>
		<td width="475"><span lang="es">Ítems</span></td>
		<td width="189"><span lang="es">Total</span></td>
		<td width="230"><span lang="es">Fecha de Estado</span></td>
	</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
If Request.QueryString("Fecha") = "" then
SQL = "Select * From ConsultaOrdenes Where (Estado = 'Entregado y Cobrado' or Estado = 'Vendido' or Estado = 'Pagado') and FechaEstado = #10-1-2008# order by Id"
else
SQL = "Select * From ConsultaOrdenes Where (Estado = 'Entregado y Cobrado' or Estado = 'Vendido' or Estado = 'Pagado') and FechaEstado = #" & FormatDateTime(Request.QueryString("Fecha")) & "# order by Id"
end if
ObRs.Open  SQL,ObConn
Total = 0
DO WHILE NOT ObRs.Eof
%>
	<tr>
		<td width="63">&nbsp;<a href="ConsultaDeOrden.asp?Id=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a></td>
		<td width="177">&nbsp;<%Response.Write ObRs ("Nombre")%></td>
		<td width="113">&nbsp;<%Response.Write ObRs ("Estado")%></td>
		<td width="475">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
 <%
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = " & ObRs("Id")
ObRs2.Open SQL,ObConn
SubTotal = 0
DO WHILE NOT ObRs2.Eof
%>
         <tr>
            <td width="100%">&nbsp;<%Response.Write ObRs2 ("Descripcion")%></td>
          </tr>
<%
SubTotal = SubTotal + ObRs2 ("Cantidad") * ObRs2 ("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
%>
        </table>
        </td>
		<td width="189">&nbsp;$<%Response.Write FormatNumber(SubTotal,2)%></td>
		<td width="230">&nbsp;<%Response.Write FormatDateTime(ObRs ("FechaEstado"))%></td>
	</tr>
<%
Total = Total + SubTotal
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>

</table>
Total:&nbsp;$<%Response.Write FormatNumber(Total,2)%>

</body>

</html>