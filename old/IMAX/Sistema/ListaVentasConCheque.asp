<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<form method="Get" action="ListaVentasConCheque.asp" webbot-action="--WEBBOT-SELF--">
<span lang="es">Fecha</span>:
	<select size="1" name="Fecha">
	<option value="">HOY</option>
<% 
Function FormatMediumDate(DateValue)
Dim strYYYY
Dim strMM
Dim strDD

strYYYY = CStr(DatePart("yyyy", DateValue))

strMM = CStr(DatePart("m", DateValue))
If Len(strMM) = 1 Then strMM = "0" & strMM

strDD = CStr(DatePart("d", DateValue))
If Len(strDD) = 1 Then strDD = "0" & strDD

FormatMediumDate = strMM & "/" & strDD & "/" & strYYYY

End Function 
%> 
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select Distinct FechaEstado From Ordenes Where Estado = 28 Order by FechaEstado desc"
ObRs.Open SQL, ObConn
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
<table border="1" width="811" id="table1">
	<tr>
		<td width="63"><span lang="es">Nº Orden</span></td>
		<td width="260"><span lang="es">Cliente</span></td>
		<td width="107"><span lang="es">Estado</span></td>
		<td width="545"><span lang="es">Ítems</span></td>
		<td width="140"><span lang="es">Total</span></td>
		<td width="156"><span lang="es">Fecha de Estado</span></td>
		<td width="156">Fecha de Cobro</td>
	</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
If Request.QueryString("Fecha") = "" then
SQL = "SELECT Ordenes.Id, Clientes.Nombre, Estados.Estado, Ordenes.FechaEstado,"
SQL = SQL & " Cheque.Fecha FROM (Estados INNER JOIN (Clientes "
SQL = SQL & "INNER JOIN Ordenes ON Clientes.Id = Ordenes.Cliente) "
SQL = SQL & "ON Estados.Id = Ordenes.Estado) INNER JOIN Cheque ON "
SQL = SQL & "Ordenes.Id = Cheque.Orden"

SQL = SQL & " Where (Estados.Estado = 'Con Cheque') and FechaEstado = #" & FormatMediumDate(DATE) & "# order by Ordenes.Id"
else
SQL = "SELECT Ordenes.Id, Clientes.Nombre, Estados.Estado, Ordenes.FechaEstado,"
SQL = SQL & " Cheque.Fecha FROM (Estados INNER JOIN (Clientes "
SQL = SQL & "INNER JOIN Ordenes ON Clientes.Id = Ordenes.Cliente) "
SQL = SQL & "ON Estados.Id = Ordenes.Estado) INNER JOIN Cheque ON "
SQL = SQL & "Ordenes.Id = Cheque.Orden"

SQL = SQL & " Where (Estados.Estado = 'Con Cheque') and FechaEstado = #" & FormatMediumDate(Request.QueryString("Fecha")) & "# order by Ordenes.Id"
end if
ObRs.Open  SQL,ObConn
Total = 0
DO WHILE NOT ObRs.Eof
%>
	<tr>
		<td width="63">&nbsp;<a href="ConsultaDeOrden.asp?Id=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a></td>
		<td width="260">&nbsp;<%Response.Write ObRs ("Nombre")%></td>
		<td width="107">&nbsp;<%Response.Write ObRs ("Estado")%></td>
		<td width="545">
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
            <td width="7%">&nbsp;<%Response.Write ObRs2 ("Cantidad")%></td>
            <td width="93%">&nbsp;<%Response.Write ObRs2 ("Descripcion")%></td>
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
		<td width="140">&nbsp;$<%Response.Write FormatNumber(SubTotal,2)%></td>
		<td width="156">&nbsp;<%Response.Write FormatDateTime(ObRs ("FechaEstado"))%></td>
		<td width="156">&nbsp;<%Response.Write FormatDateTime(ObRs ("Fecha"))%></td>
	</tr>
<%
Total = Total + SubTotal
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>

</table>
<span style="background-color: #FFFFFF">Total:&nbsp;$<%Response.Write FormatNumber(Total,2)%>
</span>

</body>

</html>