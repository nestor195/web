<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<%
if Session("IMAX") = False then
Response.Redirect ("administrativo.asp")
End if
%>

<form method="POST" action="ListadoOrdenes.asp" webbot-action="--WEBBOT-SELF--"><p>
  <span lang="es-ar">Cliente: </span><select size="1" name="Nombre">
 <%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Clientes Order By Nombre"
ObRs.Open SQL,ObConn
SubTotal = 0
DO WHILE NOT ObRs.Eof
If Request.Form ("Nombre") = ObRs ("Nombre") then
%>
  <option selected value="<%Response.Write ObRs ("Nombre")%>"><%Response.Write ObRs ("Nombre")%></option>
<%
else
%>
  <option value="<%Response.Write ObRs ("Nombre")%>"><%Response.Write ObRs ("Nombre")%></option>
<%
End if
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
  </select></p>
  <p><input type="submit" value="Enviar" name="B1"><input type="reset" value="Restablecer" name="B2"></p>
</form>

<table border="1" width="811" id="table1">
	<tr>
		<td width="63"><span lang="es">Nº Orden</span></td>
		<td width="260"><span lang="es">Cliente</span></td>
		<td width="107"><span lang="es">Estado</span></td>
		<td width="545"><span lang="es">Ítems</span></td>
		<td width="140"><span lang="es">Total</span></td>
		<td width="156"><span lang="es">Fecha de Estado</span></td>
	</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaOrdenes Where (Estado = 'Entregado y Cobrado' or Estado = 'Vendido' or Estado = 'Pagado') and Nombre = '" & Request.Form ("Nombre") & "' order by Id"
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