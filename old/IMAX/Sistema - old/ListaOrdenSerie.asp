<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<form method="Get" action="ListaOrdenSerie.asp" webbot-action="--WEBBOT-SELF--">
	<p>Serie:<input type="text" name="Serie" size="20"></p>
	<p><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>
<table border="1" width="733" id="table1">
	<tr>
		<td width="63" bgcolor="#3399FF"><b>Nº Orden</b></td>
		<td width="217" bgcolor="#3399FF"><b>Cliente</b></td>
		<td bgcolor="#3399FF"><b>Equipo</b></td>
		<td width="96" bgcolor="#3399FF"><b>Estado</b></td>
		<td width="72" bgcolor="#3399FF"><b>Fecha de Estado</b></td>
		<td width="65" bgcolor="#3399FF">Fecha de Ingreso</td>
	</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
If Request.QueryString("Serie") = "" then
SQL = "Select * From ConsultaOrdenes Order By Id"
Else
SQL = "Select * From ConsultaOrdenes where Serie Like '%" & Request.QueryString("Serie") & "%' Order By Id"
End If
ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
Select Case ObRs ("TipoCliente")
Case 0
Color = "#FFFFFF"
Case 1
Color = "#FFFFAA"
End Select
%>
	<tr>
		<td width="63" bgcolor="<%Response.Write Color%>">&nbsp;<a href="ConsultaDeOrden.asp?Id=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a></td>
		<td width="217" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Nombre")%></td>
		<td bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Modelo")%></td>
		<td width="96" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Estado")%></td>
		<td width="72" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("FechaEstado")%></td>
		<td width="65" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("FechaIngreso")%></td>
	</tr>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>

</table>


</body>

</html>