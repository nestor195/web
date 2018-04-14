<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<form method="Get" action="ListaCliente.asp" webbot-action="--WEBBOT-SELF--">
Cliente:
	<select size="1" name="Cliente">
	<option value="">Todos</option>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Clientes Order By Nombre",ObConn
DO WHILE NOT ObRs.Eof
If Request.QueryString("Cliente") = ObRs ("Nombre") THEN
%>
	<option selected value="<%Response.Write ObRs ("Nombre")%>"><%Response.Write ObRs ("Nombre")%></option>
<%
SQLCliente = " Nombre = '" & ObRs ("Nombre") & "'"

ELSE
%>
	<option value="<%Response.Write ObRs ("Nombre")%>"><%Response.Write ObRs ("Nombre")%></option>
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
<table border="1" width="733" id="table1">
	<tr>
		<td width="63">Nº Cliente</td>
		<td width="118">Cliente</td>
		<td width="279">Dirección</td>
		<td width="96">Teléfono</td>
		<td width="72">Email</td>
		<td width="65">Notas</td>
	</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
IF Request.QueryString("Cliente") <> "" then
SQL = "Select * From Clientes where Nombre = '" & Request.QueryString("Cliente") & "'"
Else
SQL = "Select * From Clientes Order by Nombre"
End If
ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<tr>
		<td width="63"><a href="ModificarCliente.asp?IdCliente=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a>&nbsp;</td>
		<td width="118">&nbsp;<%Response.Write ObRs ("Nombre")%></td>
		<td width="279">&nbsp;<%Response.Write ObRs ("Direccion")%></td>
		<td width="96">&nbsp;<%Response.Write ObRs ("Telefono")%></td>
		<td width="72">&nbsp;<%Response.Write ObRs ("Email")%></td>
		<td width="65">&nbsp;<%Response.Write ObRs ("Observaciones")%></td>
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