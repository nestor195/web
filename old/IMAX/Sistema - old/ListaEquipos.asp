<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<form method="Get" action="ListaEquipos.asp" webbot-action="--WEBBOT-SELF--">
Nº Equipo:
	<select size="1" name="Modelo">
	<option value="">Todos</option>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Equipos Order By Modelo",ObConn
DO WHILE NOT ObRs.Eof
If Request.QueryString("Modelo") = ObRs ("Id") THEN
%>
	<option selected value="<%Response.Write ObRs ("Modelo")%>"><%Response.Write ObRs ("Modelo")%></option>
<%
SQLCliente = " Modelo = '" & ObRs ("Modelo") & "'"

ELSE
%>
	<option value="<%Response.Write ObRs ("Modelo")%>"><%Response.Write ObRs ("Modelo")%></option>
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
<table border="1" width="672" id="table1">
	<tr>
		<td width="24">Nº Equipo</td>
		<td width="136">Modelo</td>
		<td width="130">Marca</td>
		<td width="148">Tipo</td>
	</tr>
<%

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
If Request.QueryString("Modelo") = "" then
SQL = "Select * From Equipos order by Modelo"
else
SQL = "Select * From Equipos where Modelo = '" & Request.QueryString("Modelo") & "'"
end if
ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<tr>
		<td width="24"><a href="ModificarEquipo.asp?IdEquipo=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a>&nbsp;</td>
		<td width="136">&nbsp;<%Response.Write ObRs ("Modelo")%></td>
		<td width="130">&nbsp;<%Response.Write ObRs ("Marca")%></td>
		<td width="148">&nbsp;<%Response.Write ObRs ("Tipo")%></td>
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