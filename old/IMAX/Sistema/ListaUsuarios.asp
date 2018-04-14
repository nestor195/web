<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>
<%
if Session("IMAX") = False then
Response.Redirect ("administrativo.asp")
End if
%>

<body>
<form method="Get" action="ListaUsuarios.asp" webbot-action="--WEBBOT-SELF--">
Nick:
	<select size="1" name="Nick">
	<option value="">Todos</option>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Usuarios Order By Nick",ObConn
DO WHILE NOT ObRs.Eof
If Request.QueryString("Nick") = ObRs ("Nick") THEN
%>
	<option selected value="<%Response.Write ObRs ("Nick")%>"><%Response.Write ObRs ("Nick")%></option>
<%
ELSE
%>
	<option value="<%Response.Write ObRs ("Nick")%>"><%Response.Write ObRs ("Nick")%></option>
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
		<td width="24">Id</td>
		<td width="136">Nick</td>
		<td width="130">Password</td>
		<td width="148">Area</td>
		<td width="148">Cliente</td>
		<td width="148">Habilitado</td>
	</tr>
<%

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
If Request.QueryString("Nick") = "" then
SQL = "Select * From Usuarios order by Id"
else
SQL = "Select * From Usuarios where Nick = '" & Request.QueryString("Nick") & "'"
end if
ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<tr>
		<td width="24">&nbsp;<a href="ModificarUsuario.asp?Id=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a></td>
		<td width="136">&nbsp;<%Response.Write ObRs ("Nick")%></td>
		<td width="130">&nbsp;<%Response.Write ObRs ("Password")%></td>
		<td width="148">&nbsp;<%Response.Write ObRs ("Area")%></td>
		<td width="148">&nbsp;<%Response.Write ObRs ("Cliente")%></td>
		<td width="148">&nbsp;<%Response.Write ObRs ("Habilitado")%></td>
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