<!-- insert.htm -->
<html>
<head>
	<title>Inserts Images into Database</title>
	<style>
		body, input { font-family:verdana,arial; font-size:10pt; }
	</style>
</head>
<body>

	<table border="0" align="center">
	<tr>
	<form method="POST" enctype="multipart/form-data" action="Insertequipo.asp">
	<td>&nbsp;</td><td>
		&nbsp;</td>
	</tr>
	<td>&nbsp;</td><td>
		&nbsp;</td>
	</tr>
	<td>Equipo:</td><td>

          <select size="1" name="Equipo"><%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * from Equipos order by Modelo"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof

If int(Request.QueryString("equipo")) = ObRs("Id") then
%>
		<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Modelo")%></option>
<%
else
%>
		<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Modelo")%></option>
<%
end if
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
          </select></p>
        </td></tr>
	<td>File :</td><td>
		<input type="file" name="file" size="40"></td></tr>
	<td> </td><td>
		<input type="submit" value="Submit"></td></tr>
	</form>
	</tr>
	</table>

</body>
</html>