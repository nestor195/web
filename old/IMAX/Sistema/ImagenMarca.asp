<!-- insert.htm -->
<html>
<head>
	<title>Inserts Images into Database</title>
	<style>
		body, input { font-family:verdana,arial; font-size:10pt; }
	</style>
</head>
<body>
	<p align="center">
		<b>Inserting Binary Data into Database</b><br>
		<a href="show.asp">To see inserted data click here</a>
	</p>
	
	<table border="0" align="center">
	<tr>
	<form method="POST" enctype="multipart/form-data" action="Insert.asp">
	<td>&nbsp;</td><td>
		&nbsp;</td>
	</tr>
	<td>&nbsp;</td><td>
		&nbsp;</td>
	</tr>
	<td>Marca:</td><td>

          <select size="1" name="Marca"><%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Marcas",ObConn
DO WHILE NOT ObRs.Eof
%>
		<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Marca")%></option>
<%
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