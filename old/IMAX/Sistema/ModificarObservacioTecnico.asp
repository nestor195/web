<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<%
IF Request.Form = "" THEN
%>

<b>Modificar Observaciones del Tecnico</b><form method="POST" action="ModificarObservacioTecnico.asp" webbot-action="--WEBBOT-SELF--">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL,ObConn
%>
	<p><b>Orden: <select size="1" name="Id">
	<option selected value="<%Response.Write ObRs ("Id")%>">
	<%Response.Write ObRs ("Id")%></option>
	</select></b></p>
	<p>&nbsp;</p>
<table border="0" width="86%">
	<tr>
		<td width="335">
<b>Observaciones del T�cnico:</b><p>
<textarea rows="4" name="ObservacionTecnico" cols="34">
<%Response.Write ObRs ("ObservacionTecnico")%></textarea></td>
		<td>
<b>Observaciones Internas:</b><p>
<textarea rows="4" name="ObservacionInterna" cols="34">
<%Response.Write ObRs ("ObservacionInterna")%></textarea></td>
	</tr>
</table>
<p> <br>
<br>
<%
ObRs.Close
ObConn.Close
%><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.Form("Id")
ObRs.Open SQL, ObConn, 3, 3

ObRs ("ObservacionTecnico") = Request.Form ("ObservacionTecnico")
ObRs ("ObservacionInterna") = Request.Form ("ObservacionInterna")
ObRs.Update

ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b>
<%
END IF
%>

</body>
</html>