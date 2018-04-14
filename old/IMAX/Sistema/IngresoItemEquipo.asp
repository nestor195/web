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
<b>Ingreso de Ítem a un Equipo</b><form method="POST" action="IngresoItemEquipo.asp" webbot-action="--WEBBOT-SELF--">
	<p><span lang="es">Ítem</span>: <select size="1" name="Item">
	<option value="<%Response.Write Request.QueryString("IdItem")%>">
	<%Response.Write Request.QueryString("IdItem")%></option>
	</select><br>
	<span lang="es">Equipo</span>: <select size="1" name="Equipo">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Equipos Order By Modelo"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Modelo")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select> <a target="_parent" href="IngresoEquipo.asp">Nuevo</a><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "EquipoItem",ObConn, 3, 3

ObRs.AddNew
ObRs ("Equipo") = Request.Form ("Equipo")
ObRs ("Item") = Request.Form ("Item")
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