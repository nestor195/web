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
<b>Ingreso de Equipo</b><form method="POST" action="IngresoTipoDeEquipo.asp" webbot-action="--WEBBOT-SELF--">
	<p>Tipo de Equipo: <input type="text" name="Tipo" size="30"><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "TiposDeEquipos",ObConn, 3, 3

ObRs.AddNew
ObRs ("Tipo") = Request.Form ("Tipo")
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