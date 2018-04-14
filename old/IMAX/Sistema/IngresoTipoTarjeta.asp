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
<b>Ingreso de Tipo de TarJeta</b><form method="POST" action="IngresoTipoTarjeta.asp" webbot-action="--WEBBOT-SELF--">
	<p>Tipo de Tarjeta: <input type="text" name="TipoTarjeta" size="21"><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "TipoTarjeta",ObConn, 3, 3

ObRs.AddNew
ObRs ("TipoTarjeta") = Request.Form ("TipoTarjeta")
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
