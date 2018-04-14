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
<b>Ingreso Banco</b><form method="POST" action="IngresoBanco.asp" webbot-action="--WEBBOT-SELF--">
	<p>Tipo de Banco: <input type="text" name="Banco" size="21"><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "BancoCheque",ObConn, 3, 3

ObRs.AddNew
ObRs ("BancoCheque") = Request.Form ("Banco")
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
