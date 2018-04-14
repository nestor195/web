<html>
<head>

<meta content="es-ar" http-equiv="Content-Language">
<title>Responsable</title>
<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type">


<link href="estilo.css" rel="stylesheet" type="text/css">

</head>
<body>
<%
IF Request.QueryString ("nuevoingreso") = "nuevo" then
Id = "Nuevo"
Responsable = ""
Inhabilitado = ""
Else

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "ACYP"
Sel = "SELECT * FROM Responsables Where Id = " & Request.Querystring ("AC")
ObRs.Open Sel,ObConn
Id = ObRs ("Id")
Responsable = ObRs ("Responsable")
Inhabilitado = ObRs ("Inhabilitado")
ObRs.Close
ObConn.Close

End If
%>

<p class="Tilulo">Consulta Responsable</p>
<form method="post" action="guardaraacc.asp">
<table cellpadding="2" cellspacing="0" class="tablas" style="width: 100%">
		<table style="width: 100%">
			<tr>
				<td class="celdaazul" style="width: 112px">N°</td>
				<td class="auto-style7">
				<input name="AC" type="text" readonly="readonly" value="<%Response.Write Id%>"></td>
			</tr>
		</table>
		<table style="width: 100%">
			<tr>
				<td class="celdaazul" style="width: 112px">Responsable:</td>
				<td class="auto-style7">
				<input name="AC" type="text" value="<%Response.Write Id%>"></td>
			</tr>
		</table>
		<table style="width: 100%">
			<tr>
				<td class="celdaazul" style="width: 112px">Inhabilitado</td>
				<td class="auto-style7">
				<input name="Checkbox1" type="checkbox"></td>
			</tr>
		</table>
</table>
<input name="Submit1" type="submit" value="Enviar">
</form>

</body>
</html>