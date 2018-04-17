<%
If Session("loginokay") = "" then
Response.redirect "login.asp"
end if
%>
<html>
<head>

<meta content="es-ar" http-equiv="Content-Language">
<title>SOLICITUD DE ACCIONES CORRECTIVAS Y PREVENTIVAS</title>

<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type">

<link href="estilo.css" rel="stylesheet" type="text/css">

</head>
<body>
<table style="width: 100%">
	<tr>
		<td><a href="listado.asp" class="Titulo2">Listado</a><br><br>
		<a href="equipos.asp" class="Titulo2">Equipos</a><br></td>
		<td><a class="Titulo2" href="logout.asp">Salir</a></td>
	</tr>
</table>

</body>
</html>