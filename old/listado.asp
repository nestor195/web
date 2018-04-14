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
<p class="Tilulo">SOLICITUD DE ACCIONES CORRECTIVAS Y PREVENTIVAS</p>
<form action="" method="get">
	Buscar <input name="Dato" type="text"> en <select name="Campo">
	<option></option>
	<option selected="" value="Id">N°</option>
	<option value="Estado">Estado</option>
	<option value="NoConformidad">No Conformidad</option>
	<option value="Fecha">Fecha</option>
	<option value="Area">Area</option>
	<option value="Solicita">Solicitante</option>
	</select> <input name="Ir" type="submit" value="Ir"> <span class="Titulo2">
	<a href="consulta.asp?nuevoingreso=nuevo">Agregar nuevo</a></span></form>
<table cellpadding="2" cellspacing="0" class="tablas" style="width: 100%">
	<tr class="celdaazul">
		<td class="celdaazul">N°</td>
		<td class="celdaazul">Estado</td>
		<td class="celdaazul">No conformidad</td>
		<td class="celdaazul">Fecha</td>
		<td class="celdaazul">Área</td>
		<td class="celdaazul">Solicitante</td>
	</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "ACYP"
SQL = "SELECT Planilla.Id, Planilla.Estado, Estados.Estado, Planilla.NoConformidad, Planilla.Fecha, Areas.Area, Responsables.Responsable"
SQL = SQL & " FROM ((Planilla INNER JOIN Estados ON Planilla.Estado = Estados.ID) INNER JOIN Areas ON Planilla.Area = Areas.Id) INNER JOIN Responsables ON Planilla.Solicita = Responsables.Id"

Select Case Request.Querystring ("Campo")

Case "Id"
If IsNumeric(Request.QueryString ("Dato")) Then
SQL = SQL & " Where Planilla.Id = " & Request.QueryString ("Dato") & " Order by Planilla.Id desc"
End If
Case "Estado"
SQL = SQL & " Where Estados.Estado Like '%" & Request.QueryString ("Dato") & "%' Order by Planilla.Id desc"

End Select

ObRs.Open SQL,ObConn
Do While ObRs.EOF = false
%>

	<tr class="celdablanca">
		<td  class="celdablanca">
		<a href="consulta.asp?AC=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a></td>
		<td class="celdablanca"><%Response.Write ObRs ("Estado")%></td>
		<td class="celdablanca"><%Response.Write ObRs ("NoConformidad")%></td>
		<td class="celdablanca"><%Response.Write ObRs ("Fecha")%></td>
		<td class="celdablanca"><%Response.Write ObRs ("Area")%></td>
		<td class="celdablanca"><%Response.Write ObRs ("Responsable")%></td>
	</tr>
<%
ObRs.MoveNext
Loop
ObRs.Close
ObConn.Close
%>

</table>

</body>
</html>