<html>

<head>
<meta http-equiv="Content-Language" content="es-ar">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Cambia Equipo Orden</title>
</head>
<%
If Request.QueryString("Cambiar") = 1 Then

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * from Ordenes where Id = " & Request.QueryString("Orden")
ObRs.Open SQL,ObConn, 3, 3

ObRs ("Equipo") = Request.QueryString("Equipo")

ObRs.Update

ObRs.Close
ObConn.Close

Response.Redirect ("ConsultaDeOrden.asp?Id=" & Request.QueryString("Orden"))

End If
%>

<body>
<%
Orden = Request.QueryString("Orden")
Equipo = Request.QueryString("Equipo")

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Equipos Where Id = " & Equipo
ObRs.Open SQL, ObConn
IdMarca = ObRs("Marca")
IdTiposdeEquipos = ObRs("Tipo")
Modelo = ObRs("Modelo")
ObRs.Close
ObConn.Close
%>
<p>
<a href="SeleccionarEquipo.asp?Orden=<%Response.Write Orden%>&Pagina=CambiaEquipoOrden.asp">
Seleccionar Equipo de la Orden <%Response.Write orden%>.</a></p>
<p>modelo: <%Response.Write Modelo%>.</p>
<p><a href="CambiaEquipoOrden.asp?Orden=<%Response.Write orden%>&Equipo=<%Response.Write Equipo%>&Cambiar=1">Cambiar</a></p>
</body>

</html>