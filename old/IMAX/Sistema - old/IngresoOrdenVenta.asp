<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
</head>

<body>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
sql = "select * from Ordenes where Id = 1"
ObRs.Open sql, ObConn


cliente = ObRs ("Cliente")
equipo = ObRs ("Equipo")
serie = ObRs ("Serie")
accesorios =ObRs ("Accesorios")
usuarioingreso = ObRs ("UsuarioIngreso")
usuarioestado = ObRs ("UsuarioEstado")
observacioningreso = ObRs ("ObservacionIngreso")

ObRs.Close
ObConn.Close
%>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Ordenes",ObConn, 3, 3

ObRs.AddNew
ObRs ("Cliente") = cliente
ObRs ("Equipo") = equipo
ObRs ("Serie") = serie
ObRs ("Estado") = 19
ObRs ("Accesorios") = accesorios
ObRs ("UsuarioIngreso") = usuarioingreso
ObRs ("UsuarioEstado") = usuarioestado
ObRs ("FechaIngreso") = DATE
ObRs ("FechaEstado") = DATE
ObRs ("ObservacionIngreso") = observacioningreso
ObRs.Update
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Order By Id"
ObRs.Open SQL, ObConn
DO WHILE NOT ObRs.Eof
Orden = ObRs("Id")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = 1"
ObRs.Open SQL,ObConn, 3, 3
DO WHILE NOT ObRs.Eof

ObRs ("Orden") = orden

ObRs.Update
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>

<b>Datos Ingresados</b><p><b>
<a target="_blank" href="ConsultaDeOrden.asp?Id=<%Response.Write orden%>">Imprimir</a></b>
</p>

</body>

</html>