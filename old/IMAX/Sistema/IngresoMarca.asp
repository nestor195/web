<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>
<%
Pagina = Request.QueryString ("Pagina")
Cliente = Request.QueryString ("Cliente")
Equipo = Request.QueryString ("Equipo")
TipoEquipo = Request.QueryString ("TipoEquipo")
Pagina2 = Request.QueryString ("Pagina2")
%>

<body>
<%
IF Request.Form = "" THEN
%>
<b>Ingreso de Marca</b><form method="POST" action="IngresoMarca.asp?Pagina=<%Response.Write Pagina%>&Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&TipoEquipo=<%Response.Write TipoEquipo%>&Pagina2=<%Response.Write Pagina2%>">
	<p>Marca: <input type="text" name="Marca" size="30"><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Marcas",ObConn, 3, 3

ObRs.AddNew
ObRs ("Marca") = Request.Form ("Marca")
ObRs.Update

ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Marcas Order By Id"
ObRs.Open SQL, ObConn
DO WHILE NOT ObRs.Eof
Ultimo = ObRs("Marca")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<%
Response.Redirect (Pagina2 & "?Cliente=" & Cliente & "&Marca=" & Ultimo & "&Pagina=" & Pagina& "&Equipo=" & Equipo & "&TipoEquipo=" & TipoEquipo)
END IF
%>

</body>
</html>