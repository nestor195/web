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
Marca = Request.QueryString ("Marca")
Pagina2 = Request.QueryString ("Pagina2")
%>

<body>
<%
IF Request.Form = "" THEN
%>
<b>Ingreso de Tipo de Equipo</b><form method="POST" action="IngresoTipoEquipo.asp?Pagina=<%Response.Write Pagina%>&Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Marca=<%Response.Write Marca%>&Pagina2=<%Response.Write Pagina2%>">
	<p>Tipo de Equipo: <input type="text" name="TipoEquipo" size="30"><br>
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
ObRs ("Tipo") = Request.Form ("TipoEquipo")
ObRs.Update

ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From TiposDeEquipos Order By Id"
ObRs.Open SQL, ObConn
DO WHILE NOT ObRs.Eof
Ultimo = ObRs("Tipo")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<%
Response.Redirect (Pagina2 & "?Cliente=" & Cliente & "&TipoEquipo=" & Ultimo & "&Pagina=" & Pagina& "&Equipo=" & Equipo & "&Marca=" & Marca)
END IF
%>

</body>
</html>