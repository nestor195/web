<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<SCRIPT language=javascript type="text/Jscript">
window.opener.history.go()
</SCRIPT>
</head>

<body>
<%
IF Request.Form = "" THEN
%>
<b>Ingreso de Equipo</b><form method="POST" action="IngresoEquipo%20-%20old.asp" webbot-action="--WEBBOT-SELF--">
	<p>Tipo: <select size="1" name="Tipo">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "TiposDeEquipos",ObConn
DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Tipo")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select> <a target="_parent" href="IngresoTipoDeEquipo.asp">Nuevo</a><br>
	Marca: <select size="1" name="Marca">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Marcas",ObConn
DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Marca")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select> <a target="_parent" href="IngresoMarca.asp">Nuevo</a><br>
	Modelo: <input type="text" name="Modelo" size="20"><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Equipos",ObConn, 3, 3

ObRs.AddNew
ObRs ("Tipo") = Request.Form ("Tipo")
ObRs ("Marca") = Request.Form ("Marca")
ObRs ("Modelo") = Request.Form ("Modelo")
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