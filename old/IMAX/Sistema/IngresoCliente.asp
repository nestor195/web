<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Ingreso Cliente</title>
</head>
<%
Pagina = Request.QueryString ("Pagina")
Cliente = Request.QueryString ("Cliente")
Equipo = Request.QueryString ("Equipo")
%>

<body>
<%
IF Request.Form = "" THEN
%>
<p>Ingreso Cliente</p>
<form method="POST" action="IngresoCliente.asp?Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Pagina=<%Response.Write Pagina%>">
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
    <tr>
      <td width="100%">Nombre:<input type="text" name="Nombre" size="31"></td>
    </tr>
    <tr>
      <td width="100%">Dirección:<input type="text" name="Direccion" size="30"></td>
    </tr>
    <tr>
      <td width="100%">Teléfono:<input type="text" name="Telefono" size="31"></td>
    </tr>
    <tr>
      <td width="100%">Email:<input type="text" name="Email" size="29"></td>
    </tr>
    <tr>
      <td width="100%">Tipo de Cliente:<select size="1" name="TipoCliente">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "TipoCliente",ObConn
DO WHILE NOT ObRs.Eof
%>
		<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("TipoCliente")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
		</select></td>
    </tr>
    <tr>
      <td width="100%">Observaciones:</td>
    </tr>
    <tr>
      <td width="100%"><textarea rows="3" name="Observacion" cols="32"></textarea></td>
    </tr>
  </table>
  <p><input type="submit" value="Enviar" name="B1"></p>
</form>
<p>&nbsp;</p>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Clientes",ObConn, 3, 3

ObRs.AddNew
ObRs ("Nombre") = Request.Form ("Nombre")
ObRs ("Direccion") = Request.Form ("Direccion")
ObRs ("Telefono") = Request.Form ("Telefono")
ObRs ("Email") = Request.Form ("Email")
ObRs ("Observaciones") = Request.Form ("Observacion")
ObRs ("TipoCliente") = Request.Form ("TipoCliente")
ObRs.Update

ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Clientes Order By Id"
ObRs.Open SQL, ObConn
DO WHILE NOT ObRs.Eof
Ultimo = ObRs("Id")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<%
Response.Redirect (Pagina & "?Cliente=" & Ultimo & "&Equipo=" & Equipo)
END IF
%>

</body>

</html>