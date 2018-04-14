<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Seleccionar Cliente</title>
</head>
<%
Pagina = Request.QueryString ("Pagina")
Cliente = Request.QueryString ("Cliente")
Equipo = Request.QueryString ("Equipo")
Orden = Request.QueryString ("Orden")

%>
<body>

<p>Seleccionar Equipo: </p>
<p><a href="IngresoEquipo.asp?Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Pagina=<%Response.Write Pagina%>">Nuevo Equipo</a></p>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
    <tr>
      <td width="96%">
      <form method="POST" action="seleccionarEquipo.asp?Orden=<%Response.Write Orden%>&Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Pagina=<%Response.Write Pagina%>">
        <p>Ingrese Modelo o Parte del Modelo:
      <input type="text" name="Palabra" size="39"><input type="submit" value="Enviar" name="B1"></p>
      </form>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
        <tr>
          <td width="35%">Modelo</td>
          <td width="29%">Marca</td>
          <td width="35%">Tipo de Equipo</td>
        </tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL="Select * From Equipos Where Modelo Like '%" & Request.Form ("Palabra") & "%' Order By Modelo"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
       <tr>
          <td width="35%">&nbsp;<a href="<%Response.Write Pagina%>?Orden=<%Response.Write Orden%>&Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write ObRs("Id")%>"><%Response.Write ObRs("Modelo")%></a></td>
          <td width="29%">&nbsp;<%Response.Write ObRs("Marca")%></td>
          <td width="35%">&nbsp;<%Response.Write ObRs("Tipo")%></td>
        </tr>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
        </table>
      </td>
    </tr>
  </table>

</body>

</html>