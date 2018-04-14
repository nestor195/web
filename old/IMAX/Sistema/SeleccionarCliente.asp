<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Seleccionar Cliente</title>
</head>
<%
Pagina = Request.QueryString ("Pagina")
Cliente = Request.QueryString ("Cliente")
Equipo = Request.QueryString ("Equipo")

%>
<body>

<p>Seleccionar Cliente: </p>
<p><a href="IngresoCliente.asp?Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Pagina=<%Response.Write Pagina%>">Nuevo Cliente</a></p>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
    <tr>
      <td width="96%">
      <form method="POST" action="seleccionarcliente.asp?Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Pagina=<%Response.Write Pagina%>">
        <p>Ingrese Nombre o Parte del Nombre:
      <input type="text" name="Palabra" size="39"><input type="submit" value="Enviar" name="B1"></p>
      </form>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
        <tr>
          <td width="20%">Nombre</td>
          <td width="26%">Dirección</td>
          <td width="15%">Teléfono</td>
          <td width="17%">Email</td>
          <td width="12%">Tipo</td>
          <td width="20%">Observaciones<%Response.Write Request.QueryString ("a")%></td>
        </tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL="Select * From Clientes Where Nombre Like '%" & Request.Form ("Palabra") & "%' Order By Nombre"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
       <tr>
          <td width="20%">&nbsp;<a href="<%Response.Write Pagina%>?Cliente=<%Response.Write ObRs("Id")%>&Equipo=<%Response.Write Equipo%>"><%Response.Write ObRs("Nombre")%></a></td>
          <td width="26%">&nbsp;<%Response.Write ObRs("Direccion")%></td>
          <td width="15%">&nbsp;<%Response.Write ObRs("Telefono")%></td>
          <td width="17%">&nbsp;<%Response.Write ObRs("Email")%></td>
          <td width="12%">&nbsp;<%Response.Write ObRs("TipoCliente")%></td>
          <td width="20%">&nbsp;<%Response.Write ObRs("Observaciones")%></td>
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