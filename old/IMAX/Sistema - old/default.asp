<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="50%" align="left" valign="top">
<form method="GET" action="ConsultaDeOrden.asp">
<p>
  Consulta de
  Orden: <input type="text" name="Id" size="8"><input type="submit" value="Enviar" name="B1"></p>
</form>
<p>
<a target="_parent" href="IngresoOrden.asp">Ingresar Orden</a></p>
<p>
<a target="_parent" href="ListaOrdenEstado.asp">Lista de Ordenes por Estado</a></p>
<p>
<a href="ListaOrdenEquipo.asp">Lista de Ordenes por Equipo</a></p>
<p><a href="ListaCliente.asp">Lista de Clientes</a></p>
<p><a href="ListaOrdenSerie.asp">Búsqueda de Orden por Numero de Serie</a></p>
    </td>
    <td width="50%" align="left" valign="top"><span lang="es">
    <a href="IngresoPago.asp">Ingreso de Pago</a></span><p><a href="ListaItems.asp">Lista de 
    Ítems</a></p>
    <p><a href="ListaEquipos.asp">Lista de Equipos</a><p>
    <span lang="es"><a href="IngresoVenta.asp">Ingreso Venta</a></span><p>
    <span lang="es"><a href="ListaVentas.asp">Lista de Ventas</a></span></td>
  </tr>
</table>
<p>&nbsp;</p>
<p><span lang="es">Tareas <a href="IngresoTarea.asp">Nueva</a></span></p>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="61%">
  <tr>
    <td width="17%">&nbsp;</td>
    <td width="153%"><span lang="es">Tarea</span></td>
    <td width="15%"><span lang="es">Completado</span></td>
  </tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Tareas Where Completado < 100 Order By Completado"
ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
%>

  <tr>
    <td width="17%"><span lang="es"><a href="ModificarTarea.asp?IdTarea=<%Response.Write ObRs ("IdTarea")%>">Modificar</a></span></td>
    <td width="153%">&nbsp;<%Response.Write ObRs ("Tarea")%></td>
    <td width="15%"><span lang="es">&nbsp;<%Response.Write ObRs ("Completado")%>%</span></td>
  </tr>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
</table>
</body>

</html>