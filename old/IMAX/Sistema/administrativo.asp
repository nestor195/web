<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
</head>

<body>
<%
if Session("IMAX") = True then
%>
Administrativo<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="50%"><a href="estadistica.asp">Estadistica
</a>
<p><a href="ListadoOrdenes.asp">Listado de Ordenes por Cliente</a> </p>

<p><a href="EstadoDeCuentas.asp">Estado de Cuentas</a></p>
	<p><a href="Cantidad_de_equipos_ingresados.asp">Cantidad de Equipos 
	Ingresados</a></p>
	<p><a href="Facturado.asp">Facturado</a></p>
	<p><a href="ListaVentasporfecha.asp">Lista de Ventas por Fecha</a></p>

    </td>
    <td width="50%"><a href="IngresoUsuario.asp">Ingreso Usuario</a><p>
	<a href="ListaUsuarios.asp">Lista de Usuarios</a><p>
	<a href="ListadeUtilidad.asp">Lista de Utilidades</a></td>
  </tr>
</table>

<%
Else
%>
<!--#include file="datos.inc"-->
<%
Session("IMAX") = true
Response.Redirect ("administrativo.asp")
else
Session("IMAX") = false
end if
%>
<form method="POST" action="administrativo.asp" webbot-action="--WEBBOT-SELF--">
<p>
  Password:<input type="password" name="pass" size="20"></p>
<p>
  <input type="submit" value="Enviar" name="B1"><input type="reset" value="Restablecer" name="B2"></p>
</form>
<%
End if
%>
</body>

</html>