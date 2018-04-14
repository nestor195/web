<%
IF Request.Form("Empresa") <> 0 THEN
Empresa = Request.Form("Empresa")
Else
Response.Redirect("default.asp")
End If
%>


<%
IF Request.Form("pass") <> "" THEN
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Gestion"
SQL = "Select Id, Nombre From Usuarios where Contrasenia = '" & Request.Form("pass") & "'"
ObRs.Open SQL,ObConn
	If NOT ObRs.Eof Then
	Session("USUARIO") = ObRs ("Id")
	End If
ObRs.Close
ObConn.Close
End If

If Session("USUARIO") = "" then
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//ES" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Gestion - Login</title>
<link href="estilos.css" rel="stylesheet" type="text/css" />
</head>
<body>
<form method="POST" action="principal.asp" webbot-action="--WEBBOT-SELF--">
<p>
  Contraseña: <input type="password" name="pass" size="20">
   <input type="text" value=<%response.write Empresa%> name="Empresa" class="Oculto"></p>
	<p>
    <input name="Button1" type="submit" value="INICIAR" />&nbsp;</p>
</form>
</body>
</html>

<%
response.write "</body>"
response.write "</html>"
Else
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Gestion"
SQL = "Select Nombre From Usuarios where Id = " & Session("USUARIO")
ObRs.Open SQL,ObConn
Usuario = ObRs("Nombre")
ObRs.Close
ObConn.Close
%>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Gestion"
SQL = "Select NombreCorto From Empresas where Id = " & Request.Form("Empresa")
ObRs.Open SQL,ObConn
NombreEmpresa = ObRs("NombreCorto")
ObRs.Close
ObConn.Close
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>EMPRESA - Iniciado</title>
<link href="estilos.css" rel="stylesheet" type="text/css" />
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" type="text/css" href="ddlevelsfiles/ddlevelsmenu-base.css" />
<link rel="stylesheet" type="text/css" href="ddlevelsfiles/ddlevelsmenu-topbar.css" />
<link rel="stylesheet" type="text/css" href="ddlevelsfiles/ddlevelsmenu-sidebar.css" />

<script type="text/javascript" src="ddlevelsfiles/ddlevelsmenu.js">

/***********************************************
* All Levels Navigational Menu- (c) Dynamic Drive DHTML code library (http://www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
***********************************************/

</script>
</head>

<body>

<p><span class="Titulo"><%Response.Write NombreEmpresa%></span></p>
<p>Bienvenido <%Response.write Usuario%></p>

<div id="ddtopmenubar" class="mattblackmenu">
<ul>
<li><a href="cpedido.asp" rel="ddsubmenu1">FACTURACION</a></li>
<li><a href="#" rel="ddsubmenu2">VENTA</a></li>
</ul>
</div>

<a class="animateddrawer" id="ddtopmenubar-mobiletoggle" href="#">
<span></span>
</a>

<script type="text/javascript">
ddlevelsmenu.setup("ddtopmenubar", "topbar") //ddlevelsmenu.setup("mainmenuid", "topbar|sidebar")
</script>


<!--HTML for the Drop Down Menus associated with Top Menu Bar-->
<!--They should be inserted OUTSIDE any element other than the BODY tag itself-->
<!--A good location would be the end of the page (right above "</BODY>")-->

<!--Top Drop Down Menu 1 HTML-->


<ul id="ddsubmenu1" class="ddsubmenustyle">
</ul>


<!--Top Drop Down Menu 2 HTML-->

<ul id="ddsubmenu2" class="ddsubmenustyle">
<li><a href="#">Facturacion...</a>
	<ul>
	<li><a href="#">Ver facturas</a></li>
	<li><a href="#">Carga masiva de remitos anulados</a></li>
	<li><a href="#">Listado de remitos anulados</a></li>
	<li><a href="#">ver remitos</a></li>
	<li><a href="#">ver pedidos</a></li>
	<li><a href="#">ver devoluciones</a></li>
	<li><a href="#">anular pedidos</a></li>
	<li><a href="#">afectacion de remitos a facturas</a></li>
	<li><a href="#">listado de remitos</a></li>
	<li><a href="#">listado de pedidos</a></li>
	<li><a href="#">listado de devoluciones</a></li>
	<li><a href="#">importacion de archivos elecronicos</a></li>
	</ul>
</li>
<li><a href="#">Cuentas Correntes...</a></li>
<li><a href="#">Listado de Ventas...</a></li>
<li><a href="#">Impuestos...</a></li>

</ul>

<br>

</body>

</html>
<%
End If
%>

