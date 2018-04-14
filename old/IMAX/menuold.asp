<html>

<head>
<meta http-equiv="Content-Language" content="es-ar">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
<base target="principal">
</head>

<body text="#FFFFFF" bgcolor="#000000">
<%
if Request.QueryString("Logout") = 1 then
Session ("Session") = ""
end if
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios where Nick = '" & Request.Form("User") &"' and Password = '" & Request.Form("Password") & "'"
ObRs.Open SQL,ObConn
If not ObRs.Eof then
Session ("Session") = ObRs("Id")
Usuario = ObRs("Nick")
end if
ObRs.Close
ObConn.Close

IF Session("Session") <> "" THEN
%>
<p style="margin-bottom: -15"><font face="Arial" size="2">Bienvenido:</font></p>
<p style="margin-bottom: -15"><font face="Arial" size="2"><%Response.Write Usuario%>
</font>
</p>
<p style="margin-bottom: -15"><b><font size="1" face="Arial">
<a target="_self" href="menu.asp?logout=1"><font color="#FFFFFF">Logout</font></a></font></b></p>
<p style="margin-bottom: -15">&nbsp;</p>
<p style="margin-bottom: -15"><a target="principal" href="inicio.asp">
<font color="#FFFFFF" face="Arial" size="2">Inicio</font></a></p>
<p style="margin-bottom: -15"><font face="Arial" size="2">*<a target="principal" href="listado.asp?Estado=4"><font color="#FFFFFF">Ordenes Listas</font></a></font></p>
<p style="margin-bottom: -15"><font face="Arial" size="2">*<a target="principal" href="listado.asp?Estado=2"><font color="#FFFFFF">Ordenes 
Presupuestadas</font></a></font></p>
<p style="margin-bottom: -15"><font face="Arial" size="2">*<a target="principal" href="listado.asp?Estado=3"><font color="#FFFFFF">Ordenes 
Confirmadas</font></a></font></p>
<p style="margin-bottom: -15"><font face="Arial" size="2">*<a target="principal" href="listado.asp?Estado=1"><font color="#FFFFFF">Ordenes No Vistas</font></a></font></p>
<p style="margin-bottom: -15"><font face="Arial" size="2">*<a href="listado.asp?Estado=5"><font color="#FFFFFF">Ordenes 
Entregadas</font></a></font></p>
<p style="margin-bottom: -15"><font face="Arial" size="2">*<a href="listado.asp?Estado=7"><font color="#FFFFFF">Ordenes 
Sin Reparacion</font></a></font></p>
<p style="margin-bottom: -15">Busqueda por numero de Orden</p>
<form name="form3" method="get" action="listado.asp">
  <p>
    <label>
      <input name="Orden" type="text" id="Serie0" size="15">
    </label>
    <input type="submit" name="button3" id="button3" value="Enviar">
  </p>
</form>
<p style="margin-bottom: -15">Busqueda por numero de serie</p>
<form name="form1" method="get" action="listado.asp">
  <p>
    <label>
      <input name="Serie" type="text" id="Serie" size="15">
    </label>
    <input type="submit" name="button" id="button" value="Enviar">
  </p>
</form>
<form name="form2" method="get" action="listado.asp">
  <p><label>
    Busqueda por numero de referencia<br>
    <input name="referencia" type="text" id="referencia" value="" size="15">
  </label>
  <label>
    <input type="submit" name="button2" id="button2" value="Enviar">
  </label>
	</p>
</form>
<%
ELSE
%>
<form method="POST" action="menuold.asp" target="_self" webbot-action="--WEBBOT-SELF--">
  <p style="margin-bottom: -15">Usuario</p>
  <p style="margin-bottom: -15"><input type="text" name="User" size="17"></p>
  <p style="margin-bottom: -15">Contraseña</p>
  <p style="margin-bottom: -15">
  <input type="password" name="password" size="17"></p>
  <p style="margin-bottom: -15"><input type="submit" value="Enviar" name="B1"></p>

  <p style="margin-bottom: -15">&nbsp;</p>
</form>

<%
END IF
%>
</body>

</html>