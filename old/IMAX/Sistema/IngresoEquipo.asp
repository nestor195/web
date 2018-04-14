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
TipoEquipo = Request.QueryString ("TipoEquipo")
Marca = Request.QueryString ("Marca")
%>

<body>
<%
IF Request.Form = "" THEN
%>
<p>Ingreso Equipo</p>
<form method="POST" action="IngresoEquipo.asp?Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Pagina=<%Response.Write Pagina%>&TipoEquipo=<%Response.Write TipoEquipo%>&Marca=<%Response.Write Marca%>">
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
    <tr>
      <td width="100%">Marca:<select size="1" name="Marca">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Marcas",ObConn
DO WHILE NOT ObRs.Eof
If Marca = ObRs ("Marca") Then
%>
		<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Marca")%></option>
<%
Else
%>
		<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Marca")%></option>
<%
End If
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
		</select> <a href="IngresoMarca.asp?Pagina=<%Response.Write Pagina%>&Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&TipoEquipo=<%Response.Write TipoEquipo%>&Pagina2=IngresoEquipo.asp">Nuevo</a></td>
    </tr>
    <tr>
      <td width="100%">Tipo de Equipo:<select size="1" name="TipoEquipo">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "TiposDeEquipos",ObConn
DO WHILE NOT ObRs.Eof
If TipoEquipo = ObRs ("Tipo") Then
%>
		<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Tipo")%></option>
<%
Else
%>
		<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Tipo")%></option>
<%
End IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
		</select> <a href="IngresoTipoEquipo.asp?Pagina=<%Response.Write Pagina%>&Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Marca=<%Response.Write Marca%>&Pagina2=IngresoEquipo.asp">Nuevo</a></td>
    </tr>
    <tr>
      <td width="100%">Modelo:<input type="text" name="Modelo" size="22"></td>
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
ObRs.Open "Equipos",ObConn, 3, 3

ObRs.AddNew
ObRs ("Marca") = Request.Form ("Marca")
ObRs ("Tipo") = Request.Form ("TipoEquipo")
ObRs ("Modelo") = Request.Form ("Modelo")
ObRs.Update

ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Equipos Order By Id"
ObRs.Open SQL, ObConn
DO WHILE NOT ObRs.Eof
Ultimo = ObRs("Id")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<%
Response.Redirect (Pagina & "?Cliente=" & Cliente & "&Equipo=" & Ultimo)
END IF
%>

</body>

</html>