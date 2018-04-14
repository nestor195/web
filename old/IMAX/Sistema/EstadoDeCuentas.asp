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
Else
Response.Redirect ("administrativo.asp")
End if
%>
<p>Cuentas</p>

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="58%" id="AutoNumber1">
  <tr>
    <td width="46%">Cliente</td>
    <td width="24%">Se debe</td>
    <td width="21%">Nos deben</td>
    <td width="32%">&nbsp;</td>
    <td width="24%">&nbsp;</td>
  </tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Clientes Where Cuenta <> 0"
ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
  <tr>
    <td width="46%">&nbsp;<%Response.Write ObRs("Nombre")%></td>
<%
If ObRs("Cuenta") < 0 then
%>
    <td width="24%">&nbsp;<%Response.Write -ObRs("Cuenta")%></td>
    <td width="21%">&nbsp;</td>
<%
Else
%>
    <td width="24%">&nbsp;</td>
    <td width="21%">&nbsp;<%Response.Write ObRs("Cuenta")%></td>
<%
End If
%>
    <td width="32%"><a href="IngresoPagoaCuenta.asp?Cliente=<%Response.Write ObRs("Id")%>">Ingreso Pago</a></td>
    <td width="24%"><a href="IngresoCobroaCuenta.asp?Cliente=<%Response.Write ObRs("Id")%>">Ingreso Cobro</a></td>
  </tr>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
  <tr>
    <td width="46%">Total</td>
    <td width="24%">&nbsp;</td>
    <td width="21%">&nbsp;</td>
    <td width="32%">&nbsp;</td>
    <td width="24%">&nbsp;</td>
  </tr>
</table>

</body>

</html>