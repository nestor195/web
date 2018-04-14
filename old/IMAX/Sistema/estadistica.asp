<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
</head>
<%
Dim mesini(12)
Dim mesfin(12)

mesini(1)="#12/31/2008#"
mesfin(1)="#2/01/2009#"

mesini(2)="#1/31/2009#"
mesfin(2)="#3/01/2009#"

mesini(3)="#2/28/2009#"
mesfin(3)="#4/01/2009#"

mesini(4)="#3/31/2009#"
mesfin(4)="#5/01/2009#"

mesini(5)="#4/30/2009#"
mesfin(5)="#6/01/2009#"

mesini(6)="#5/31/2009#"
mesfin(6)="#7/01/2009#"

mesini(7)="#6/30/2009#"
mesfin(7)="#8/01/2009#"

mesini(8)="#7/31/2009#"
mesfin(8)="#9/01/2009#"

mesini(9)="#8/31/2009#"
mesfin(9)="#10/01/2009#"

mesini(10)="#9/30/2009#"
mesfin(10)="#11/01/2009#"

mesini(11)="#10/31/2009#"
mesfin(11)="#12/01/2009#"

mesini(12)="#11/30/2009#"
mesfin(12)="#1/01/2010#"

%>
<%
If Request.QueryString("Mes") = "" THEN
i = 1
else
select case Request.QueryString("Mes")
case 1
i = 1
case 2
i = 2
case 3
i = 3
case 4
i = 4
case 5
i = 5
case 6
i = 6
case 7
i = 7
case 8
i = 8
case 9
i = 9
case 10
i = 10
case 11
i = 11
case 12
i = 12
end select
end if
%>
<body>
<%
if Session("IMAX") = True then
%>

<form method="GET" action="estadistica.asp" webbot-action="--WEBBOT-SELF--">
<p><font face="Arial">Estadísticas 
  mes <select size="1" name="Mes">
  <option value="1" <%if i =1 then response.write "selected"%>>Enero</option>
  <option value="2" <%if i =2 then response.write "selected"%>>Febrero</option>
  <option value="3" <%if i =3 then response.write "selected"%>>Marzo</option>
  <option value="4" <%if i =4 then response.write "selected"%>>Abril</option>
  <option value="5" <%if i =5 then response.write "selected"%>>Mayo</option>
  <option value="6" <%if i =6 then response.write "selected"%>>Junio</option>
  <option value="7" <%if i =7 then response.write "selected"%>>Julio</option>
  <option value="8" <%if i =8 then response.write "selected"%>>Agosto</option>
  <option value="9" <%if i =9 then response.write "selected"%>>Septiembre</option>
  <option value="10" <%if i =10 then response.write "selected"%>>Octubre</option>
  <option value="11" <%if i =11 then response.write "selected"%>>Noviembre</option>
  <option value="12" <%if i =12 then response.write "selected"%>>Diciembre</option>
  </select></font></p>
  <p><input type="submit" value="Enviar" name="B1"></p>
</form>

<p>&nbsp;</p>

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="1065" id="AutoNumber1">
  <tr>
    <td width="140" valign="top"><font face="Arial">Ingresos Mensuales</font>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >" & mesini(i) & " and FechaEstado <" & mesfin(i) & " And (Estado = 9 or Estado = 19 or Estado = 20 or Estado = 24 or Estado = 25 or Estado = 6)",ObConn
numer=0
DO WHILE NOT ObRs.Eof

SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext

LOOP
ObRs.Close
ObConn.Close
%>
    <p>
<font face="Arial">Enero : $<%Response.Write(FormatNumber(numer))%> <br>

    </font> </td>
    <td width="140" valign="top"><font face="Arial">Entregado y cobrado</font><%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >" & mesini(i) & " and FechaEstado <" & mesfin(i) & " And (Estado = 9)",ObConn
numer=0
DO WHILE NOT ObRs.Eof

SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext

LOOP
ObRs.Close
ObConn.Close
%>
    <p>
<font face="Arial">Enero : $<%Response.Write(FormatNumber(numer))%> </font> 
    </td>
    <td width="140" valign="top"><font face="Arial">Vendido</font>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >" & mesini(i) & " and FechaEstado <" & mesfin(i) & " And (Estado = 19)",ObConn
numer=0
DO WHILE NOT ObRs.Eof

SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext

LOOP
ObRs.Close
ObConn.Close
%>
    <p>
<font face="Arial">Enero : $<%Response.Write(FormatNumber(numer))%> </font> 
    </td>
    <td width="140" valign="top">Cobrada
    <%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >" & mesini(i) & " and FechaEstado <" & mesfin(i) & " And (Estado = 6)",ObConn
numer=0
DO WHILE NOT ObRs.Eof

SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext

LOOP
ObRs.Close
ObConn.Close
%>
    <p>
<font face="Arial">Enero : $<%Response.Write(FormatNumber(numer))%> </font> 

    <td width="140" valign="top"><font face="Arial">Pago a Cuenta
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >" & mesini(i) & " and FechaEstado <" & mesfin(i) & " And (Estado = 24)",ObConn
numer=0
DO WHILE NOT ObRs.Eof

SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext

LOOP
ObRs.Close
ObConn.Close
%>
    </font>
    <p>
<font face="Arial">Enero : $<%Response.Write(FormatNumber(numer))%> </font> 

    
    </td>
    <td width="140" valign="top"><font face="Arial">Cobro a Cuenta</font>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >" & mesini(i) & " and FechaEstado <" & mesfin(i) & " And (Estado = 25)",ObConn
numer=0
DO WHILE NOT ObRs.Eof

SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext

LOOP
ObRs.Close
ObConn.Close
%>
    <p>
<font face="Arial">Enero : $<%Response.Write(FormatNumber(numer))%> </font> 


    </td>
    <td width="140" valign="top"><font face="Arial">Pagado</font><%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Ordenes Where FechaEstado >" & mesini(i) & " and FechaEstado <" & mesfin(i) & " And (Estado = 20)",ObConn
numer=0
DO WHILE NOT ObRs.Eof

SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "Select * From OrdenItem Where Orden = " & ObRs("Id"),ObConn
DO WHILE NOT ObRs2.Eof
numer = numer + ObRs2("Cantidad") * ObRs2("PrecioUnitario")
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
ObRs.MoveNext

LOOP
ObRs.Close
ObConn.Close
%>
    <p>
<font face="Arial">Enero : $<%Response.Write(FormatNumber(numer))%> <br>



    </font> </td>
  </tr>
</table>
<p> <br>

</p>
<p> <br>
</p>


<%
Else
Response.Redirect ("administrativo.asp")
End if
%>

</body>

</html>