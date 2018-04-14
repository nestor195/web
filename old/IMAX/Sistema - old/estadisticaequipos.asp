<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>
<body>
<%
contadore = 0
contadorh = 0
contadorl = 0
contadorc = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaOrdenes where Estado <> 'Anulado' and (Tipo = 'Impresora InkJet' Or Tipo = 'Multifunción')"
ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
select case ObRs ("Marca")
Case "HP"
contadorh = contadorh + 1
Case "Epson"
contadore = contadore + 1
Case "Lexmark"
contadorl = contadorl + 1
Case "Canon"
contadorc = contadorc + 1
end select
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<p>Epson: <%response.write contadore%></p>
<p><%response.write FormatNumber(contadore * 100 / (contadore + contadorh + contadorl + contadorc))%>%</p>
<p>HP: <%response.write contadorh%></p>
<p><%response.write FormatNumber(contadorh * 100 / (contadore + contadorh + contadorl + contadorc))%>%</p>
<p>Lexmark: <%response.write contadorl%></p>
<p><%response.write FormatNumber(contadorl * 100 / (contadore + contadorh + contadorl + contadorc))%>%</p>
<p>Canon: <%response.write contadorc%></p>
<p><%response.write FormatNumber(contadorc * 100 / (contadore + contadorh + contadorl + contadorc))%>%</p>

</body>
</html>