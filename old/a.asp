<html> 
<head> 
<title>.....</title>
<style>
table {
    font-family: arial, sans-serif;
    border-collapse: collapse;
    width: 100%;
}

td, th {
    border: 1px solid #dddddd;
    text-align: left;
    padding: 8px;
}

tr:nth-child(even) {
    background-color: #dddddd;
}
</style>
</head> 


<body> 
<%

Dim columna(3,1)
columna(0,0) = "Columna 1"
columna(0,1) = "Id"
columna(1,0) = "Columna 2"
columna(1,1) = "fecha_importacion"
columna(2,0) = "Columna 3"
columna(2,1) = "nro_int"
columna(3,0) = "Columna 4"
columna(3,1) = "nombre"
Selector = "SELECT * FROM Worksheet order by nro_int"
Conector = "a"
vinculo = "consulta.asp?Id="
parametro = "id"
tabla columna, selector, conector, vinculo, parametro

%>


<%
sub tabla(a,b,c,d,e)
'a ---> array con el nombre de las columnas y cual campo consulta
'b ---> select
'c ---> conector ODBC
'd ---> vinculo con parametro ej: "consulta.asp?Id="
'e ---> Parametro: es el nombre de un campo del selec
%>
<table border="1">
  <thead>
	<tr>
<%
For i = 0 to uBound(a)
%>
		<td><%response.write a(i,0)%>&nbsp;</td>
<%
Next
%>
	</tr>
  </thead>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open c
ObRs.Open b,ObConn
Do While ObRs.EOF = false
%>
	<tr>
<%
For i = 0 to uBound(a)
If i = 0 Then
%>
		<td><a href="<%response.write d & ObRs(e) %>"><%response.write ObRs(a(i,1))%></a>&nbsp;</td>
<%
Else
%>
		<td><%response.write ObRs(a(i,1))%>&nbsp;</td>
<%
End If
Next
%>
	</tr>
<%
ObRs.MoveNext
Loop
ObRs.Close
%>
</table>
<%
end sub
%>
</body> 
</html>