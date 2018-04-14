<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
</head>

<body text="#FFFFFF" bgcolor="#000000">
<% 
Function FormatMediumDate(DateValue)
Dim strYYYY
Dim strMM
Dim strDD

strYYYY = CStr(DatePart("yyyy", DateValue))

strMM = CStr(DatePart("m", DateValue))
If Len(strMM) = 1 Then strMM = "0" & strMM

strDD = CStr(DatePart("d", DateValue))
If Len(strDD) = 1 Then strDD = "0" & strDD

FormatMediumDate = strMM & "/" & strDD & "/" & strYYYY

End Function 
%> 

<%
IF Session("Session") = "" THEN
Response.Redirect ("inicio.asp")
End If
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios where Id = "& Session("Session")
ObRs.Open  SQL,ObConn
IdCliente = ObRs("Cliente")
Nick = ObRs("Nick")
ObRs.Close
ObConn.Close
%>

<p><%Response.Write Nick%></p>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolor="#0000FF">
  <tr>
    <td width="8%" bgcolor="#FF0000"><font face="Arial Black" size="2"><b>Orden</b></font></td>
    <td width="23%" bgcolor="#FF0000"><font face="Arial Black" size="2"><b>
    Equipo</b></font></td>
    <td width="16%" bgcolor="#FF0000"><font face="Arial Black" size="2"><b>Nº 
    serie</b></font></td>
    <td width="13%" bgcolor="#FF0000">&nbsp;</td>
    <td width="16%" bgcolor="#FF0000"><font face="Arial Black" size="2"><b>Fecha 
    Ingreso</b></font></td>
    <td width="9%" bgcolor="#FF0000"><font face="Arial Black" size="2"><b>
    Estado</b></font></td>
    <td width="14%" bgcolor="#FF0000"><font face="Arial Black" size="2"><b>Fecha 
    de Estado</b></font></td>
  </tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"

If Request.QueryString("Serie") <> "" then
	Serie = Request.QueryString("Serie")
	SQL = "Select * From Ordenes where Cliente = " & IdCliente & " and (Estado = 1 or Estado = 2 or Estado = 3 or Estado = 4 or Estado = 5 or Estado = 7 or Estado = 8 or Estado = 9 or Estado = 13 or Estado = 14 or Estado = 15 or Estado = 16 or Estado = 17 or Estado = 18) and Serie like '%" & Serie & "%'" & " Order by Id"
Else	
	If Request.QueryString("Estado") <> "" then
		Estado = Request.QueryString("Estado")
		SQL = "Select * From Ordenes where Cliente = " & IdCliente & " and (Estado = 1 or Estado = 2 or Estado = 3 or Estado = 4 or Estado = 5 or Estado = 7 or Estado = 8 or Estado = 9 or Estado = 13 or Estado = 14 or Estado = 15 or Estado = 16 or Estado = 17 or Estado = 18) and Estado = " & Estado & " Order by Id"
	Else
		If Request.QueryString("Referencia") <> "" then
			Referencia = Request.QueryString("Referencia")
			SQL = "Select * From Ordenes where Cliente = " & IdCliente & " and (Estado = 1 or Estado = 2 or Estado = 3 or Estado = 4 or Estado = 5 or Estado = 7 or Estado = 8 or Estado = 9 or Estado = 13 or Estado = 14 or Estado = 15 or Estado = 16 or Estado = 17 or Estado = 18) and Referencia like '%" & Referencia & "%'" & " Order by Id"
		Else
			If Request.QueryString("Orden") <> "" then
				Orden = Request.QueryString("Orden")
				SQL = "Select * From Ordenes where Cliente = " & IdCliente & " and (Estado = 1 or Estado = 2 or Estado = 3 or Estado = 4 or Estado = 5 or Estado = 7 or Estado = 8 or Estado = 9 or Estado = 13 or Estado = 14 or Estado = 15 or Estado = 16 or Estado = 17 or Estado = 18) and Id = " & Orden & " Order by Id"
			Else
				SQL = "Select * From Ordenes where Cliente = " & IdCliente & " and (Estado = 1 or Estado = 2 or Estado = 3 or Estado = 4 or Estado = 5 or Estado = 7 or Estado = 8 or Estado = 9 or Estado = 13 or Estado = 14 or Estado = 15 or Estado = 16 or Estado = 17 or Estado = 18)" & " Order by Id"
			End If
		End If
	End If
End If

ObRs.Open  SQL,ObConn

DO WHILE NOT ObRs.Eof
Orden = ObRs("Id")
IdEquipo = ObRs("Equipo")
IdEstado = ObRs("Estado")
Serie = ObRs("Serie")
Referencia = ObRs("Referencia")
FechaIngreso = ObRs("FechaIngreso")
FechaEstado = ObRs("FechaEstado")

	SET ObConn2 = Server.CreateObject ("ADODB.Connection")
	SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
	ObConn2.Open "Sistema"
	SQL = "Select * From Equipos where Id = " & IdEquipo
	ObRs2.Open SQL,ObConn2
	
	Modelo = ObRs2("Modelo")
	
	ObRs2.Close
	ObConn2.Close

	SET ObConn2 = Server.CreateObject ("ADODB.Connection")
	SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
	ObConn2.Open "Sistema"
	SQL = "Select * From Estados where Id = " & IdEstado
	ObRs2.Open SQL,ObConn2
	
	Estado = ObRs2("Estado")
	
	ObRs2.Close
	ObConn2.Close
%>
  <tr>
    <td width="8%" bgcolor="#FFFFFF"><b><font color="#000000" face="Arial">
    <a target="principal" href="orden.asp?orden=<%Response.Write Orden%>"><font color="#000000"><%Response.Write Orden%></font></a></font></b>&nbsp;</td>
    <td width="23%" bgcolor="#FFFFFF"><b><font face="Arial" color="#000000">
    <%Response.Write Modelo%></font></b>&nbsp;</td>
    <td width="16%" bgcolor="#FFFFFF"><b><font face="Arial" color="#000000">
    <%Response.Write Serie%></font></b>&nbsp;</td>
    <td width="13%" bgcolor="#FFFFFF"><b><font face="Arial" color="#000000">
    <%Response.Write Referencia%></font></b>&nbsp;</b></td>
    <td width="16%" bgcolor="#FFFFFF"><b><font face="Arial" color="#000000">
    <%Response.Write FechaIngreso%></font></b>&nbsp;</td>
    <td width="9%" bgcolor="#FFFFFF"><b><font face="Arial" color="#000000"> 
    <%Response.Write Estado%></font></b>&nbsp;</td>
    <td width="14%" bgcolor="#FFFFFF"><b><font face="Arial" color="#000000">
    <%Response.Write FechaEstado%></font></b>&nbsp;</td>
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