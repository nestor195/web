<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString ("Id")
ObRs.Open SQL,ObConn, 3, 3

ObRs ("Serie") = Request.Form ("Serie")
ObRs.Update

ObRs.Close
ObConn.Close
response.redirect "ConsultaDeOrden.asp?Id=" & Request.QueryString ("Id")
%>
