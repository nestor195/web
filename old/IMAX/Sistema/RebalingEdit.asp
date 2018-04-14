<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString ("Id")
ObRs.Open SQL,ObConn, 3, 3

ObRs ("Modelo") = Request.Form ("Modelo")
ObRs ("Nbridge") = Request.Form ("Nbridge")
ObRs ("GPU") = Request.Form ("GPU")
ObRs ("Notas") = Request.Form ("Notas")
if Request.Form ("Reballing") = "si" then
ObRs ("Reballing") = True
else
ObRs ("Reballing") = False
end if
ObRs.Update

ObRs.Close
ObConn.Close
response.redirect "ConsultaDeOrden.asp?Id=" & Request.QueryString ("Id")
%>
