<%
   ' -- show.asp --
   ' Generates a list of uploaded files
   
   Response.Buffer = True
   
   ' Connection String
   Dim connStr
      connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("FileDB.mdb")
%>
<html>
<head>
   <title>Inserts Images into Database</title>
   <style>
      body, input, td { font-family:verdana,arial; font-size:10pt; }
   </style>
</head>
<body>
   <p align="center">
      <b>Showing Binary Data from the Database</b><br>
      <a href="ImagenMarca.asp">To insert data click here</a>
   </p>
   
   <table width="700" border="1" align="center">
<%
   ' Recordset Object
   Dim rs
      Set rs = Server.CreateObject("ADODB.Recordset")
      
      ' opening connection
      rs.Open "select [ID],[File Name],[File Size],[Content Type],[First Name]," & _
         "[Last Name],[Profession] from Files order by [ID] desc", connStr, 3, 4

      If Not rs.EOF Then
         Response.Write "<tr><td colspan=""7"" align=""center""><i>"
         Response.Write "No. of records : " & rs.RecordCount
         Response.Write ", Table : Files</i><br>"
         Response.Write "</td></tr>"
   
         While Not rs.EOF
            Response.Write "<tr><td>"
            Response.Write rs("ID") & "</td><td>"
            Response.Write "<a href=""file.asp?ID=" & rs("ID") & """>"
            Response.Write rs("File Name") & "</a></td><td>"
            Response.Write rs("File Size") & "</td><td>"
            Response.Write rs("Content Type") & "</td><td>"
            Response.Write rs("First Name") & "</td><td>"
            Response.Write rs("Last Name") & "</td><td>"
            Response.Write rs("Profession")
            Response.Write "</td></tr>"
            rs.MoveNext
         Wend
      Else
         Response.Write "No Record Found"
      End If
      
      rs.Close
      Set rs = Nothing
%>
   </table>
</body>
</html>