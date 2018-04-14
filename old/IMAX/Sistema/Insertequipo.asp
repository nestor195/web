<% ' Insert.asp %>
<!--#include file="Loader.asp"-->
<%
  Response.Buffer = True

  ' load object
  Dim load
    Set load = new Loader
    
    ' calling initialize method
    load.initialize
    
  ' File binary data
  Dim fileData
    fileData = load.getFileData("file")
  ' File name
  Dim fileName
    fileName = LCase(load.getFileName("file"))
  ' File path
  Dim filePath
    filePath = load.getFilePath("file")
  ' File path complete
  Dim filePathComplete
    filePathComplete = load.getFilePathComplete("file")
  ' File size
  Dim fileSize
    fileSize = load.getFileSize("file")
  ' File size translated
  Dim fileSizeTranslated
    fileSizeTranslated = load.getFileSizeTranslated("file")
  ' Content Type
  Dim contentType
    contentType = load.getContentType("file")
  ' No. of Form elements
  Dim countElements
    countElements = load.Count
  ' Value of text input field "fname"
  Dim fnameInput
    fnameInput = load.getValue("fname")
  ' Value of text input field "lname"
  Dim lnameInput
    Equipo = load.getValue("Equipo")
  ' Value of text input field "profession"
  Dim profession
    profession = load.getValue("profession")  
    
  ' destroying load object
  Set load = Nothing
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
    <b>Inserting Binary Data into Database</b><br>
    <a href="show.asp">To see inserted data click here</a>
  </p>
  
  <table width="700" border="1" align="center">
  <tr>
    <td>File Name</td><td><%= fileName %>&nbsp;</td>
  </tr><tr>
    <td>File Path</td><td><%= filePath %>&nbsp;</td>
  </tr><tr>
    <td>File Path Complete</td><td><%= filePathComplete %>&nbsp;</td>
  </tr><tr>
    <td>File Size</td><td><%= fileSize %>&nbsp;</td>
  </tr><tr>
    <td>File Size Translated</td><td><%= fileSizeTranslated %>&nbsp;</td>
  </tr><tr>
    <td>Content Type</td><td><%= contentType %>&nbsp;</td>
  </tr><tr>
    <td>No. of Form Elements</td><td><%= countElements %>&nbsp;</td>
  </tr><tr>
    <td>First Name</td><td><%= fnameInput %>&nbsp;</td>
  </tr><tr>
    <td>Last Name</td><td><%= lnameInput %>&nbsp;</td>
  </tr>
  <tr>
    <td>Profession</td><td><%= profession %>&nbsp;</td>
  </tr>
  </table><br><br>
  
  <p style="padding-left:220;">
  <%= fileName %> data received ...<br>
  <%
    ' Checking to make sure if file was uploaded
    If fileSize > 0 Then
    
      ' Connection string
      Dim connStr

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET Rs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * from Equipos Where Id = " & Equipo
        Rs.Open SQL , ObConn, 3, 3
        ' Adding data
          rs("Imagen").AppendChunk fileData
          rs("ContentType") = contentType
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
      Response.Write "<font color=""green"">File was successfully uploaded..."
      Response.Write "</font>"
    Else
      Response.Write "<font color=""brown"">No file was selected for uploading"
      Response.Write "...</font>"
    End If
      
      
    If Err.number <> 0 Then
      Response.Write "<br><font color=""red"">Something went wrong..."
      Response.Write "</font>"
    End If
  %>
  </p>
  
  <br>
  <table border="0" align="center">
  <tr>
  <form method="POST" enctype="multipart/form-data" action="Insertequipo.asp">
  <td>First Name :</td><td>
    <input type="text" name="fname" size="40" ></td>
  </tr>
  <td>Last Name :</td><td>
    <input type="text" name="lname" size="40" ></td>
  </tr>
  <td>Profession :</td><td>
    <input type="text" name="profession" size="40" ></td>
  </tr>
  <td>File :</td><td>
    <input type="file" name="file" size="40"></td>
  </tr>
  <td> </td><td>
    <input type="submit" value="Submit"></td>
  </tr>
  </form>
  </tr>
  </table>

</body>
</html>