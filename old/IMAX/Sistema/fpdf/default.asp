<%@language=vbscript%>
<!--#include file="fpdf.asp"-->
<%
 
Set pdf=CreateJsObject("FPDF")
pdf.CreatePDF "p","mm","A4"
pdf.SetPath("fpdf/")
pdf.SetFont "Arial","",16
pdf.Open()
pdf.AddPage()
pdf.Image "lambdaWeb.jpg", 0, 0, 210, 297, "JPG","http://www.lambdasi.com.ar"
pdf.Cell 40,10,"Esto es un PDF creado desde ASP."
pdf.Close()
pdf.Output()
%> 
