<%@language=vbscript%>
<!--#include file="fpdf.asp"-->
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaOrdenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL, ObConn

Set pdf=CreateJsObject("FPDF")
pdf.CreatePDF "p","mm","A4"
pdf.SetFont "Arial","b",13
pdf.Open()
pdf.AddPage()

pdf.Image "Orden.jpg", 10, 40, 180, 10, "JPG"
pdf.Image "Imax.jpg", 10, 12, 60, 25, "JPG"

pdf.SetFont "Arial","b",8
pdf.Text 70,18,"Servicio T�cnico Inform�tico"
pdf.SetFont "Arial","b",8
pdf.Text 70,21.5,"Impresoras - PC - Notebooks"
pdf.SetFont "Arial","b",8
pdf.Text 70,25,"Insumos y Accesorios"
pdf.SetFont "Arial","b",8
pdf.Text 70,28.5,"FRAGUEIRO 152 - CENTRO - CORDOBA"
pdf.SetFont "Arial","b",8
pdf.Text 70,32,"TEL. 4271942 - 152444548"
pdf.SetFont "Arial","b",8
pdf.Text 70,35.5,"email: soporte@imaxcba.com.ar"

pdf.SetFont "Arial","b",12
pdf.Text 140,20,"ORDEN DE REPARACION"
pdf.SetFont "Arial","b",12
pdf.Text 150,25,"N� " & ObRs ("Id")
pdf.SetFont "Arial","b",12
pdf.Text 140,30,"FECHA " & ObRs ("FechaIngreso")

pdf.SetFont "Arial","b",12
pdf.Text 10,60,"Cliente: " & ObRs ("Nombre")
pdf.SetFont "Arial","b",12
pdf.Text 100,60,"Telefono: " & ObRs ("Telefono")
pdf.SetFont "Arial","b",12
pdf.Text 10,65,"Domicilio: " & ObRs ("Direccion")
pdf.SetFont "Arial","b",12
pdf.Text 10,70,"email: " & ObRs ("Email")

pdf.SetFont "Arial","b",12
pdf.Text 10,77,"Equipo: " & ObRs ("Tipo")
pdf.SetFont "Arial","b",12
pdf.Text 80,77,"Marca: " & ObRs ("Marca")
pdf.SetFont "Arial","b",12
pdf.Text 10,82,"Modelo: " & ObRs ("Modelo")
pdf.SetFont "Arial","b",12
pdf.Text 70,82,"Numero de Serie: " & ObRs ("Serie")
pdf.SetFont "Arial","b",12
pdf.Text 10,87,"Accesorios: " & ObRs ("Accesorios")

pdf.SetFont "Arial","b",12
pdf.Text 10,94,"Falla descripta por el cliente:"
pdf.SetFont "Arial","b",12
pdf.Text 10,99, ObRs ("ObservacionIngreso")

pdf.SetFont "Arial","b",12
pdf.Text 10,104,"Informe t�cnico. Falla Encontrada:"
pdf.SetFont "Arial","b",12
pdf.Text 10,112,"......................................................................................................................................................."
pdf.SetFont "Arial","b",12
pdf.Text 10,120,"......................................................................................................................................................."
pdf.SetFont "Arial","b",12
pdf.Text 10,128,"Presupuesto $................Informado al cliente el ......./......./......."
pdf.SetFont "Arial","b",12
pdf.Text 10,136,"Repuestos: .................................................................................................................................."

fuente = 1.12
pdf.SetFont "Arial","b",10 * fuente
pdf.Text 55,135 * fuente,"CONDICIONES GENERALES A LEER POR EL CLIENTE"

pdf.SetFont "Arial","",6.15 * fuente
pdf.Text 6,143 * fuente,"1)"
pdf.SetFont "Arial","",6.15 * fuente
pdf.Text 10,143 * fuente,"EMISION DE PRESUPUESTO Y PLAZO DE ACEPTACION O DENEGATORIA: Ser� puesto a disposici�n del cliente personalmente en el local de IMPRECOM.COM por"
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,147 * fuente,"tel�fono, fax o e-mail. El cliente tendr� 48 hs. para aceptar o denegar el presupuesto despu�s de haber sido puesto a su disposici�n personalmente, por tel�fono, fax o e-mail."

pdf.SetFont "Arial","",6.15 * fuente
pdf.Text 6,153 * fuente,"2)"
pdf.SetFont "Arial","",5.9 * fuente
pdf.Text 10,153 * fuente,"TIEMPO DE REPARACION Y DIAGNOSTICO: Por tratarse de un equipo con componentes importados, el tiempo de reparaci�n puede llegar hasta los 60 d�as. Se deja aclarado"
pdf.SetFont "Arial","",5.95 * fuente
pdf.Text 10,157 * fuente,"que IMPRECOMP.COM no se hace responsable por la demora incurrida en tal sentido porque depende de la llegada de esos componentes. El diagn�stico se informar� dentro"
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,161 * fuente,"de los siete d�as h�biles desde la fecha de ingreso del equipo a partir de los cuales se dar� fecha de retiro del mismo, se acepte o no el presupuesto."

pdf.SetFont "Arial","",6.15 * fuente
pdf.Text 6,167 * fuente,"3)"
pdf.SetFont "Arial","",5.95 * fuente
pdf.Text 10,167 * fuente,"RETIRO DE EQUIPOS: El cliente deber� retirar el equipo dentro de los cinco d�as h�biles de notificado de la puesta a su disposici�n del mismo, debiendo presentar esta orden"
pdf.SetFont "Arial","",5.93 * fuente
pdf.Text 10,171 * fuente,"de reparaci�n en original. A partir del sexto d�a IMPRECOMP.COM podr� cobrar al cliente la suma de un d�lar estadounidenses billete (U$S 1) por d�a de estadia, en calidad de"
pdf.SetFont "Arial","",5.88 * fuente
pdf.Text 10,175 * fuente,"dep�sito. Asimismo y en dicho caso, IMPRECOMP.COM no ser� responsable por la p�rdida o deterioro que sufra el equipo despu�s de ser notificado. Pasados los 60 d�as desde"
pdf.SetFont "Arial","",5.89 * fuente
pdf.Text 10,179 * fuente,"la fecha de la orden de reparaci�n (se haya aceptado o no e presupuesto) sin que se retire el equipo, se entender� que el cliente renunci� a la propiedad y/o posesi�n de equipo"
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,183 * fuente,"transfiriendole todos sus derechos (articulos 2525, 2526 y 3939 del C�digo Civil) a IMPRECOMP.COM para que disponga del mismo."

pdf.SetFont "Arial","",6.15 * fuente
pdf.Text 6,189 * fuente,"4)"
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,189 * fuente,"Es obligaci�n del cliente tener un Back Up con su informaci�n porque IMPRECOMP.COM no se hace responsable por la p�rdida de softwqare o informaci�n. Tampoco es res-"
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,193 * fuente,"ponsable del origen, procedencia o destino del equipo (reparado o no)."

pdf.SetFont "Arial","",6.15 * fuente
pdf.Text 6,199 * fuente,"5)"
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,199 * fuente,"No ser� responsable IMPRECOMP.COM por las fallas o visios ocultos no declarados por el cliente (conocidos o no por �l) en la orden de reparaci�n."
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,205 * fuente,""
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,207 * fuente,""

pdf.SetFont "Arial","",6.15 * fuente
pdf.Text 6,205 * fuente,"6)"
pdf.SetFont "Arial","",5.85 * fuente
pdf.Text 10,205 * fuente,"GARANTIA: La reparaci�n tiene una garant�a de sesenta d�as corridos a partir de la fecha de retiro del equipo siempre y cuando haya sido retirado dentro de los cinco d�as h�biles"
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,210 * fuente,"de notificado conforme a lo expresado en el punto 3. La garant�a cubrir� solamente la mano de obra y no los repuestos o partes reemplazadas salvo las que otorque el provee-"
pdf.SetFont "Arial","",6 * fuente
pdf.Text 10,215 * fuente,"dor de dichos repuestos o partes."

pdf.SetFont "Arial","b",10
pdf.Text 20,250 * fuente,".............................................................."
pdf.SetFont "Arial","",10
pdf.Text 20,255 * fuente,"      EN CONFORMIDAD CLIENTE"

pdf.SetFont "Arial","b",10
pdf.Text 120,250 * fuente,".............................................................."
pdf.SetFont "Arial","",10
pdf.Text 120,255 * fuente,"               IMPRECOMP.COM"

pdf.SetPath("fpdf/")
pdf.Close()
nombre = "Orden" & ObRs("Id") & ".pdf"
pdf.Output()
'pdf.Output server.mappath(nombre),"T"

ObRs.Close
ObConn.Close
%>