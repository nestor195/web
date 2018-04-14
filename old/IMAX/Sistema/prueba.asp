<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Language" content="es-ar" />
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
<title>prueba</title>
</head>

<body>

<p>prueba</p>
<%
fechaactual = date
diasprogramados = 10
fechaprogramada = date + diasprogramados

diaactual = DatePart("w", fechaactual)
diaprogramado = DatePart("w", fechaprogramada)

response.write fechaactual
response.write "<br>"
response.write fechaprogramada
response.write "<br>"
response.write diaactual
response.write "<br>"
response.write diaprogramado
response.write "<br>"
response.write "<br>"

'calculo de cuantos dias habiles faltan
diashabiles = 0
i = date
do while i <= fechaprogramada - 1
	i = i+1
	j = DatePart("w", i)
	if j <> 1 and j <> 7 then
		diashabiles = diashabiles + 1
	end if
loop  
response.write "dias habiles : " & diashabiles
response.write "<br>"



'calculo de fecha programada
response.write "dias programados : " & diasprogramados
response.write "<br>"

fechaprogramada = date
for i = 1 to diasprogramados
	fechaprogramada = fechaprogramada + 1
	j = DatePart("w", fechaprogramada)
	select case j
		case 1
			fechaprogramada = fechaprogramada + 1
		case 7
			fechaprogramada = fechaprogramada + 2
	end select
	response.write "fecha programada : " & fechaprogramada
	response.write "<br>"
next
response.write "****fecha programada : " & fechaprogramada
response.write "<br>"
%>


</body>

</html>
