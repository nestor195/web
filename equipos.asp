<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Sin título 1</title>
</head>

<body>

<p>Listado de Equipos Ingresados para Reparación o Configuración</p>
<p><a href="ingresoequipo.asp">Ingreso de Equipo</a></p>
<table style="width: 100%">
	<tr>
		<td id="id">Id</td>
		<td id="cliente">Cliente</td>
		<td id="equipo" style="width: 141px">Equipo</td>
		<td id="serie">Serie</td>
		<td id="fecha" style="width: 71px">Fecha de Ingreso</td>
		<td id="falla" style="width: 135px">Falla Declarada</td>
		<td id="estado">Estado</td>
	</tr>
	<tr>
		<td><input name="Id" style="width: 33px" type="text" onkeypress="filtrar('id')" /></td>
		<td><input name="Text2" style="width: 117px" type="text" onkeypress="filtrar('cliente')" /></td>
		<td style="width: 141px">
		<input name="Text3" style="width: 83px" type="text" onkeypress="filtrar('equipo')" /></td>
		<td><input name="Text4" style="width: 97px" type="text" onkeypress="filtrar('serie')" /></td>
		<td style="width: 71px">
		<input name="Text5" style="width: 101px" type="text" onkeypress="filtrar('fecha')" /></td>
		<td style="width: 135px">
		<input name="Text6" style="width: 239px" type="text" onkeypress="filtrar('falla')" /></td>
		<td><input name="Text7" style="width: 98px" type="text" onkeypress="filtrar('estado')" /></td>
	</tr>
</table>

</body>

</html>
