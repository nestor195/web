<%
If Session("loginokay") = "" then
Response.redirect "login.asp"
end if
%>
<html>
<head>

<meta content="es-ar" http-equiv="Content-Language">
<title>SOLICITUD DE ACCIONES CORRECTIVAS Y PREVENTIVAS</title>

<link rel="stylesheet" href="./tabla/style.css">
<link rel="preload" href="tabla/integrator.js" as="script">
<script src="./tabla/ca-pub-0652501482160045.js.descarga"></script>
<script type="text/javascript" src="tabla/integrator.js"></script>
<link rel="preload" href="tabla/integrator2.js" as="script">
<script type="text/javascript" src="tabla/integrator2.js"></script>

</head>
<body style="">
<p class="Tilulo">GC.P2.F1 - SOLICITUD DE ACCIONES CORRECTIVAS Y PREVENTIVAS - 
REV 02</p>
	<div id="tablewrapper">
		<div id="tableheader">
        	<div class="search">
                <select id="columns" onchange="sorter.search(&#39;query&#39;)">
                <option value="-1">All Columns</option>
                <option value="0">Estado</option>
                <option value="1">No conformidad</option>
                <option value="2">Fecha</option>
                <option value="3">Area</option>
                <option value="4">Solicitante</option>
                </select>
                <input type="text" id="query" onkeyup="sorter.search(&#39;query&#39;)">
            </div>
            <span class="details">
				<div>Records <span id="startrecord">1</span>-<span id="endrecord">10</span> of <span id="totalrecords">49</span></div>
        		<div><a href="javascript:sorter.reset()">reset</a></div>
        	</span>
        </div>
        
        <table cellpadding="0" cellspacing="0" border="0" id="table" class="tinytable">
            <thead>
                <tr>
                    <th class="nosort"><h3>Numero</h3></th>
                    <th class="asc"><h3>Estado</h3></th>
                    <th class="head"><h3>No conformidad</h3></th>
                    <th class="head"><h3>Fecha</h3></th>
                    <th class="head"><h3>Area</h3></th>
                    <th class="head"><h3>Solicitante</h3></th>
                </tr>
            </thead>
            <tbody>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "ACYP"
SQL = "SELECT Planilla.Id, Planilla.Estado, Estados.Estado, Planilla.NoConformidad, Planilla.Fecha, Areas.Area, Responsables.Responsable"
SQL = SQL & " FROM ((Planilla INNER JOIN Estados ON Planilla.Estado = Estados.ID) INNER JOIN Areas ON Planilla.Area = Areas.Id) INNER JOIN Responsables ON Planilla.Solicita = Responsables.Id"

Select Case Request.Querystring ("Campo")

Case "Id"
If IsNumeric(Request.QueryString ("Dato")) Then
SQL = SQL & " Where Planilla.Id = " & Request.QueryString ("Dato") & " Order by Planilla.Id desc"
End If
Case "Estado"
SQL = SQL & " Where Estados.Estado Like '%" & Request.QueryString ("Dato") & "%' Order by Planilla.Id desc"

End Select

ObRs.Open SQL,ObConn
numeroregistro = 0
Do While ObRs.EOF = false
%>
            	<tr class="evenrow" onmouseover="sorter.hover(<%Response.Write numeroregistro%>,1)" onmouseout="sorter.hover(<%Response.Write numeroregistro%>,0)" id="">
                    <td class=""><a href="consulta.asp?AC=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a></td>
                    <td class="evenselected"><%Response.Write ObRs ("Estado")%></td>
                    <td class=""><%Response.Write ObRs ("NoConformidad")%></td>
                    <td class=""><%Response.Write ObRs ("Fecha")%></td>
                    <td class=""><%Response.Write ObRs ("Area")%></td>
                    <td class=""><%Response.Write ObRs ("Responsable")%></td>
                </tr>
<%
numeroregistro = numeroregistro + 1
ObRs.MoveNext
Loop
ObRs.Close
ObConn.Close
%>

</table>
        <div id="tablefooter">
          <div id="tablenav" style="display: block;">
            	<div>
                    <img src="./tabla/first.gif" width="16" height="16" alt="First Page" onclick="sorter.move(-1,true)">
                    <img src="./tabla/previous.gif" width="16" height="16" alt="First Page" onclick="sorter.move(-1)">
                    <img src="./tabla/next.gif" width="16" height="16" alt="First Page" onclick="sorter.move(1)">
                    <img src="./tabla/last.gif" width="16" height="16" alt="Last Page" onclick="sorter.move(1,true)">
                </div>
                <div>
                	<select id="pagedropdown" onchange="sorter.goto(this.value)"><option value="1">1</option><option value="2">2</option><option value="3">3</option><option value="4">4</option><option value="5">5</option></select>
				</div>
                <div>
                	<a href="javascript:sorter.showall()">view all</a>
                </div>
            </div>
			<div id="tablelocation">
            	<div>
                    <select onchange="sorter.size(this.value)">
                    <option value="5">5</option>
                        <option value="10" selected="selected">10</option>
                        <option value="20">20</option>
                        <option value="50">50</option>
                        <option value="100">100</option>
                    </select>
                    <span>Entries Per Page</span>
                </div>
                <div class="page">Page <span id="currentpage">1</span> of <span id="totalpages">5</span></div>
            </div>
        </div>
    </div>

	<script type="text/javascript" src="./tabla/script.js.descarga"></script>
	<script type="text/javascript">
	var sorter = new TINY.table.sorter('sorter','table',{
		headclass:'head',
		ascclass:'asc',
		descclass:'desc',
		evenclass:'evenrow',
		oddclass:'oddrow',
		evenselclass:'evenselected',
		oddselclass:'oddselected',
		paginate:true,
		size:10,
		colddid:'columns',
		currentid:'currentpage',
		totalid:'totalpages',
		startingrecid:'startrecord',
		endingrecid:'endrecord',
		totalrecid:'totalrecords',
		hoverid:'selectedrow',
		pageddid:'pagedropdown',
		navid:'tablenav',
		sortcolumn:1,
		sortdir:1,
		sum:[1],
		avg:[1],
		columns:[{index:1, format:'%', decimals:2},{index:1, format:'$', decimals:0}],
		init:true
	});
  </script>



</body>
</html>