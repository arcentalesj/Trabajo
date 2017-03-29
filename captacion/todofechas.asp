<%@LANGUAGE="VBSCRIPT"%>
<%
If (Request.QueryString("asigna") <> "") Then 
	session("reporte")=Request.QueryString("asigna")
else  
	session("reporte")="Reporte de Cotizadores a Excel"
End If
%>
<%
hoy = date
dia = day(hoy)
anio = year(hoy)
mes = month(hoy)
'hoy = mes&"/"&dia&"/"&anio
hoy = dia&"/"&mes&"/"&anio
fecfin=hoy
if mes=1 then
	mes1=12
	anio1=anio-1
else 
 	mes1 = mes-1
	anio1 = anio
end if
'hoy1 = (mes1)&"/"&dia&"/"&anio1
hoy1 = (dia)&"/"&mes1&"/"&anio1

'hoy1 = dia&"/"&mes1&"/"&anio1
fecini=hoy1
session("voy") = 0
%>
<%
select case session("reporte")
	case "Reporte ingresados a Excel"
		titulo = "REPORTE DE PROSPECTOS INGRESADOS A EXCEL"
	case "Citas Excel"
		titulo = "REPORTE DE CITAS A EXCEL"
end select
%>
<html>
	<head>
		<title>Registro de Fechas</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
		<script language="JavaScript" src="js/calendar.js" type="text/javascript"></script>
		<script language="JavaScript" src="js/calendar-es.js" type="text/javascript"></script>
		<script language="JavaScript" src="js/calendar-setup.js" type="text/javascript"></script>
		<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
		<link rel="stylesheet" type="text/css" href="css/calendario.css" >
	</head>
<body class="cuerpo">
	<table width="100%" class="gradient">
		<tr> 
			<td width="55%" rowspan="7" class="Estilo4" ><div align="Center"><%response.write(session("Sistema"))%></div></td>
			<td width="15%" height="50%">&nbsp;</td>
			<td width="20%" rowspan="5" class="gradient"><img name="texto" src="imagenes/logo.gif" width="249" height="153" border="0" alt=""></td>
		</tr>
		<tr>
			<td height="50%"></td>
		</tr>
		<tr>
			<td class="fecha"><script type="text/javascript">fecha();</script></td>
		</tr>
		<tr>
			<td height="100%">&nbsp;</td>
		</tr>
	</table>
	<table cellspacing="0" class="Estilo2">
		<tr></tr>
	</table>
	<p>&nbsp;</p>
	<table align="center">
		<tr> 
			<td class="titulo"><%response.write(titulo)%><%response.write("  ")%><%response.write(titulo1)%></div>
            </td>
		</tr>
	</table>
	<table width="100%" height ="25" border="0">
		<tr><td>&nbsp;</td></tr>
	</table>
	<%select case session("reporte")
		case "Reporte ingresados a Excel","Citas Excel" %>
		  <form name="formulario" method=post onSubmit='return veriFechas();' >
				<input type=hidden name="pagina" value="resultado"/>
				<table width="364" border="0" align="center">
					<tr> 
						<td width="125" class="presenta1">Fecha inicial</td>
						<td width="209" class="presenta2">
							<input type="text" name="fecini" id="fecini" value=<%response.write(fecini)%> tabindex="13"/> 
							<img src="imagenes/calendario.png" width="16" height="16" border="0" title="Fecha Inicial" id="lanzador">
						  <script type="text/javascript"> 
								Calendar.setup({ 
								inputField:"fecini",     // id del campo de texto 
								ifFormat  :"%d/%m/%Y",     // formato de la fecha que se escriba en el campo de texto 
								button    :"lanzador"     // el id del botón que lanzará el calendario 
							}); 
							</script></td>
					</tr>
					<tr> 
						<td width="125" class="presenta1">Fecha Final</td>
						<td width="209" class="presenta2">
							<input type="text" name="fecfin" id="fecfin" value=<%response.write(fecfin)%> tabindex="13"/> 
							<img src="imagenes/calendario.png" width="16" height="16" border="0" title="Fecha Final" id="lanzador1">
						  <script type="text/javascript"> 
								Calendar.setup({ 
								inputField:"fecfin",     // id del campo de texto 
								ifFormat  :"%d/%m/%Y",     // formato de la fecha que se escriba en el campo de texto 
								button    :"lanzador1"     // el id del botón que lanzará el calendario 
							}); 
							</script></td>
					</tr>
					<td colspan="7" align="center">&nbsp;</td>
					<tr align="center"> 
					<td colspan="7" align="center">
					<input type="button" class="boton" name="btnBuscar" value="Realizar Consulta" onclick='javascript:veriFechas();'/></td>
					</tr>
			</table>
		</form>
	<%end select%>
	<table width="100%" height ="25" border="0">
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	<p align="center"><font color="#990000" size="3"><b>EL FORMATO DE FECHAS ES MM/DD/YYYY</b></font></p>
	<p align="center">
	<table width="100%"  align="center">
		<tr>
			<td width="14%">				</td>
			<td width="53%">
				<%select case session("reporte")
						case "Reporte ingresados a Excel"%>
							<td width="53%">
								<FORM ACTION="Clientes_ingresados.asp" NAME="volver">
									<input type="submit" name="" value="Regresar">
								</FORM>
							</td>
						<%case "Citas Excel" %>
								<FORM ACTION="menu.asp" NAME="volver">
									<input type="submit" name="" value="Regresar">
								</FORM>
					<%end select%>
			</td><td width="33%"></td>
		</tr>
	</table>
</body>
</html>
