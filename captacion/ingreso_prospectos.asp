<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<%
' Este programa es para ingresar los clientes que han sido captados por los vendedores
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Response.Expires = 0
titulo="INGRESO DE PROSPECTOS"
hoy = date()
dia = day(hoy)
anio = year(hoy)
mes = month(hoy)
%>
<%
session("estado")= 1
select case session("bien")
	case "AL"
		titulo1= "ALEVINES" 
	case "TV"
		titulo1= "TRUCHA VIVA"
	case "FH"
		titulo1= "FILETE CON HUESO"
	case "FS"
		titulo1= "FILETE SIN HUESO"
	case "MA"
		titulo1= "FILETE MARIPOSA"
	case "AA"
		titulo1= "ASESORIA EN ACUACULTURA"
end select
' para punto de venta
Set Conpuntos = Server.CreateObject("ADODB.Connection")
Conpuntos.Open ("provider=SQLOLEDB.1;Persist Security Info=False;User ID=inform; PASSword=inform;Initial Catalog=csc;Data Source=ANYJUAN-PC\srv_aplicc")
set puntos = CreateObject("ADODB.Recordset")
' estados
Set Conestados = Server.CreateObject("ADODB.Connection")
Conestados.Open("provider=SQLOLEDB.1;Persist Security Info=False;User ID=inform; PASSword=inform;Initial Catalog=csc;Data Source=ANYJUAN-PC\srv_aplicc")
set estados = CreateObject("ADODB.Recordset")
' prefijos telefono 1
Set Conprefijo1 = Server.CreateObject("ADODB.Connection")
Conprefijo1.Open("provider=SQLOLEDB.1;Persist Security Info=False;User ID=inform; PASSword=inform;Initial Catalog=csc;Data Source=ANYJUAN-PC\srv_aplicc")
set prefijo1 = CreateObject("ADODB.Recordset")
' prefijos telefono 2
Set ConPrefijo2 = Server.CreateObject("ADODB.Connection")
Conprefijo2.Open("provider=SQLOLEDB.1;Persist Security Info=False;User ID=inform; PASSword=inform;Initial Catalog=csc;Data Source=ANYJUAN-PC\srv_aplicc")
set prefijo2 = CreateObject("ADODB.Recordset")
%>
<html>
	<head>
		<title>Captacion de clientes desde CN</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
		<script language="JavaScript" type="text/javascript" src="js/calendar.js"></script>
		<script language="JavaScript" type="text/javascript" src="js/calendar-es.js"></script>
		<script language="JavaScript" type="text/javascript" src="js/calendar-setup.js"></script>
		<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
		<link rel="stylesheet" type="text/css" href="css/calendario.css" >
	</head>
<body class="cuerpo">
	<table width="100%" class="gradient" >
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
	<table width="100%" height ="2%"><tr></tr></table>
	<table align="center">
		<tr> 
			<td class="titulo"><%response.write(titulo)%></td>
		</tr>
		<tr><td height="10"></td></tr>
		<tr> 
			<td class="subtitulo"><%response.write(session("titulo"))%></td>
		</tr>
		<tr> 
			<td class="subtitulo"><%response.write(session("titulo1"))%></td>
		</tr>
	</table>
	<table width="100%" height ="2%"><tr></tr></table>
	<form name="formulario" action="grabadatos.asp" method="post" >
		<table width="837" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ACE5EE">
			<tr> 
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">Cedula</td>
				<td width="165" class="presenta2"><%response.write(session("cedula"))%></td>
				<td width="24">&nbsp;</td>
				<td width="12">&nbsp;</td>
				<td width="131" class="presenta1">Fecha de Registro</td>
				<td width="163" class="presenta2"><%response.write(DATE())%></td>
			</tr>
			<tr> 
				<td width="12"  class="presenta1">*</td>
				<td width="131" class="presenta1">Apellidos</td>
				<td width="165"> <font size="1">
					<input name="apellido" type="text" value="<%response.write(apellido)%>"  maxlength="30" required autofocus tabindex ="1">
				</td>
				<td width="24">&nbsp;</td>
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">Nombres</td>
				<td width="163"><font size="1">
                   	<input name="nombre" type="text" value="<%response.write(nombre)%>" maxlength="30" required tabindex ="2">
				</td>
			</tr>
			<tr> 
				<td width="12" class="presenta1">*</td>
				<td width="131"><font color="#990000" size="2"><div align="left"><b>Direccion Domicilio</td>
				<td width="165" class="presenta1">
					<input type="text" name="direccion"  value="<%response.write(direccion)%>"  size="25" maxlength="80" required tabindex ="3" ></td>
				<td width="24">&nbsp;</td>
				<td width="12">&nbsp;</td>
				<td width="131" class="presenta1">Direccion Trabajo</td>
				<td width="163"> <font size="1">
					<input name="trabajo" type="text" value="<%response.write(trabajo)%>" size="25" maxlength="80" tabindex ="4" >
				</td>
			</tr>
			<tr> 
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">E_Mail</td>
				<td width="165" class="presenta2"><input name="correo" type="email" value="<%response.write(correo)%>" 
					pattern="^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$" size="25" maxlength="40" required tabindex ="5"></td>
				<td width="24">&nbsp;</td>
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">Celular</td>
				<td width="163" class="presenta2">
					<input name="celular" type="text" value="<%response.write(celular)%>" pattern="[0-9]{10}" size="10" maxlength="10" required tabindex ="6"></td>
			</tr>
		</table>
		<table width="100%" height ="3%" border="0"><tr></tr></table>
		<table width="788"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#ACE5EE">
			<tr> 
				<td width="13" class="presenta1">*</td>
				<td width="108" class="presenta1">Provincia</td>
				<td width="108" class="presenta1">Telefono 1</td>
				<td width="13" class="presenta1"></td>
				<td width="108" class="presenta1">Provincia</td>
				<td width="109" class="presenta1">Telefono2</td>
				<td width="13" class="presenta1">*</td>
				<td width="129" class="presenta1">Año de Nacimiento</td>
				<td width="13" class="presenta1">*</td>
				<td width="114" class="presenta1">Estado Civil</td>
			</tr>
			<tr> 
				<td width="13" class="presenta1"></td>
				<td width="108" class="presenta2">
					<select size="1" id="codarea1" name="codarea1" tabindex="7" required onBlur="valprefijo1()">
						<option selected>Seleccione</option>
							<%prefijo1_Sql="Select * from prefijos"
								prefijo1.Open prefijo1_Sql, Conprefijo1
							do while not prefijo1.eof
							%>
							<option value="<%=prefijo1("prenutel")%>"><%=prefijo1("dsubi")%></option>
							<%prefijo1.movenext
								loop
								set prefijo1=nothing
								Conprefijo1.close
							set Conprefijo1=nothing%>
					</select>
			  </td>
				<td width="108" class="presenta2">
				  <input name="telefono1" type ="numeric" value="<%response.write(telefono1)%>" pattern="[0-9]{7}" size="7" maxlength="7" required tabindex ="8"></td>
				<td width="13"></td>
			  <td width="108" class="presenta2">
					<select size="1" id="codarea2" name="codarea2" tabindex="9">
						<option selected>Seleccione</option>
							<%prefijo2_Sql="Select * from prefijos"
								prefijo2.Open prefijo2_Sql, Conprefijo2
							do while not prefijo2.eof
							%>
							<option value="<%=prefijo2("prenutel")%>"><%=prefijo2("dsubi")%></option>
							<%prefijo2.movenext
								loop
								set prefijo2=nothing
								Conprefijo2.close
							set Conprefijo2=nothing%>
					</select>
				</td>

				<td width="109" class="presenta2">
				  <input name="telefono2" type=  "numeric" value="<%response.write(telefono2)%>" size="7" maxlength="7" tabindex="10"></td>
				<td width="13"></td>
				<td width="129" class="presenta2">
				  <input name="nacer" type="numeric" value="<%response.write(nacer)%>" pattern="[0-9]{4}" size="4" maxlength="4" tabindex="11" required></td>
				<td width="13" class="presenta1"></td>
				<td width="114" class="presenta2">
					<select size="1" id="ecivil" name="ecivil"  tabindex="12" required onBlur="valestado()">
						<option selected>Seleccione</option>
							<%estados_Sql="Select * from estado WHERE ID_estado>0 order by nom_estado"
							estados.Open estados_Sql, Conestados
							do while not estados.eof
							%>
								<option value="<%=estados("id_estado")%>"><%=estados("nom_estado")%></option>
							<%estados.movenext
								loop
								set estados=nothing
								Conestados.close
							set Conestados=nothing%>
					</select>
				</td>
			</tr>
		</table>              
		<table width="100%" height ="3%"><tr></tr></table>
		<table width="800"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#ACE5EE">
			<tr> 
				<td width="13" class="presenta1">*</td>
				<td width="129" class="presenta1">Fecha Contacto</td>
				<td width="17" class="presenta1">*</td>
				<td width="129" class="presenta1">Fecha de Cita</td>
				<td width="17" class="presenta1">*</td>
				<td width="113" class="presenta1">Punto de Venta</td>
			</tr>
			<tr> 
				<td width="13"></td>
				<td width="129">
					<input type="text" name="feregis" id="feregis" value="<%response.write(feregis)%>" tabindex ="13"/> 
					<img src="imagenes/calendario.png" width="16" height="16" border="0" title="Fecha Contacto" id="lanzador">
					<script type="text/javascript"> 
						Calendar.setup({ 
						inputField:"feregis",     // id del campo de texto 
						ifFormat  :"%d-%m-%Y",     // formato de la fecha que se escriba en el campo de texto 
						button    :"lanzador"     // el id del botón que lanzará el calendario 
						}); 
					</script>
				</td>
				<td width="17"></td>
				<td width="129"><div>
					<input type="text" name="cita" id="cita" value="<%response.write(cita)%>" tabindex ="14"/> 
					<img src="imagenes/calendario.png" width="16" height="16" border="0" title="Fecha Cita" id="lanzador1">
					<script type="text/javascript"> 
						Calendar.setup({ 
						inputField:"cita",     // id del campo de texto 
						ifFormat  :"%d-%m-%Y",     // formato de la fecha que se escriba en el campo de texto 
						button    :"lanzador1"     // el id del botón que lanzará el calendario 
						}); 
					</script></div>
				</td>
				<td width="17"></td>
				<td width="208" class="presenta2">
				<select size="1" id="pto1" name="pto1"  tabindex="15" required onBlur="valpunto()">
					<option selected>Seleccione</option>
						<%puntos_Sql="Select * from PUNTOVENTA WHERE ID_puntoventa>0 order by nom_puntoventa"
						puntos.Open puntos_Sql, Conpuntos
						do while not puntos.eof
						%>
							<option value="<%=puntos("id_puntoventa")%>"><%=puntos("nom_puntoventa")%></option>
						<%puntos.movenext
							loop
							set puntos=nothing
							Conpuntos.close
						set Conpuntos=nothing%>
				</select>
				</td>
		</table>
		<table width="100%" height ="6%"><tr></tr></table>
		<table align="center">
			<tr>
				<td width="33%" align="center">&nbsp;</td>
				<td width="33%" align="center">&nbsp;</td>
				<td colspan="2" align="center"><input type="submit"  name="btnBuscar" value="grabar"></button>
			</tr>
		</table>
	</form>
	<table width="100%"  align="center">
		<tr>
			<td width="26%" align="center">
				<FORM NAME="volver" ACTION="cedula_prospectos.asp">
					<input type="submit" name="" value="Regresar">
				</FORM>
			</td>
			<td width="26%"></td>
		</tr>
	</table>
</body>
</html>
