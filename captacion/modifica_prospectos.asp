<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<%
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Response.Expires = 0
titulo="MODIFICACION DATOS DE CLIENTES"
hoy = date()
dia = day(hoy)
anio = year(hoy)
mes = month(hoy)
session("estado")= 2
%>
<%
If (Request.QueryString("cedula") <> "") Then 
  session("cedula") = Request.QueryString("cedula")
  session("bien") = Request.QueryString("tipobien")
End If
strSQL= "SELECT * FROM captados WHERE cedula = '" & session("cedula") &"'"
%>
<%
Dim captados
Dim captados_numRows
Set captados = Server.CreateObject("ADODB.Recordset")
captados.ActiveConnection = MM_usuarios_STRING
captados.Source = strSQL
captados.CursorType = 0
captados.CursorLocation = 2
captados.LockType = 1
captados.Open()
if captados.eof and captados.bof then
	' No es cliente de nosotros
	response.redirect("cedula_prospectos.asp?mensajeerror=" &"cliente_errado")
else
	ecivil =trim(captados.Fields.Item("idestado").value)
	pto1 =trim(captados.Fields.Item("ptoventa").value)
	arti1=captados.Fields.Item("idproducto").value
	codarea1=captados.Fields.Item("codarea1").value
	codarea2=captados.Fields.Item("codarea2").value
	apellido=captados.Fields.Item("apellido").value
	nombre=captados.Fields.Item("nombre").value

end if
%>
<%
Dim estados
Dim estados_numRows
Set estados = Server.CreateObject("ADODB.Recordset")
estados.ActiveConnection = MM_usuarios_STRING
estados.Source = "SELECT * FROM estado order by 2"
estados.CursorType = 0
estados.CursorLocation = 2
estados.LockType = 1
estados.Open()
estados_numRows = 0
%>
<%
Dim ptoventa
Dim ptoventa_numRows
Set ptoventa = Server.CreateObject("ADODB.Recordset")
ptoventa.ActiveConnection = MM_usuarios_STRING
ptoventa.Source = "SELECT * FROM puntoventa order by 2"
ptoventa.CursorType = 0
ptoventa.CursorLocation = 2
ptoventa.LockType = 1
ptoventa.Open()
ptoventa_numRows = 0
%>
<%
Dim prefidos
Dim prefidos_numRows
Set prefidos = Server.CreateObject("ADODB.Recordset")
prefidos.ActiveConnection = MM_usuarios_STRING
prefidos.CursorType = 0
prefidos.CursorLocation = 2
prefidos.LockType = 1
prefidos.Source = "Select * from prefijos"
prefidos.Open()
prefidos_numRows = 0
%>
<%
Dim bdproducto
Dim bdproducto_numRows
Set bdproducto = Server.CreateObject("ADODB.Recordset")
bdproducto.ActiveConnection = MM_usuarios_STRING
bdproducto.CursorType = 0
bdproducto.CursorLocation = 2
bdproducto.LockType = 1
bdproducto.Source = "select * from articulos order by dsmod"
bdproducto.Open()
bdproducto_numRows = 0
%>
<html>
	<head>
		<title>Modificacion de datos desde CN</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
		<script language="JavaScript" src="js/calendar.js" type="text/javascript"></script>
		<script language="JavaScript" src="js/calendar-es.js" type="text/javascript"></script>
		<script language="JavaScript" src="js/calendar-setup.js" type="text/javascript"></script>
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
	<table cellspacing="0" class="Estilo2"><tr></tr></table>
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
				<td width="163" class="presenta2"><%response.write(captados("ingresado"))%></td>
			</tr>
			<tr> 
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">Apellidos</div></td>
				<td width="165" class="presenta2">
				<%if session("cargo")="71" then %>
					<input name="apellido" type="text" value="<%response.write(captados("apellido"))%>" size="25" maxlength="30" tabindex="1" readonly="">
				<%else %>
					<input name="apellido" type="text" value="<%response.write(captados("apellido"))%>" size="25" maxlength="30" tabindex="1" >
				<%end if%>
				</td>
				<td width="24">&nbsp;</td>
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">Nombres</td>
				<td width="163" class="presenta2">
				<%if session("cargo")="71" then %>
					<input name="nombre" type="text" value="<%response.write(captados("nombre"))%>" size="25" maxlength="30" tabindex="2" readonly="">
				<%else %>
					<input name="nombre" type="text" value="<%response.write(captados("nombre"))%>" size="25" maxlength="30" tabindex="2">
				<%end if%>
				</td>
			</tr>
			<tr> 
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">Direccion Domicilio</td>
				<td width="165" class="presenta2">
					<input name="direccion" type="text" value="<%response.write(captados("direccion"))%>" size="25" maxlength="80" required tabindex ="3">
				</td>
				<td width="24">&nbsp;</td>
				<td width="12">&nbsp;</td>
				<td width="131" class="presenta1">Direccion Trabajo</b></td>
				<td width="163" class="presenta2">
					<input name="trabajo" type="text" value="<%response.write(captados("trabajo"))%>" size="25" maxlength="80" tabindex ="4">
				</td>
			</tr>
			<tr> 
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">E_Mail</td>
				<td width="165" class="presenta2">
					<input name="correo" type="text" value="<%response.write(captados("correo"))%>" size="25" maxlength="40" tabindex ="5"></td>
				<td width="24">&nbsp;</td>
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">Celular</b></td>
				<td width="163" class="presenta2">
					<input name="celular" type="text" value="<%response.write(captados("celular"))%>" size="25" maxlength="40" tabindex ="6"></td>
			</tr>
		</table>
		<table width="100%" height ="3%"><tr></tr></table>
		<table width="840"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#ACE5EE">
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
				<td width="108" class="presenta1">
					<select size="1" id="select" name="codarea1" class="PPRFiel" tabindex="7">
					<%While (NOT prefidos.EOF)%>
						<option value="<%=(prefidos.Fields.Item("prenutel").Value)%>"
						<%If (Not isNull(codarea1)) Then If (trim(prefidos.Fields.Item("prenutel").Value) = trim(codarea1)) Then Response.Write("SELECTED") :
							Response.Write("")%> >
						<%=(prefidos.Fields.Item("dsubi").Value)%></option>
						<%prefidos.MoveNext()
					Wend
					If (prefidos.CursorType > 0) Then
						prefidos.MoveFirst
					Else
						prefidos.Requery
					End If%>
					</select>
				</td>
				<td width="108" class="presenta1">
				<input name="telefono1" type ="numeric" value="<%response.write(captados("telefono"))%>" size="7" maxlength="7" tabindex ="8"></td>
				<td width="13"></td>
				<td width="108" class="presenta1">
					<select size="1" id="codarea2" name="codarea2" tabindex="9">
					<%While (NOT prefidos.EOF)%>
						<option value="<%=(prefidos.Fields.Item("prenutel").Value)%>"
						<%If (Not isNull(codarea2)) Then If (trim(prefidos.Fields.Item("prenutel").Value) = trim(codarea2)) Then Response.Write("SELECTED") :
						Response.Write("")%> >
						<%=(prefidos.Fields.Item("dsubi").Value)%></option>
						<%prefidos.MoveNext()
					Wend
					If (prefidos.CursorType > 0) Then
						prefidos.MoveFirst
					Else
						prefidos.Requery
					End If%>
				</select></td>
				<td width="109" class="presenta1">
				<input name="telefono2" type=  "numeric" value="<%response.write(captados("fono2"))%>" size="7" maxlength="7" tabindex="10"></td>
				<td width="13"></td>
				<td width="129" class="presenta2">
				 <input name="nacer" type="numeric" value="<%response.write(captados("nacer"))%>" pattern="[0-9]{4}" size="4" maxlength="4" tabindex="11" required></td>
				<td width="13" class="presenta1"></td>
				<td width="114" class="presenta1">
					<select name="ecivil" class="PPRFiel" id="select" tabindex="12">
					<%While (NOT estados.EOF)%>
						<option value="<%=(estados.Fields.Item("id_estado").Value)%>"
						<%If (Not isNull(ecivil)) Then If (trim(estados.Fields.Item("id_estado").Value) = trim(ecivil)) Then Response.Write("SELECTED") :
						Response.Write("")%> >
						<%=(estados.Fields.Item("nom_estado").Value)%></option>
						<%estados.MoveNext()
					Wend
					If (estados.CursorType > 0) Then
						estados.MoveFirst
					Else
						estados.Requery
					End If%>
				</select></td>
			</tr>
		</table>  
		<table width="100%" height ="3%"><tr></tr></table>
		<table width="1013"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#ACE5EE">
			<tr> 
				<td width="11" class="presenta1">*</td>
				<td width="231" class="presenta1">Fecha Contacto</td>
				<td width="12" class="presenta1">*</td>
				<td width="235" class="presenta1">Fecha de Cita</td>
				<td width="12" class="presenta1">*</td>
				<td width="238" class="presenta1">Punto de Venta</td>
				<td width="11" class="presenta1">*</td>
				<td width="247" class="presenta1">Producto</td>
			</tr>
			<tr> 
				<td width="11"></td>
				<td width="231"><div>
					<input type="text" name="feregis" id="feregis" value="<%response.write(captados("contactado"))%>" tabindex="13"/> 
					<img src="imagenes/calendario.png" width="16" height="16" border="0" title="Fecha Contacto" id="lanzador">
					<script type="text/javascript"> 
						Calendar.setup({ 
						inputField:"feregis",     // id del campo de texto 
						ifFormat  :"%d-%m-%Y",     // formato de la fecha que se escriba en el campo de texto 
						button    :"lanzador"     // el id del botón que lanzará el calendario 
						}); 
					</script></div>
				</td>
				<td width="12"></td>
				<td width="235"><div>
					<input type="text" name="cita" id="cita" value="<%response.write(captados("cita"))%>" tabindex="14"/> 
					<img src="imagenes/calendario.png" width="16" height="16" border="0" title="Fecha Cita" id="lanzador1">
					<script type="text/javascript"> 
						Calendar.setup({ 
						inputField:"cita",     // id del campo de texto 
						ifFormat  :"%d-%m-%Y",     // formato de la fecha que se escriba en el campo de texto 
						button    :"lanzador1"     // el id del botón que lanzará el calendario 
						}); 
					</script></div>
				</td>
				<td width="12"></td>
				<td width="238" class="presenta2">
					<select name="pto1" class="PPRFiel" id="select" tabindex="15">
						<%While (NOT ptoventa.EOF)%>
							<option value="<%=(ptoventa.Fields.Item("ID_puntoventa").Value)%>"
							<%If (Not isNull(pto1)) Then If (trim(ptoventa.Fields.Item("ID_puntoventa").Value) = trim(pto1)) Then Response.Write("SELECTED") :
							Response.Write("")%> >
							<%=(ptoventa.Fields.Item("nom_puntoventa").Value)%></option>
							<%ptoventa.MoveNext()
						Wend
						If (ptoventa.CursorType > 0) Then
							ptoventa.MoveFirst
						Else
							estados.Requery
						End If%>
					</select>
				</td>
				<td width="11"></td>
				<td width="247" class="presenta2">
					<select name="arti1" class="PPRFiel" id="select" tabindex="16" >
						<%While (NOT bdproducto.EOF)%>
							<option value="<%=(bdproducto.Fields.Item("cdart").Value)%>"
							 <%If (Not isNull((arti1))) Then If (trim(bdproducto.Fields.Item("cdart").Value)) = (trim(arti1)) then Response.Write("SELECTED") :
							  Response.Write("")%> ><%=(bdproducto.Fields.Item("cdart").Value)%><%=(response.write(" --- "))%>
							  <%=(bdproducto.Fields.Item("dsmod").Value)%></option>
						<%
							bdproducto.MoveNext()
						Wend
						If (bdproducto.CursorType > 0) Then
							bdproducto.MoveFirst
						Else
							bdproducto.Requery
						End If
						%>
					</select>
				</td>
			</tr>
		</table>
		<table width="100%" height ="6%"><tr></tr></table>
		<table align="center">
			<tr>
				<td colspan="2" align="center"><input type="submit"  name="btnBuscar" value="Modificar Registro"></button>
			</tr>
		</table>
	</form>
	    <table width="100%"  align="center">
    	    <tr>
				<td width="13%"></td>
				<td width="22%">
					<FORM NAME="volver" ACTION="menu.asp">
						<input type="submit" name="" value="Menú Principal">
					</FORM>
				</td>
				<td width="13%"></td>
				<td width="22%" align="right">
					<FORM NAME="volver" ACTION="modificadatos.asp">
						<input type="submit" name="asigna" value="Modificar Otro Prospecto">
					</FORM>
				</td>
				<td width="16%"></td>
	      </tr>
    	</table>
	</body>
</html>
<%
'dbplazo.Close()
'Set dbplazo = Nothing
'estados.Close()
'Set estados = Nothing
'captados.Close()
'Set captados = Nothing
'prefijos.Close()
'Set prefijos = Nothing
'tiempo_emp.Close()
'Set tiempo_emp = Nothing
'empleado.Close()
'Set empleado = Nothing
'ingreso.Close()
'Set ingreso = Nothing
'estados.Close()
'Set estados = Nothing
%>

 