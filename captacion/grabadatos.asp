<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<%
'
' Objeto de Conexion de bases de datos
Set captacion = CreateObject("ADODB.Recordset")
captacion.ActiveConnection = MM_usuarios_STRING
dim arr, fecha, fecha1,hoy,fecha2,fecha3
hoy= date()
dia = day(hoy)
anio = year(hoy)
mes = month(hoy)
hoy = dia&"/"&mes&"/"&anio
cestado = hoy+" "+cstr(hour(time()))+":"+cstr(minute(time()))+":"+cstr(second(time()))
select case session("estado")
	case 1
		' graba el ingreso de un Prospecto
		if trim(request.form("telefono2"))="" then
			codare2=""
			fono2=""
		else
			codare2=request.form("codarea2")
			fono2=request.form("telefono2")
		end if
		' fecha de contacto
		CadenaDLL = split(request.form("feregis"),"-",-1,1)
		dia = CadenaDLL(0)
		mes = CadenaDLL(1)
		anio = CadenaDLL(2)
		fechacon = dia&"/"&mes&"/"&anio
		if trim(request.form("cita")) <> "" then
			CadenaDLL = split(request.form("cita"),"-",-1,1)
			dia = CadenaDLL(0)
			mes = CadenaDLL(1)
			anio = CadenaDLL(2)
			fechacita = dia&"/"&mes&"/"&anio
		else
			fechacita = request.form("cita")
		end if		
		strSQL ="insert into seguimiento (seg_cedula, seg_nombre, seg_apellido,"
		strSQL = strSQL&"seg_direccion, seg_trabajo,"
		strSQL = strSQL&"seg_mail, seg_celular,"
		strSQL = strSQL&"seg_codarea1,seg_telefono,"
		strSQL = strSQL&"seg_codarea2,seg_telefono2,"
		strSQL = strSQL&"seg_civil,"
		strSQL = strSQL&"seg_articulo,"
		strSQL = strSQL&"seg_feccaptacion,"
		strSQL = strSQL&"seg_idvendedor,"
		strSQL = strSQL&"seg_idptovnta,"
		strSQL = strSQL&"seg_tipodoc,"
		strSQL = strSQL&"seg_feccontacto,"
		strSQL = strSQL&"seg_fec_cita,"
		strSQL = strSQL&"seg_anonac)"
		strSQL = strSQL&" values ('"& (Session("cedula"))&"','"&(ucase(request.form("nombre")))&"','"&(ucase(request.form("apellido")))
		strSQL = strSQL&"','"&(ucase(request.form("direccion"))) 
		strSQL = strSQL&"','"&(ucase(request.form("trabajo"))) 
		strSQL = strSQL&"','"&((request.form("correo"))) 
		strSQL = strSQL&"','"&((request.form("celular"))) 
		strSQL = strSQL&"','"&((request.form("codarea1")))&"','"&((request.form("telefono1")))
		strSQL = strSQL&"','"&(codare2)&"','"&(fono2)
		strSQL = strSQL&"','"&((request.form("ecivil"))) 
		strSQL = strSQL&"','"&(ltrim(rtrim(session("bien"))))
		strSQL = strSQL&"','"&(fechacon)
		strSQL = strSQL&"',"&(session("usuario"))
		strSQL = strSQL&",'" &(request.form("pto1"))
		strSQL = strSQL&"','"&(session("docum"))
		strSQL = strSQL&"','"&(cestado)
		strSQL = strSQL&"','"&(fechacita)
		strSQL = strSQL&"','"&(request.form("nacer"))&"')"
		captacion.Open strSQL, MM_usuarios_STRING
		strSQL = "SELECT * FROM captados WHERE cedula= '"+(Session("cedula"))+"'"
		captacion.Open strSQL, MM_usuarios_STRING
		if captacion.bof and captacion.eof then
			grabesi=0
		else
			grabesi=1
		end if
		captacion.close()
	case 2
		' modificacion de datos desde Centro de Negocios para clientes de CREDIAUTO
		' verifica codigo de area
		if trim(request.form("telefono2"))="" then
			codare2=""
			fono2=""
		else
			codare2=request.form("codarea2")
			fono2=request.form("telefono2")
		end if
		if InStr(request.form("feregis"),"-") then
			CadenaDLL = split(request.form("feregis"),"-",-1,1)
			dia = CadenaDLL(0)
			mes = CadenaDLL(1)
			anio = CadenaDLL(2)
			fechacon = dia&"/"&mes&"/"&anio
		else
			fechacon = request.form("feregis")
		end if
		if trim(request.form("cita")) <> "" then
			if InStr(request.form("cita"),"-") then
				CadenaDLL = split(request.form("cita"),"-",-1,1)
				dia = CadenaDLL(0)
				mes = CadenaDLL(1)
				anio = CadenaDLL(2)
				fechacita = dia&"/"&mes&"/"&anio
			else
				fechacita = request.form("cita")
			end if		
		else
			fechacita = ""
		end if		
		' arma el SQL para hacer el update
		strSQL ="update dbo.seguimiento set "
		strSQL = strSQL&"seg_nombre='"&(ucase(request.form("nombre")))&"', seg_apellido='"&(ucase(request.form("apellido")))& "', "
		strSQL = strSQL&"seg_direccion ='"&(ucase(request.form("direccion")))&"', seg_trabajo ='"&(ucase(request.form("trabajo")))&"', "
		strSQL = strSQL&"seg_mail ='"&(lcase(request.form("correo")))&"', seg_celular ='"&((request.form("celular")))& "', " 
		strSQL = strSQL&"seg_codarea1 ='"&((request.form("codarea1")))&"', seg_telefono ='"&((request.form("telefono1")))&  "'," 
		strSQL = strSQL&"seg_codarea2 ='"&(codare2)&"', seg_telefono2 ='"&(fono2)&  "'," 
		strSQL = strSQL&"seg_civil ='"&((request.form("ecivil")))& "', " 
		strSQL = strSQL&"seg_articulo ='"& (ltrim(rtrim(request.form("arti1"))))&"', " 
		strSQL = strSQL&"seg_anonac ='"&(request.form("nacer"))&"', " 
		strSQL = strSQL&"seg_feccaptacion ='"&(fechacon)&"', " 
		strSQL = strSQL&"seg_fec_cita ='"&(fechacita)&"' " 
		strSQL = strSQL&"where seg_cedula='"&(session("cedula"))&"'"
		captacion.Open strSQL, MM_usuarios_STRING
		strSQL = "SELECT * FROM captados WHERE cedula= '"+(Session("cedula"))+"'"
		captacion.Open strSQL, MM_usuarios_STRING
		if captacion.bof and captacion.eof then
			grabesi=0
		else
			grabesi=1
		end if
		captacion.close()
		'
	case 3
		'
		' pone estatus de eliminado al cliente
		'
		strSQL = "update seguimiento set seg_seguimiento = 'N',"
		strSQL = strSQL & "seg_feceli= '" & (hoy) & "' ,"
		strSQL = strSQL & "seg_elimina= '" & session("quien") & "' "
		strSQL = strSQL & " where seg_cedula = '" & session("cedula") & "'"
		captacion.Open strSQL, MM_usuarios_STRING
		strSQL = "SELECT * FROM seguimiento WHERE seg_cedula= '"+(Session("cedula"))+"'"
		captacion.Open strSQL, MM_usuarios_STRING
		elimino=(captacion.Fields.Item("seg_seguimiento").value)
		if ltrim(rtrim(elimino)) ="N" then
			grabesi=1
		else
			grabesi=0
		end if
		captacion.close()
		'
	case 4
		'
		' graba las cedulas no deseadas
		'
		grabesi=0
        if session("siexiste")=1 then
			if (request.form("paraactivar"))=1 then
				' ingresa el registro lo que desean buro
				strSQL ="update cedula_problemas set "
				strSQL = strSQL& "ced_nombre ='"&(ucase(request.form("nombre")))&"',"
				strSQL = strSQL& "ced_apellido ='"&(ucase(request.form("apellido")))&"',"
				strSQL = strSQL& "ced_motivo ='"&(ucase(request.form("motivo")))&"',"
				strSQL = strSQL& "ced_usrhabilita ="&(ucase(session("quien")))
				strSQL = strSQL& "ced_estatus ='"&((request.form("paraactivar")))&"'"
				strSQL = strSQL&" where ced_cedula = '"&session("cedula")&"'"
'				response.write (strSql)
				captacion.Open strSQL, MM_usuarios_STRING
				grabesi=1
			else
				strSQL ="update cedula_problemas set "
				strSQL = strSQL& "ced_nombre ='"&(ucase(request.form("nombre")))&"',"
				strSQL = strSQL& "ced_apellido ='"&(ucase(request.form("apellido")))&"',"
				strSQL = strSQL& "ced_motivo ='"&(ucase(request.form("motivo")))&"',"
				strSQL = strSQL& "ced_estatus ='"&((request.form("paraactivar")))&"'"
				strSQL = strSQL&" where ced_cedula = '"&session("cedula")&"'"
'				response.write (strSql)
				captacion.Open strSQL, MM_usuarios_STRING
				grabesi=1
			end if
		else
			' ingresa el registro lo que no desean buro
			strSQL ="insert into cedula_problemas (ced_cedula, ced_nombre, ced_apellido,"
			strSQL = strSQL&"ced_motivo,"
			strSQL = strSQL&"ced_usrregistra,"
			strSQL = strSQL&"ced_fecregistro,"
			strSQL = strSQL&"ced_estatus)"
			strSQL = strSQL&" values ('"& (Session("cedula"))&"','"&(ucase(request.form("nombre")))&"','"&(ucase(request.form("apellido")))
			strSQL = strSQL&"','"&(ucase(request.form("motivo")))
			strSQL = strSQL&"','"&((session("quien")))
			strSQL = strSQL&"','"&(cestado)
			strSQL = strSQL&"','2')"
'			response.write (strSql)
			grabesi=1
			captacion.Open strSQL, MM_usuarios_STRING
		end if
		strSQL = "SELECT * FROM seguimiento WHERE seg_cedula= '"+(Session("cedula"))+"'"
'		response.write (strSql)
		captacion.Open strSQL, MM_usuarios_STRING
		if captacion.bof and captacion.eof then
			'no ha estado ingresado como prospecto
		else
			captacion.close()
			strSQL = "update seguimiento set seg_seguimiento = 'N',"
			strSQL = strSQL & "seg_feceli= '" & (hoy) & "' ,"
			strSQL = strSQL & "seg_elimina= '" & session("quien") & "' "
			strSQL = strSQL & " where seg_cedula = '" & session("cedula") & "'"
'			response.write (strSql)
			captacion.Open strSQL, MM_usuarios_STRING
		end if
	case 5
		'
		' cambia asessor comercial a un cliente
		'
		strSQL = "update seguimiento set seg_idvendedor = " & (request.form("vende1")) 
		strSQL = strSQL & " where seg_cedula = '" & session("cedula") & "'"
'		response.write(strSQL)
		captacion.Open strSQL, MM_usuarios_STRING
		grabesi=1

end select
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
		<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
		<title>Captacion de Clientes</title>
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
		<p>&nbsp;</p>
		<table width="58%" border="1" align="center" bgcolor="#FFFFFF" bordercolor="#990000" cellspacing="0" cellpadding="0">
			<tr> 
				<%
					if (grabesi=1) then%>
						<td align="center"><p>Ingreso de datos satisfactorios</p></td>
					<%else%>                       
						<td align="center"><p>No se actualizo ningun dato</p></td>
					<%end if
				%>
			</tr>
		</table>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<table width="64%"  align="center">
			<tr> 
				<td width="26%" align="center">
					<%
					select case session("estado")
						case 1%>
							<FORM NAME="volver" ACTION="cedula_prospectos.asp">
								<input type="submit" name="" value="Ingresar otro prospecto">
							</FORM> 
						<%case 2%>
							<FORM NAME="volver" ACTION="modificadatos.asp">
								<input type="submit" name="" value="Modificar otro prospecto">
							</FORM> 
						<%case 3%>
							<FORM NAME="volver" ACTION="busqueda.asp">
								<div align="center"><input type="submit" name="asigna" value="Suspender Prospecto"></div>
							</FORM>
						<%case 4%>
							<FORM NAME="actualizar" ACTION="cedula_problemas.asp">
								<input type="submit" name="" value="Clientes no deseados">
							</FORM>
					<%end select%>
				</td> 
				<td width="59%" align="left"></td>
				<td width="41%" align="left"> 
					<FORM NAME="volver" ACTION="menu.asp">
						<input type="submit" name="" value="Menu Principal">
					</FORM>
				</td>
			</tr>
		</table>
	</body>
</html>
<%
set captacion = nothing
%>