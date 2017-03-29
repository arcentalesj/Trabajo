<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/usuarios.asp" -->
<% 
dim quien
Dim mensajeerror
If (Request.QueryString("mensajeerror") <> "") Then 
	session("dd") = Request.QueryString("mensajeerror")
End If
if session("quien") <> "" then
	session("quien1")=session("quien")
	session("codemi1")=session("codemi")
else
	session("quien")=session("quien1")
	session("codemi")=session("codemi1")
end if
%>
<html>
<head>
	<title>Captacion y seguimiento de clientes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
	<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
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
	<%session("titulo")= (session("nombrease"))&" -- CODIGO "&session("usuario")
	session("titulo1")= session("tipocliente")%>
	<table align="center">
		<tr> 
			<td class="titulo"><%response.write(session("titulo"))%></td>
		</tr>
		<tr> 
			<td class="subtitulo"><%response.write(session("titulo1"))%></td>
		</tr>
	</table>
	<%
	select case session("cargo")
		case "01"%>
			<p>&nbsp;</p>
			<table width="600"  align="center">
				<tr>
					<td width="300"  align="center">
						<FORM NAME="actualizar" ACTION="clientes_ingresados.asp">
							<input type="submit" name="" value="Consultar Prospectos">
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="300" align="center"> 
						<FORM NAME="actualizar" ACTION="busqueda.asp">
							<input type="submit" name="asigna" value="Busqueda de Prospectos">
						</FORM>					</td>
				</tr>
				<tr> 
					<td width="300" align="center">
						<FORM NAME="actualizar" ACTION="muestra_citas.asp">
							<input type="submit" name="" value="Citas para hoy">
						</FORM>
					</td>
				</tr>
			</table>
		<%case "02"
			' Menú para el Analista de Operaciones
			%>
			<p>&nbsp;</p>
			<table width="600"  align="center">
				<tr> 
					<td width="200"  align="center">
						<FORM NAME="actualizar" ACTION="clientes_ingresados.asp">
							<input type="submit" name="asigna" value="Clientes Ingresados">
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="300" align="center"> 
						<FORM NAME="actualizar" ACTION="busqueda.asp">
							<input type="submit" name="asigna" value="Busqueda de Prospectos">
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="300" align="center">
						<FORM NAME="actualizar" ACTION="cambia_asesor.asp">
							<input type="submit" name="" value="Reasignación de Vendedor">
						</FORM>
					</td>
				</tr>
				<tr> 
					<td width="300" align="center">
						<FORM NAME="actualizar" ACTION="muestra_citas.asp">
							<input type="submit" name="" value="Citas para hoy">
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="300" align="center">
						<FORM NAME="volver" ACTION="captados_punto.asp">
							<input type="submit" name="" value="Prospectos Captados por Punto">
						</FORM>
					</td>
				</tr>
			</table>

            <%'end if
		case "71"
			' Menú de ventas
			session("consulta") = 0
			session("opcion") = 0%> 
			<p>&nbsp;</p>
			<table width="304"  align="center">
				<tr> 
					<td width="300"  align="center">
						<FORM NAME="actualizar" ACTION="cedula_prospectos.asp">
							<input type="submit" name="" value="Prospectos a Clientes">
						</FORM>
					</td>
				</tr>
				<tr> 
					<td width="300"  align="center">
						<FORM NAME="actualizar" ACTION="modificadatos.asp">
							<input type="submit" name="" value="Modificar datos">
						</FORM>
					</td>
				</tr>
				<tr> 
					<td width="300"  align="center">
						<FORM NAME="actualizar" ACTION="clientes_ingresados.asp">
							<input type="submit" name="" value="Consultar Prospectos">
						</FORM>
					</td>
				</tr>
				<tr> 
					<td width="300" align="center">
						<FORM NAME="actualizar" ACTION="control_de_citas.asp">
							<input type="submit" name="" value="Control de Citas a Clientes">
						</FORM>
					</td>
				</tr>
				<tr> 
					<td width="300" align="center">
						<FORM NAME="actualizar" ACTION="muestra_citas.asp">
							<input type="submit" name="" value="Citas para hoy">
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="300" align="center"> 
						<FORM NAME="actualizar" ACTION="busqueda.asp">
							<input type="submit" name="asigna" value="Busqueda de Prospectos">
						</FORM>					</td>
				</tr>
			</table>
		<%case "11"
			' Menú para el Analista de Sistemas
			session("consulta") = 0
			session("opcion") = 0
			session("codemi")=0%>
			<p>&nbsp;</p>
			<table width="15%"  align="center">
				<tr>
					<td width="200"  height="5" align="center">
						<FORM NAME="actualizar" ACTION="ingptoventa.asp">
							<input type="submit" name="" value="Tabla Puntos de ventas">
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="200"  align="center">
						<FORM NAME="actualizar" ACTION="clientes_ingresados.asp">
							<input type="submit" name="asigna" value="Consultar Prospectos">
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="200" align="center"> 
						<FORM NAME="actualizar" ACTION="busqueda.asp">
							<input type="submit" name="asigna" value="Busqueda de Prospectos">
						</FORM>
					</td>
				</tr>
				<tr> 
					<td width="200">
						<FORM NAME="actualizar" ACTION="captados_punto.asp">
							<div align="center"><input type="submit" name="" value="Reporte Clientes captados por Punto"></div>
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="200">
						<FORM NAME="volver" ACTION="busqueda.asp">
							<div align="center"><input type="submit" name="asigna" value="Suspender Prospecto"></div>
						</FORM>
					</td>
				</tr>
			</table>
		<%case "04"
			session("consulta") = 0
			session("opcion") = 0
			session("codemi")=0%>
			<p>&nbsp;</p>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td width="200"  align="center">
						<FORM NAME="actualizar" ACTION="clientes_ingresados.asp">
							<input type="submit" name="asigna" value="Consultar Prospectos">
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="200" align="center"> 
						<FORM NAME="actualizar" ACTION="busqueda.asp">
							<input type="submit" name="asigna" value="Busqueda de Prospectos">
						</FORM>
					</td>
				</tr>
				<tr> 
					<td width="300" align="center">
						<FORM NAME="actualizar" ACTION="control_de_citas.asp">
							<input type="submit" name="" value="Control de Citas a Clientes">
						</FORM>
					</td>
				</tr>
				<tr> 
					<td width="300" align="center">
						<FORM NAME="actualizar" ACTION="muestra_citas.asp">
							<input type="submit" name="" value="Citas para hoy">
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="200">
						<FORM NAME="volver" ACTION="busqueda.asp">
							<div align="center"><input type="submit" name="asigna" value="Suspender Prospecto"></div>
						</FORM>
					</td>
				</tr>
				<tr>
					<td width="55%">
						<div align="center"> 
							<FORM NAME="actualizar" ACTION="cedula_problemas.asp">
								<input type="submit" name="" value="Clientes no deseados">
							</FORM>
						</div>
					</td>
				</tr>
				<tr>
					<td width="300" align="center">
						<FORM NAME="actualizar" ACTION="cambia_asesor.asp">
							<input type="submit" name="" value="Reasignación de Vendedor">
						</FORM>
					</td>
				</tr>
			</table>
		<%case else
		if trim(session("cargo"))="" then
			response.redirect("inicio.asp")
		else %>
			<p>&nbsp;</p>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr> 
					<td>
						<div align="center"><img src="imagenes/prohibido.png" alt="" name="Prohibido" width="320" height="335" border="0" align="middle">
							</div>
					</td>
				</tr>
			</table>
		<%end if%>
<%end select%>
	<table width="100%"  align="center">
		<tr>
			<td width="14%"></td>
			<td width="86%">
				<FORM NAME="volver" ACTION="principal.asp">
					<input type="submit" name="" value="Regresar">
				</FORM>
			</td>
		</tr>
	</table>			

<!--		
	'if control>hora then
'		session("tipocliente")=""
 '       strSQL = "select * from equipos where ltrim(rtrim(alias)) = '" + session("quien") +"'"
  '      Set requipos = CreateObject("ADODB.Recordset")
   '     requipos.Open strSQL, MM_usuarios_STRING

				<%' end select
				' no es nada, un usuario de caja que quiere ingresar %>
			<%'else
				' toma todos los equipos
'				equipos = 1
'				While (NOT rregional.EOF)
'					if equipos = 1 then
'						session("equipo") = cstr(rregional.Fields.Item("equipos").Value)
'					else
'						session("equipo") = session("equipo")+","+cstr(rregional.Fields.Item("equipos").Value)
'					end if
'					equipos=equipos+1
'					rregional.MoveNext()
'				Wend
'				session("equipo") = session("equipo")+",1"
				' toma el nombre el regional
'				strSQLREN = "SELECT * FROM regionaln WHERE codigo ='"&session("quien")&"'"
'				Set rregionaln = CreateObject("ADODB.Recordset")
'				rregionaln.Open strSQLREN, MM_usuarios_STRING
'				if rregionaln.bof and rregionaln.eof then 
					' no tiene registrado el nombre en base de datos fox %>
					<p>&nbsp;</p>
					<p>&nbsp;</p>
                    
					<table border="0" cellpadding="0" cellspacing="0" width="91%">
						<tr> 
							<td><div align="center">
<!--                   	        	<img src="imagenes/prohibido.gif" alt="" name="Prohibido" width="320" height="335" border="0" align="middle"></div>-->
<!--							</td>
						</tr>
					</table> -->
				<%'else
'					session("nombrereg")=(rregionaln.Fields.Item("nombre").Value)
'					response.redirect("submenu_regional.asp")
'				end if
'			end if
'		else
			' aqui ingresa cuando es superviosr
			'
'			session("tipocliente")="sup"
'			session("nombresup")=(requipos.Fields.Item("nombre").Value)
'			session("equipo")=""
'			session("codemi")=(requipos.Fields.Item("cdemi").Value)
'			session("quien")=(requipos.Fields.Item("codven").Value)
'			equipos = 1
'			While (NOT requipos.EOF)
'				if equipos = 1 then
'					session("equipo") = cstr(requipos.Fields.Item("codeqv").Value)
'				else
'					session("equipo") = session("equipo")+","+cstr(requipos.Fields.Item("codeqv").Value)
'				end if
'				equipos=equipos+1
'				requipos.MoveNext()
'			Wend
'			if (session("quien"))="faresm" then
'				session("equipo") = session("equipo")+",93252"
'			end if
'			session("equipo") = session("equipo")+",1"
'			if session("codemi") > 0  then 
'				titulo= "SUPERVISOR -- "+(session("nombresup"))+" -- EQUIPO "+session("equipo")	%> 

<%'			response.write	session("codemi")
	'		response.write	session("quien")
	'		response.write	session("nombrease")
	'		response.write	session("equipo")
	'		response.write	session("tipocliente")
	'		response.write	session("cargo")%>
<!--				<table align="center">
					<tr> 
						<td><div align="center"><font color="#000000"><strong><%'response.write(titulo)%></strong></font></div></td>
					</tr>
					<tr> 
						<td height="20"><div align="center"><font color="#000000"><strong></strong></font></div></td>
					</tr>
				</table>
				<table width="300" align="center">
					<tr> 
                      <td width="300" align="center">
                          <FORM NAME="ingresar" ACTION="submenuclientes.asp">
                              <input type="submit" name="input" value="Ingreso de Clientes">
                          </FORM>					</td>
                    </tr>
                    <tr>
                        <td width="300" align="center">
                            <FORM NAME="actualizar" ACTION="submenu.asp">
                                <input type="submit" name="" value="Seguimiento de Clientes">
                            </FORM>					</td>
                    </tr>
                    <tr>
                        <td width="300" align="center"> 
                            <FORM NAME="actualizar" ACTION="busqueda.asp">
                                <input type="submit" name="" value="Busqueda de Prospectos">
                            </FORM>					</td>
                    </tr>
                    <tr>
                        <td width="300" align="center"> 
                            <FORM NAME="actualizar" ACTION="submenu_graficos.asp">
                                <input type="submit" name="" value="graficos">
                            </FORM>					</td>
                    </tr>
					<tr>
							<td width="300" align="center"> 
								<FORM NAME="actualizar" ACTION="cierre_mes.asp">
									<input type="submit" name="" value="Cierre de Mes">
								</FORM>					</td>
					</tr>
                 </table>
            <%'end if
       ' end if%>
	<%'else%>
<!--		<table width="63%" border="0" align="center" cellpadding="0" cellspacing="0">
				<p>&nbsp;</p>
				<p>&nbsp;</p>
			<tr> 
				<td width="47%"><div align="center">
				<img src="imagenes/llegartarde.gif" alt="" name="Prohibido" width="180" height="163" border="0" align="middle"></div></td>
				<td width="53%"><div align="center">
				<strong>
				<%'response.write("Lo sentimos, No esta permitido ingresar mas clientes a partir de las ")%><%'response.write(control)%><%'response.write(":00 horas")%>
				</strong>
				<div align="center"><strong><%'response.write("En nuestros servidores son las ")%><%'response.write((time()))%></strong></div>      </div></td>
			</tr>
		</table>
	<%'end if%>-->

	<table width="59%" height="24"  align="center">
		<tr>
			<td width="26%" align="center"> 
				<p>
					<%
					select case request.querystring("mensajeerror")
'						case "cedula_nohay"
'							response.write ("<center><b>No existe este Asesor</b></center>")	
'						case "ccliente_nohay"
'							response.write ("<center><b>Lo sentimos, este cliente ya esta ingresado en el PSC</b></center>")	
'						case "centro_nohay"
'							response.write ("<center><b>No existe Informacion para ese centro de negocios</b></center>")	
						case "vende_nohay"
							response.write ("<center><b>No existe Vendedores en ese Centro de Negocios</b></center>")	
'						case "vencidas_noadj"
'							response.write ("<center><b>No existen coberturas vencidas a la fecha</b></center>")	
						case "cclientet_nohay"
							response.write ("<center><b>No existe datos para esta Consulta</b></center>")	
'						case "asesorpend_nohay"
'							response.write ("<center><b>No existe Asesores pendientes de comunicarse</b></center>")	
'						case "cagentet_nohay"
'							response.write ("<center><b>No existen clientes asignados a Usted</b></center>")	
'						case "agente_nohay"
'							response.write ("<center><b>No existe este agente en las bases de datos</b></center>")	
'						case "cliente_yagrabado"
'							response.write ("<center><b>El seguro del Cliente ya fue ingresado</b></center>")	
'						case "clienteest_nohay"
'							response.write ("<center><b>No existen clientes en este estatus</b></center>")	
'						case "llamando"
'							response.write ("<center><b>Este asesor comercial ya esta siendo contactado,intente con otro</b></center>")	
'						case "noautoriza"
'							response.write ("<center><b>Usted no esta autorizado para esta opción</b></center>")	
'						case "nocierre"
'							response.write ("<center><b>Usted no ha realizado el cierre de mes, no puede ingrear prospectos</b></center>")	
					end select
					%>
				</p>
			</td>
		</tr>
	</table>
</body>
</html>
