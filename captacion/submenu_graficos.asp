<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% 
dim quien
if request.querystring("usuario") <> "" then
	session("quien")=request.querystring("usuario")
	session("codemi")=request.querystring("emisor")
end if
 %>

<html>
<head>
	<title>Captacion y seguimiento de clientes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
	<script LANGUAGE="JavaScript">
        function ChangeUrl(formulaire)
            {
            if (formulaire.ListeUrl.selectedIndex != 0)
                {
                location.href = formulaire.ListeUrl.options[formulaire.ListeUrl.selectedIndex].value;
                }
            else 
                {
                alert('Tienes que elegir un destino.');
                }
            }
    </script>
<body background="imagenes/paginapsc.jpg">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr> 
			<td><img name="logo" src="imagenes/logo.gif" width="314" height="82" border="0" alt=""></td>
			<td "imagenes/blancos.gif" width="100%" height="92" border="0" alt="" ><div align="center">
				<SCRIPT>
                    dows = new Array("Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado");
                    months = new Array("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre");
                    now = new Date();
                    dow = now.getDay();
                    d = now.getDate();
                    m = now.getMonth();
                    h = now.getTime();
                    y = now.getYear();
                    document.write(dows[dow]+" "+d+" de "+months[m]+" de "+y);
                   </SCRIPT></div>
			</td>
			<td><img name="texto" src="imagenes/texto.gif" width="289" height="92" border="0" alt=""></td>
		</tr>
	</table>
	<table width="85%" border = "0"  height="60">
		<td width="60%"  height="60">&nbsp;</td>
		<td width="40%" >&nbsp;</td>
	</table >
	<%' if isnumeric(session("quien")) then 
		session("consulta") = 0
		session("opcion") = 0
	'	if session("codemi") > 0  then 
			titulo= "SUPERVISOR -- "+(session("nombresup"))+" -- EQUIPO "+session("equipo")%> 
			<table align="center">
				<tr> 
					<td><div align="center"><font color="#000000"><strong><%response.write(titulo)%></strong></font></div></td>
				</tr>
				<tr> 
					<td height="20"><div align="center"><font color="#000000"><strong></strong></font></div></td>
				</tr>
			</table>
			<table width="200"  align="center">
				<tr> 
					<td width="200"  height="5" align="center">
						<FORM NAME="actualizar" ACTION="muestra_asesores.asp">
							<input type="submit" name="asigna" value="Puntos Vs Asesores">
						</FORM>
					</td>
				</tr>
				<tr> 
					<td width="200"  height="5" align="center">
						<FORM NAME="actualizar" ACTION="muestra_asesores.asp">
							<input type="submit" name="asigna" value="Cerrados Vs Punto">
						</FORM>
					</td>
				</tr>
			</table>
			<table width="300"  align="center">
				<tr> 
					<td width="300" align="center"> 
					</td>
				</tr>
			</table>
			<table width="600"  align="center">
				<tr>
					<td width="300" align="center">
					</td>
					<td width="300" align="center">
					</td>
				</tr>
			</table>
			<table width="300"  align="center">
				<tr> 
					<td width="300" align="center">
					</td>
				</tr>
			</table>
			<table width="700"  align="center">
					<tr>
						<td width="350">
						</td>
						<td width="350">
						</td>
					</tr>
			</table>
			<table width="100%"  align="center">
                <tr>
                <td width="14%">				</td>
                    <td width="86%">
						<FORM NAME="volver" ACTION="principal.asp">
                            <input type="submit" name="" value="Regresar">
                        </FORM>
					</td>
				<tr> 
				</tr>
				</tr>
			</table>
		<%'else
			'response.redirect("inicio.asp")
		'end if
	'else
	'	select case session("quien")%>
			<%'case "egasf","arellanow","arcentalesj","silvap","rodrigueze","cabrerad" %>
			<table width="200"  align="center">
				<tr> 
					<td width="200"  height="5" align="center">
						<FORM NAME="actualizar" ACTION="captados_asesor.asp">
							<input type="submit" name="asigna" value="Puntos Vs Asesores">
						</FORM>
					</td>
				</tr>
				<tr> 
				</tr>
				<tr> 
					<td width="400"  height="5" align="center">
					</td>
				</tr>
				<tr> 
					<td width="400"  height="5" align="center">
					</td>
				</tr>
			</table>
		<% 'end select
	'end if%>
	<table width="59%" height="24"  align="center">
		<tr>
			<td width="26%" align="center"> 
				<p>
					<%
					select case request.querystring("mensajeerror")
						case "scliente_nohay"
							response.write ("<center><b>No existe cientes con coberturas CON ENDOSOS</b></center>")	
						case "ccliente_nohay"
							response.write ("<center><b>Lo sentimos, este cliente ya esta ingresado en el PSC</b></center>")	
						case "centro_nohay"
							response.write ("<center><b>No existe Informacion para ese centro de negocios</b></center>")	
						case "vende_nohay"
							response.write ("<center><b>No existe Vendedores en ese Centro de Negocios</b></center>")	
						case "vencidas_noadj"
							response.write ("<center><b>No existen coberturas vencidas a la fecha</b></center>")	
						case "cclientet_nohay"
							response.write ("<center><b>No existe datos para esta Consulta</b></center>")	
						case "asesorpend_nohay"
							response.write ("<center><b>No existe Asesores pendientes de comunicarse</b></center>")	
						case "cagentet_nohay"
							response.write ("<center><b>No existen clientes asignados a Usted</b></center>")	
						case "agente_nohay"
							response.write ("<center><b>No existe este agente en las bases de datos</b></center>")	
						case "cliente_yagrabado"
							response.write ("<center><b>El seguro del Cliente ya fue ingresado</b></center>")	
						case "clienteest_nohay"
							response.write ("<center><b>No existen clientes en este estatus</b></center>")	
						case "llamando"
							response.write ("<center><b>Este asesor comercial ya esta siendo contactado,intente con otro</b></center>")	
					end select
					%>
				</p>
			</td>
		</tr>
	</table>
</body>
</html>