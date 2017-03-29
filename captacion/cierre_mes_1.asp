<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<%
' Objeto para correos
Set nMail = Server.CreateObject("Persits.Mailsender")

' Objeto de Conexion de bases de datos
Set cerrar = CreateObject("ADODB.Recordset")
cerrar.ActiveConnection = MM_usuarios_STRING
Set asesor = CreateObject("ADODB.Recordset")
asesor.ActiveConnection = MM_usuarios_STRING
Set historico = CreateObject("ADODB.Recordset")
historico.ActiveConnection = MM_usuarios_STRING
Set seguir = CreateObject("ADODB.Recordset")
seguir.ActiveConnection = MM_usuarios_STRING
dim hoy
'
hoy = date()
dia = day(hoy)
anio = year(hoy)
mes = month(hoy)
hoy = anio&"/"&mes&"/"&dia
hora= time()
cestado=hoy+" "+cstr(hora)
if mes = 1 then
	anio= anio-1
	mescierre=12
else
	mescierre= mes-1
end if
'
equipos=0
equiposact = split(session("equipo"),",",-1,1)
for i = 1 to len(trim(session("equipo")))
	if mid(session("equipo"),i,1)="," then
		equipos=equipos+1
	end if
next
equipos=equipos+1
for i=0 to equipos-1
	strSQL0 = "SELECT * FROM cierre where cierre_equipo ='"&trim(equiposact(i))&"' and cierre_mes='"&(mescierre)&"' and cierre_anio='"&(anio)&"'"
	cerrar.Source = strSQL0
	cerrar.Open()
	if cerrar.bof and cerrar.eof then
		' No hay información de este select
		cerrar.close()
'	 	Procede con el cierre de mes definitivo
		strSQL ="insert into cierre (cierre_equipo, cierre_fecha,cierre_mes,cierre_anio) "
		strSQL = strSQL + "values ('"&trim(equiposact(i))&"','"&(cestado)&"','" &(mescierre)&"','"&(anio)&"')"
'		response.write (strSql)
		cerrar.Open strSQL, MM_usuarios_STRING
		nocierra= "N"
	else
		' Si hay información de este select
		cerrar.close()
		nocierra= "S"
	end if
next
if nocierra="N" then
	strSQL1= "Select * from asesores where cdequipo in ("&session("equipo")&") order by codven"
'	response.write (strSQL1)
	asesor.Source = strSQL1
	asesor.Open()
	While ((NOT asesor.EOF))
		vendedor=(asesor.Fields.Item("codven").Value)
		' copiar a historicos para graficos anteriores
		strSQL3= "update seguimiento set seg_seguimiento='N'"
		strSQL3= strSQL3&"WHERE (seg_idvendedor='"&(trim(vendedor))&"') AND (YEAR(seg_fec_cita)="&(anio)&") AND (seg_grupo = 'X')"
		strSQL3= strSQL3&" AND (MONTH(seg_feccontacto)<="&(anio)&") OR (seg_idvendedor='"&(trim(vendedor))&"') AND (seg_grupo = 'X')"
		strSQL3= strSQL3&" AND (MONTH(seg_feccontacto)<="&(anio)&") AND (seg_fec_cita IS NULL)"
'		response.write (strSQL3)
		seguir.open strSQL3, MM_usuarios_STRING
		asesor.MoveNext()
	Wend
end if
%>
<html>
	<head>
		<meta http-equiv="Content-Type"
			content="text/html; charset=iso-8859-1">
		<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">
		<title>Captacion de Clientes</title>
	</head>
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
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<table width="58%" border="1" align="center" bgcolor="#FFFFFF" bordercolor="#990000" cellspacing="0" cellpadding="0">
	        <% if nocierra="N" then %>
				<tr> 
					<td align="center"><p>Cierre de mes realizado correctamente</p></td>
				</tr>
	        <% else %>
				<tr> 
					<td align="center"><p>El cierre de este Equipo de ventas ya fue realizado</p></td>
				</tr>
	        <% end if %>
		</table>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<table width="27%"  align="center">
			<tr> 
				<td width="26%" align="center">
				</td> 
			</tr>
		</table>
		<table width="30%" border="0" align="left" cellspacing="0" cellpadding="0">
			<tr> 
				<td align="left"> 
					<FORM NAME="volver" ACTION="principal.asp">
						<input type="submit" name="" value="Regresar al menu principal">
					</FORM>
				</td>
			</tr>
		</table>
	</body>
</html>
<%
set rspagado = nothing
%>
