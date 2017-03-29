<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<% 
   Response.addHeader "pragma", "no-cache"
   Response.CacheControl = "Private"
   Response.Expires = 0
%>
<%
If (Request.QueryString("offset") = "") Then 
	session("inicio")=request.form("fecini")+" 00:00:00"
	session("fin")=request.form("fecfin")+" 23:59:00"
End If
titulo = "CLIENTES DEL PSC QUE INGRESARON AL SISTEMA "
if session("codemi") >0 then
	select case session("tipocliente")
		case "Asesor"
			strSQL = "SELECT *  FROM psc_con_caja where idvendedor = '"& trim(session("usuario"))& "'"
			strSQL = strSQL + " and (ingresopsc >= '"&session("inicio") &"' and ingresopsc <= '"&session("fin")&"')"
			titulo1= "ASESOR -- "+(session("nombrease"))+" -- CODIGO "+session("USUARIO") 
		case "sup"
			strSQL = "SELECT * FROM psc_con_caja where equipo in("&(session("equipo"))&")"
			strSQL = strSQL + " and (ingresopsc >= '"&session("inicio") &"' and ingresopsc <= '"&session("fin")&"')"
			titulo1= "SUPERVISOR -- "+(session("nombresup"))+" -- EQUIPO "+session("equipo") 
		case "reg"
			strSQL = "SELECT *  FROM psc_con_caja where equipo in("&(session("equipo"))&")"
			strSQL = strSQL + " and (ingresopsc >= '"&session("inicio") &"' and ingresopsc <= '"&session("fin")&"')"
			titulo1= (session("nombrereg"))
	end select
else
	strSQL = "SELECT * FROM psc_con_caja"
	strSQL = strSQL + " where (ingresopsc >= '"&session("inicio") &"' and ingresopsc <= '"&session("fin")&"')"
end if
%>
<%
Dim captados
Dim captados_numRows
Dim captados__MMColParam
captados__MMColParam = "1"
Set captados = Server.CreateObject("ADODB.Recordset")
captados.ActiveConnection = MM_usuarios_STRING
captados.Source = strSQL
captados.CursorType = 0
captados.CursorLocation = 2
captados.LockType = 1
captados.Open()
if captados.eof and captados.bof then
	if session("codemi") >0 then
		select case session("tipocliente")
			case "sup"
				response.redirect("submenu.asp?mensajeerror=" &"cclientet_nohay")
			case "Asesor"
				response.redirect("menu.asp?mensajeerror=" &"cclientet_nohay")
			case "reg"
				response.redirect("submenu_regional.asp?mensajeerror=" &"cclientet_nohay")
		end select
	else
		response.redirect("principal.asp?mensajeerror=" &"cclientet_nohay")
	end if
end if
if session("codemi") >0 then
	select case session("tipocliente")
		case "sup"
			nombre_archivo="attachment;filename=clien_ingcaja_super.xls;" 
		case "Asesor"
			nombre_archivo="attachment;filename=clien_ingcaja_"&(trim(session("usuario")))+".xls;" 
		case "reg"
			titulo = titulo + " DE LA REGIONAL DE "
			nombre_archivo="attachment;filename=clien_ingcaja_reg.xls;" 
	end select
else
	titulo = titulo + " A NIVEL NACIONAL"
	nombre_archivo="attachment;filename=clien_ingcaja_nac.xls;" 
end if
%>
<% 
rs_numRows = 0
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", nombre_archivo
%>
<html>
<head>
<title>Envio de Informacion A excel</title>
</head>

<body>
	<table width="100%" border="0">
		<tr> 
			<td align="center"><font color="#990000" size="3"><b><%response.write(titulo)%></b></font></td>
		</tr>
		<tr> 
			<td align="center"><font color="#990000" size="3"><b><%response.write(titulo1)%></b></font></td>
		</tr>
	</table>
	<table border="1" align="center">
		<tr> 
			<td width="30" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Ord.</font></strong></td>
			<td width="80" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Cedula</font></strong></td>
			<td width="280" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Nombre</font></strong></td>
			<td width="300" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Telefono1</font></strong></td>
			<td width="300" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Telefono2</font></strong></td>
			<td width="300" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Celular</font></strong></td>
			<td width="85" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Fecha Contacto</font></strong></td>
			<td width="85" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Fecha Ing. Caja</font></strong></td>
			<% if session("codemi") >0 then
				select case session("tipocliente")
					case "reg"%>
						<td width="150" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Vendedor</font></strong></td>
						<td width="150" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Centro de Negocios</font></strong></td>
					<%case "sup"%>
						<td width="150" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Vendedor</font></strong></td>
					<%end select%>
			<% else%>
				<td width="150" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Vendedor</font></strong></td>
				<td width="150" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Centro de Negocios</font></strong></td>
			<% end if%>
			<td width="280" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Punto de Venta</font></strong></td>
			<td width="60" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Tiempo en días</font></strong></td>
			<td width="60" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Contrato No.</font></strong></td>
			<td width="60" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Cotizador No.</font></strong></td>
			<td width="90" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Grupo</font></strong></td>
			<td width="90" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Orden</font></strong></td>
			<td width="90" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Bien Adquirido</font></strong></td>
		</tr>
		<% contador=1
		While (NOT captados.EOF)
			if (captados.Fields.Item("idvendedor").Value)>1 then%>
				<tr> 
					<%fechareg = ""
					fechacap = ""
					fecingca = ""
					if isdate(captados.Fields.Item("ingresopsc").Value) then
						fechareg = captados.Fields.Item("ingresopsc")
						anio = year(fechareg)
						mes = month(fechareg)
						dia = day(fechareg)
						fechareg = dia &"/" &mes &"/" &anio
					end if
					if isdate(captados.Fields.Item("ingresocaja").Value) then
						fecingca = captados.Fields.Item("ingresocaja")
						anio3 = year(fecingca)
						mes3 = month(fecingca)
						dia3 = day(fecingca)
						fecingca = dia3 &"/" &mes3 &"/" &anio3
					end if
					if len(trim(fechareg)) = 0 then 
						tiempo = datediff("d",captados.Fields.Item("ingresopsc").Value,captados.Fields.Item("ingresocaja").Value)
					ELSE 
						tiempo = datediff("d",captados.Fields.Item("ingresopsc").Value,captados.Fields.Item("ingresocaja").Value)
					END IF
					nombre = trim(captados("apellidos")) & " "& trim(captados("nombres"))%>
					<td><div align="center"><font size="2"><%=(response.write(contador))%></font></div></td>
					<td><font size="2" face="Times New Roman, Times, serif"><%=(response.write(nombre))%></font></td>
					<td><div align="center"><font size="2"><%=captados.Fields.Item("cedula").Value %></font></div></td>
					<td><font size="2"><%=(captados.Fields.Item("telefono1").Value)%></font></td>
					<td><font size="2"><%=(captados.Fields.Item("telefono2").Value)%></font></td>
					<td><font size="2"><%=(captados.Fields.Item("celular").Value)%></font></td>
					<td><div align="right"><font size="2"><%=(response.write(fechareg))%></font></div></td>
					<td><div align="right"><font size="2"><%=(response.write(fecingca))%></font></div></td>
					<% if session("codemi") >0 then
						select case session("tipocliente")
							case "reg"%>
								<td><font size="2"><%=(captados.Fields.Item("nombrevende").Value)%></font></td>
						    	<td><font size="2"><%=(captados.Fields.Item("centro").Value)%></font></td>
							<%case "sup"%>
								<td><font size="2"><%=(captados.Fields.Item("nombrevende").Value)%></font></td>
							<%end select%>
					<% else%>
						<td><font size="2"><%=(captados.Fields.Item("nombrevende").Value)%></font></td>
				    	<td><font size="2"><%=(captados.Fields.Item("centro").Value)%></font></td>
					<% end if%>
				    <td><font size="2"><%=(captados.Fields.Item("Ptoventa").Value)%></font></td>
					<td><div align="right"><font size="2"><%=(response.write(tiempo))%></font></div></td>
					<td><div align="right"><font size="2"><%=(response.write(captados("contrato")))%></font></div></td>
					<td><div align="right"><font size="2"><%=(response.write(captados("Cotizador")))%></font></div></td>
				    <td><font size="2"><%=(captados.Fields.Item("grupo").Value)%></font></td>
				    <td><font size="2"><%=(captados.Fields.Item("orden").Value)%></font></td>
				    <td><font size="2"><%=(captados.Fields.Item("bien").Value)%></font></td>
				</tr>
			<%contador=contador+1
			end if
			Repeat1__index=Repeat1__index+1
			Repeat1__numRows=Repeat1__numRows-1
			captados.MoveNext()
		Wend%>
	</table>
</body>
</html>
<%
captados.Close()
Set captados = Nothing
%> 
