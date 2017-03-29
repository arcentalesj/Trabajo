<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<% 
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Response.Expires = 0
%>
<%
session("inicio") = request.querystring("fecini")
session("fin")= request.querystring("fecfin")
select case session("reporte")
	case "Reporte ingresados a Excel"
		titulo = "PROSPECTOS INGRESADOS AL CSC"
		if session("cargo")="71" then
			strSQL = "SELECT *  FROM clientes_ingresados where idvendedor = '"& trim(session("usuario"))& "'"
			strSQL = strSQL + " and (contactado >= '"&session("inicio") &"' and contactado <= '"&session("fin")&"') order by nombres"
		else 
			strSQL = "SELECT *  FROM clientes_ingresados"
			strSQL = strSQL + " where (contactado >= '"&session("inicio") &"' and contactado <= '"&session("fin")&"') order by idvendedor,nombres"
		end if
	case "Citas Excel"
		titulo = "REPORTE DE CITAS A EXCEL"
		if session("cargo")="71" then
			strSQL = "SELECT *  FROM citas_del_dia where idvendedor = '"& trim(session("usuario"))& "'"
			strSQL = strSQL + " and (cita >= '"&session("inicio") &"' and cita <= '"&session("fin")&"')"
		else 
			strSQL = "SELECT * FROM citas_del_dia where "
			strSQL = strSQL + " and (cita >= '"&session("inicio") &"' and cita <= '"&session("fin")&"')"
		end if
end select
response.write strsql
%>
<%
Dim captados
Dim captados_numRows
Dim Repeat1__numRows
Set captados = Server.CreateObject("ADODB.Recordset")
captados.ActiveConnection = MM_usuarios_STRING
captados.Source = strSql
captados.CursorType = 1
captados.CursorLocation = 2
captados.LockType = 1
captados.Open()
if captados.bof and captados.eof then
	response.redirect("menu.asp?mensajeerror=" &"cclientet_nohay")
end if
select case session("reporte")
	case "Reporte ingresados a Excel"
		if session("cargo")="71" then
			nombre_archivo="attachment;filename=ClienVende_"&(trim(session("usuario")))+".xls;" 
		else
			nombre_archivo="attachment;filename=clienIngTot.xls;" 
		end if
	case "Citas Excel"
		if session("cargo")="71" then
			nombre_archivo="attachment;filename=Citas_"&(trim(session("usuario")))+".xls;" 
		else
			nombre_archivo="attachment;filename=citasTot.xls;" 
		end if
end select
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
	<table align="center">
		<tr> 
			<td width="30"></td>
			<td width="70"></td>
			<td width="300"><%response.write(titulo)%></td>
		</tr>
		<tr> 
			<td width="30"></td>
			<td width="70"></td>
			<td width="300"><%response.write(session("titulo1"))%><%response.write(" -- ")%><%response.write(session("titulo"))%></td>
		</tr>
	</table>
	<table border="1" align="center">
		<tr> 
			<td width="30">Ord.</td>
			<td width="70">Cedula</td>
			<td width="300">Nombre del Cliente</td>
			<td width="90">Celular</td>
			<td width="90">Telefono1</td>
			<td width="90">Telefono2</td>
			<td width="150">Producto</td>
			<td width="150">Punto Venta</td>
			<td width="100">Fecha de Contacto</td>
			<td width="100">Fecha de Registro</td>
			<td width="90">Fecha de Cita</td>
			<%if session("cargo") <> "71" then%>
				<td width="90">Vendedor</td>
			<%end if%>
		</tr>
		<% contador=1
		While (NOT captados.EOF)%>
			<tr> 
				<%
				fono1= ("'")+trim(captados.Fields.Item("codarea1").Value)+(captados.Fields.Item("telefono").Value)
				fono2= ("'")+trim(captados.Fields.Item("codarea2").Value)+(captados.Fields.Item("fono2").Value)
				celu = ("'")+trim(captados.Fields.Item("celular").Value)
				if len(ltrim(rtrim(fono1)))<8 then
					fono1 =""
				end if
				if len(ltrim(rtrim(fono2)))<8 then
					fono2 =""
				end if
				if len(ltrim(rtrim(celu)))<8 then
					celu =""
				end if
				bien= captados.Fields.Item("producto").Value
				%>
				<td><%=(response.write(contador))%></td>
				<td><%=(response.write("'"))%><%=(captados.Fields.Item("cedula").Value)%></td>
				<td><%=(captados.Fields.Item("nombres").Value)%></td>
				<td><%=(response.write(celu))%></td>
				<td><%=(response.write(fono1))%></td>
				<td><%=(response.write(fono2))%></td>
				<td><%=(response.write(bien))%></td>
				<td><%=(captados.Fields.Item("ptoventa").Value)%></td>
				<td><%=(captados.Fields.Item("Contactado").Value)%></td>
				<td><%=(captados.Fields.Item("captado").Value)%></td>
			    <td><%=(captados.Fields.Item("cita").Value)%></td>
				<%if session("cargo") <> "71" then%>
					<td><%=(captados.Fields.Item("nomvende").Value)%></td>
				<%end if%>
			</tr>
			<%contador=contador+1
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
