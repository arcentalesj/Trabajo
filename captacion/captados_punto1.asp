<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<%

if trim(request.form("cboMeses"))<>"" then
	session("mes")=request.form("cboMeses")
end if
if trim(request.form("cboAnio"))<>"" then
	session("anio")=request.form("cboAnio")
end if
%>

<%
titulo = "CLIENTES CAPTADOS POR PUNTO DE VENTA DEL "&session("mes")&"/"&session("anio")
strSQL = "SELECT * FROM puntosdeventa where anio ="&session("anio")&" and mes = "&session("mes")& " order by anio,mes,codpunto,codvende"
response.write strSQL
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
if captados.eof and captados.bof then
	response.redirect("menu.asp?mensajeerror=" &"cclientet_nohay")
end if
nombre_archivo="attachment;filename=clientes_captados.xls;" 
%>
<% 
captados_numRows = 0
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", nombre_archivo
%>

<html>

<head>
<title>Captacion del Cliente por mes y año</title>
</head>
<body background="imagenes/paginapsc.jpg">
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
			<td width="300" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Punto de Venta</font></strong></td>
			<td width="300" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Asesor</font></strong></td>
			<td width="50" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">No.Captados</font></strong></td>
			<td width="40" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Anio</font></strong></td>
			<td width="40" align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Mes</font></strong></td>
		</tr>
			<% While (NOT captados.EOF)%>
			<td width="300"><div align="left"><font size="2"><%=(captados.Fields.Item("nombre_punto").Value)%></font></div></td>
			<td width="300"><div align="left"><font size="2"><%=(captados.Fields.Item("nombre_vende").Value)%></font></div></td>
			<td width="50"><div align="right"><font size="2"><%=(captados.Fields.Item("cuantos").Value)%></font></div></td>
			<td width="40"><div align="right"><font size="2"><%=(captados.Fields.Item("anio").Value)%></font></div></td>
			<td width="40"><div align="right"><font size="2"><%=(captados.Fields.Item("mes").Value)%></font></div></td>
		</tr>
				<%Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				captados.MoveNext()
			Wend%>
	</table>
</body>
</html>