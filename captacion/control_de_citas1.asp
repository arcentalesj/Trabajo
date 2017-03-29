<%@LANGUAGE="VBSCRIPT"%>
<% session("estado")= 6 %>
<!--#include file="Connections/usuarios.asp" -->
<%
hoy = date()
dia = day(hoy)
anio = year(hoy)
mes = month(hoy)
hoy = mes&"/"&dia&"/"&anio
hora= time()
cestado=hoy+" "+cstr(hora)
If (Request.QueryString("cedula1") <> "") Then 
  session("cedula") = Request.QueryString("cedula1")
  session("codigo") = Request.QueryString("codigo")
End If
 Response.addHeader "pragma", "no-cache"
 Response.CacheControl = "Private"
 Response.Expires = 0
%>
<%
titulo = "REGISTRO DE CITAS"
usuario=session("quien")
Dim captados
Dim captados_numRows
Set captados = Server.CreateObject("ADODB.Recordset")
captados.ActiveConnection = MM_usuarios_STRING
captados.Source = "SELECT *  FROM control_citas  WHERE cedula ='" + session("cedula") +"'"
captados.CursorType = 0
captados.CursorLocation = 2
captados.LockType = 1
captados.Open()
if session("codemi") >0 then
	titulo1 = "ASESOR "+session("codigo")+ " "+ captados.Fields.Item("nomvende").value
else
	titulo = titulo 
end if
titulo1= "ASESOR -- "&(captados("nomvende"))&" -- CODIGO "&(session("asesor")) 
captados_numRows = 0
%>
<%
Dim obsseguim
Dim obsseguim_numRows
Set obsseguim = Server.CreateObject("ADODB.Recordset")
obsseguim.ActiveConnection = MM_usuarios_STRING
obsseguim.Source = "SELECT *  FROM dbo.obsseguim  WHERE obs_cedula ='"&(session("cedula"))&"'"
obsseguim.CursorType = 0
obsseguim.CursorLocation = 2
obsseguim.LockType = 1
obsseguim.Open()
obsseguim_numRows = 0
%>
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
	MM_editAction = MM_editAction & "?" & Request.QueryString
End If
MM_abortEdit = false
MM_editQuery = ""
%>
<%
If (CStr(Request("MM_insert")) = "formEvalua") Then
	MM_editConnection = MM_usuarios_STRING
	MM_editTable = "dbo.ObsSeguim"
	MM_editRedirectUrl = "control_de_citas1.asp"
	'  MM_fieldsStr  = "observa|value|fecha|value|hiddenField|value"
	' MM_columnsStr = "obs_Observa|',none,''|obs_fecha|',none,NULL|obs_ingresa|',none,''"
'				      obs_cedula	obs_fecha	obs_Observa	obs_ingresa	obs_evalua
	MM_fieldsStr  = "observa|value|fechao|value|hiddenField|value|usuario|value|evalua|value"
	MM_columnsStr = "obs_Observa|',none,''|obs_fecha|',none,''|obs_cedula |',none,''|obs_ingresa|',none,''|obs_evalua|',none,''"
	MM_fields = Split(MM_fieldsStr, "|")
	MM_columns = Split(MM_columnsStr, "|")
	For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
		MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
	Next
	If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
		If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
			MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
		Else
			MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
		End If
	End If
End If
%>
<%
If (CStr(Request("MM_update")) = "form2" And CStr(Request("MM_recordId")) <> "") Then
	MM_editConnection = MM_usuarios_STRING
	MM_editTable = "dbo.seguimiento"
	MM_editColumn = "seg_cedula"
	MM_recordId = "'" + Request.Form("MM_recordId") + "'"
	MM_editRedirectUrl = "control_de_citas1.asp"
	MM_fieldsStr  = "fecita1|value"
	MM_columnsStr = "seg_fec_cita|',none,'"
	MM_fields = Split(MM_fieldsStr, "|")
	MM_columns = Split(MM_columnsStr, "|")
	For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
		MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
	Next
	If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
		If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
			MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
		Else
			MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
		End If
	End If
End If
%>
<%
Dim MM_tableValues
Dim MM_dbValues
If (CStr(Request("MM_insert")) <> "") Then
	MM_tableValues = ""
	MM_dbValues = ""
	For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
		MM_formVal = MM_fields(MM_i+1)
		MM_typeArray = Split(MM_columns(MM_i+1),",")
		MM_delim = MM_typeArray(0)
		If (MM_delim = "none") Then MM_delim = ""
		MM_altVal = MM_typeArray(1)
		If (MM_altVal = "none") Then MM_altVal = ""
			MM_emptyVal = MM_typeArray(2)
			If (MM_emptyVal = "none") Then MM_emptyVal = ""
				If (MM_formVal = "") Then
					MM_formVal = MM_emptyVal
				Else
					If (MM_altVal <> "") Then
						MM_formVal = MM_altVal
				ElseIf (MM_delim = "'") Then  
					MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
				Else
					MM_formVal = MM_delim + MM_formVal + MM_delim
			End If
		End If
		If (MM_i <> LBound(MM_fields)) Then
			MM_tableValues = MM_tableValues & ","
			MM_dbValues = MM_dbValues & ","
		End If
		MM_tableValues = MM_tableValues & MM_columns(MM_i)
		MM_dbValues = MM_dbValues & MM_formVal
	Next
	MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"
	If (Not MM_abortEdit) Then
		' execute the insert
		Set MM_editCmd = Server.CreateObject("ADODB.Command")
		MM_editCmd.ActiveConnection = MM_editConnection
		MM_editCmd.CommandText = MM_editQuery
		MM_editCmd.Execute
		MM_editCmd.ActiveConnection.Close
		If (MM_editRedirectUrl <> "") Then
			Response.Redirect(MM_editRedirectUrl)
		End If
	End If
End If
%>
<%
If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then
	MM_editQuery = "update " & MM_editTable & " set "
	For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
		MM_formVal = MM_fields(MM_i+1)
		MM_typeArray = Split(MM_columns(MM_i+1),",")
		MM_delim = MM_typeArray(0)
		If (MM_delim = "none") Then MM_delim = ""
			MM_altVal = MM_typeArray(1)
			If (MM_altVal = "none") Then MM_altVal = ""
				MM_emptyVal = MM_typeArray(2)
				If (MM_emptyVal = "none") Then MM_emptyVal = ""
					If (MM_formVal = "") Then
						MM_formVal = MM_emptyVal
					Else
						If (MM_altVal <> "") Then
							MM_formVal = MM_altVal
					ElseIf (MM_delim = "'") Then
						MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
			Else
				MM_formVal = MM_delim + MM_formVal + MM_delim
			End If
		End If
		If (MM_i <> LBound(MM_fields)) Then
			MM_editQuery = MM_editQuery & ","
		End If
		MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
	Next
	MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId
	If (Not MM_abortEdit) Then
		Set MM_editCmd = Server.CreateObject("ADODB.Command")
		MM_editCmd.ActiveConnection = MM_editConnection
		MM_editCmd.CommandText = MM_editQuery
		MM_editCmd.Execute
		MM_editCmd.ActiveConnection.Close
		If (MM_editRedirectUrl <> "") Then
			Response.Redirect(MM_editRedirectUrl)
		End If
	End If
End If
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = 10
Repeat1__index = 0
obsseguim_numRows = obsseguim_numRows + Repeat1__numRows
%>
<%
Dim obsseguim_total
Dim obsseguim_first
Dim obsseguim_last
obsseguim_total = obsseguim.RecordCount
If (obsseguim_numRows < 0) Then
	obsseguim_numRows = obsseguim_total
Elseif (obsseguim_numRows = 0) Then
	obsseguim_numRows = 1
End If
obsseguim_first = 1
obsseguim_last  = obsseguim_first + obsseguim_numRows - 1
If (obsseguim_total <> -1) Then
	If (obsseguim_first > obsseguim_total) Then
		obsseguim_first = obsseguim_total
	End If
	If (obsseguim_last > obsseguim_total) Then
		obsseguim_last = obsseguim_total
	End If
	If (obsseguim_numRows > obsseguim_total) Then
		obsseguim_numRows = obsseguim_total
	End If
End If
%>
<%
If (obsseguim_total = -1) Then
	obsseguim_total=0
	While (Not obsseguim.EOF)
		obsseguim_total = obsseguim_total + 1
		obsseguim.MoveNext
	Wend
	If (obsseguim.CursorType > 0) Then
		obsseguim.MoveFirst
	Else
		obsseguim.Requery
	End If
	If (obsseguim_numRows < 0 Or obsseguim_numRows > obsseguim_total) Then
		obsseguim_numRows = obsseguim_total
	End If
	obsseguim_first = 1
	obsseguim_last = obsseguim_first + obsseguim_numRows - 1
	If (obsseguim_first > obsseguim_total) Then
		obsseguim_first = obsseguim_total
	End If
	If (obsseguim_last > obsseguim_total) Then
		obsseguim_last = obsseguim_total
	End If
End If
%>
<%
Dim MM_paramName 
%>
<%
Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined
Dim MM_param
Dim MM_index
Set MM_rs    = obsseguim
MM_rsCount   = obsseguim_total
MM_size      = obsseguim_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
	MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
if (Not MM_paramIsDefined And MM_rsCount <> 0) then
	MM_param = Request.QueryString("index")
	If (MM_param = "") Then
		MM_param = Request.QueryString("offset")
	End If
	If (MM_param <> "") Then
		MM_offset = Int(MM_param)
	End If
	If (MM_rsCount <> -1) Then
		If (MM_offset >= MM_rsCount Or MM_offset = -1) Then
			If ((MM_rsCount Mod MM_size) > 0) Then
				MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
			Else
				MM_offset = MM_rsCount - MM_size
			End If
		End If
	End If
	MM_index = 0
	While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
		MM_rs.MoveNext
		MM_index = MM_index + 1
	Wend
	If (MM_rs.EOF) Then 
		MM_offset = MM_index
	End If
End If
%>
<%
If (MM_rsCount = -1) Then
	MM_index = MM_offset
	While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
		MM_rs.MoveNext
		MM_index = MM_index + 1
	Wend
	If (MM_rs.EOF) Then
		MM_rsCount = MM_index
		If (MM_size < 0 Or MM_size > MM_rsCount) Then
			MM_size = MM_rsCount
		End If
	End If
	If (MM_rs.EOF And Not MM_paramIsDefined) Then
		If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
			If ((MM_rsCount Mod MM_size) > 0) Then
				MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
			Else
				MM_offset = MM_rsCount - MM_size
			End If
		End If
	End If
	If (MM_rs.CursorType > 0) Then
		MM_rs.MoveFirst
	Else
		MM_rs.Requery
	End If
	MM_index = 0
	While (Not MM_rs.EOF And MM_index < MM_offset)
		MM_rs.MoveNext
		MM_index = MM_index + 1
	Wend
End If
%>
<%
obsseguim_first = MM_offset + 1
obsseguim_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
	If (obsseguim_first > MM_rsCount) Then
		obsseguim_first = MM_rsCount
	End If
	If (obsseguim_last > MM_rsCount) Then
		obsseguim_last = MM_rsCount
	End If
End If
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth
Dim MM_removeList
Dim MM_item
Dim MM_nextItem
MM_removeList = "&index="
If (MM_paramName <> "") Then
	MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If
MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""
For Each MM_item In Request.QueryString
	MM_nextItem = "&" & MM_item & "="
	If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
		MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
	End If
Next
For Each MM_item In Request.Form
	MM_nextItem = "&" & MM_item & "="
	If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
		MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
	End If
Next
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
	MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
	MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
	MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If
Function MM_joinChar(firstItem)
	If (firstItem <> "") Then
		MM_joinChar = "&"
	Else
		MM_joinChar = ""
	End If
End Function
%>
<%
Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev
Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam
MM_keepMove = MM_keepBoth
MM_moveParam = "index"
If (MM_size > 1) Then
	MM_moveParam = "offset"
	If (MM_keepMove <> "") Then
		MM_paramList = Split(MM_keepMove, "&")
		MM_keepMove = ""
		For MM_paramIndex = 0 To UBound(MM_paramList)
			MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
			If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
				MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
			End If
		Next
		If (MM_keepMove <> "") Then
			MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
		End If
	End If
End If
If (MM_keepMove <> "") Then 
	MM_keepMove = MM_keepMove & "&"
End If
MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
	MM_movePrev = MM_urlStr & "0"
Else
	MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<html>
<head>
	<title>Control de Citas</title>
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
			<td class="subtitulo"><%response.write(titulo1)%></div></td>
		</tr>
	</table>
	<table width="100%" height ="2%"><tr></tr></table>
	<table width="837" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ACE5EE">
		<tr bgcolor="<%Response.Write(session("color"))%>"> 
			<td width="12">&nbsp;</td>
			<td width="139" class="presenta1">Cedula</td>
			<td width="288" class="presenta2"><%response.write(session("cedula"))%></td>
			<td width="13">&nbsp;</td>
			<td width="15">&nbsp;</td>
			<td width="139"class="presenta1">Fecha de Registro</td>
			<td width="289" class="presenta2"><%response.write(captados("registro"))%></td>
		</tr>
		<tr> 
			<td width="12">&nbsp;</td>
			<td width="139" class="presenta1">Apellidos</td>
			<td width="288" class="presenta2"><%response.write(captados("apellido"))%></td>
			<td width="13">&nbsp;</td>
			<td width="15" class="presenta1"></td>
			<td width="139" class="presenta1">Nombres</td>
			<td width="289" class="presenta2"><%response.write(captados("nombre"))%></td>
		</tr>
		<tr bgcolor="<%Response.Write(session("color"))%>"> 
			<td width="12">&nbsp;</td>
			<td width="139" class="presenta1">Direccion Domicilio</td>
			<td width="288" class="presenta2"><%response.write(TRIM(captados("direccion")))%></td>
			<td width="13">&nbsp;</td>
			<td width="15">&nbsp;</td>
			<td width="139" class="presenta1">Direccion Trabajo</td>
			<td width="289" class="presenta2"><%response.write(captados("trabajo"))%></td>
		</tr>
		<tr> 
			<td width="12">&nbsp;</td>
			<td width="139" class="presenta1">E_Mail</td>
			<td width="288" class="presenta2"><%response.write(captados("correo"))%></td>
			<td width="13">&nbsp;</td>
			<td width="15" class="presenta1"></td>
			<td width="139"class="presenta1">Celular</td>
			<td width="289" class="presenta2"><%response.write(captados("celular"))%></td>
		</tr>
		<tr bgcolor="<%Response.Write(session("color"))%>"> 
			<td width="12">&nbsp;</td>
			<td width="139" class="presenta1">Telefono 1</b></td>
			<td width="288" class="presenta2"><%response.write(captados("codarea1"))%><%response.write("-")%><%response.write(captados("telefono"))%></td>
			<td width="13">&nbsp;</td>
			<td width="15" class="presenta1"></td>
			<td width="139" class="presenta1">Telefono2</td>
			<td width="289" class="presenta2"><%response.write(captados("codarea2"))%><%response.write("-")%><%response.write(captados("fono2"))%></td>
	  </tr>
		<tr> 
			<td width="12">&nbsp;</td>
			<td width="139"><font color="#990000" size="2"><div align="Left"><b>Estado Civil</td>
			<td width="288" class="presenta2"><%response.write(captados("idestado"))%><%response.write(" - ")%><%response.write(captados("nomestado"))%></td>
			<td width="13">&nbsp;</td>
			<td width="15">&nbsp;</td>
			<td width="139" class="presenta1">Descripcion del Bien</td>
			<td width="289" class="presenta2"><%response.write(captados("idproducto"))%><%response.write(" - ")%><%response.write(captados("nomarticulo"))%></td>
		</tr>
	</table>
	<table width="100%" height ="2%"><tr></tr></table>
	<table width="837" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ACE5EE">
		<tr bgcolor="<%Response.Write(session("color"))%>">
			<td width="8"><div align="center"><font color="#0000FF" size="2"></td>
			<td width="138" class="presenta1">Punto de Venta</td>
			<td width="468" class="presenta2"><%response.write(captados("nomptoventa"))%></td>
			<td width="124" class="presenta1">Fecha de Contacto</td>
			<td width="99" class="presenta2"><%response.write(captados("contactado"))%></td>
		</tr>
	</table>

	<!-- aqui cambia las fechas de las citas 1,2 o 3 fechas -->
	<table width="542"  border="0" align="center" cellpadding="1" cellspacing="0" >		
		<form method="post" action="<%=MM_editAction%>" name="form2">
			<tr> 
				<td width="24">&nbsp;</td>
				<td width="200"class="presenta1"></td>
			</tr>
			<tr> 
				<td width="24">&nbsp;</td>
				<td width="200"class="presenta1">CITA</td>
				<td width="170"><div align="Center"><font color="#990000" size="2"><b>Fecha y hora de actualizacion</td>
			</tr>
			<tr> 
				<td width="24"><strong><font color="#990000"></strong></td>
				<td width="200"><div>
						<input type="text" name="fecita1" id="fecita1" value="<%response.write(captados.Fields.Item("cita"))%>" /> 
						<img src="imagenes/calendario.png" width="16" height="16" border="0" title="Fecha Cita" id="lanzador1">
						<script type="text/javascript"> 
							Calendar.setup({ 
							inputField:"fecita1",     // id del campo de texto 
							ifFormat  :"%d-%m-%Y",     // formato de la fecha que se escriba en el campo de texto 
							button    :"lanzador1"     // el id del botón que lanzará el calendario 
							}); 
						</script></div>
				</td>
				<td width="150"><div align="center">
					<input type="text" name="fecita2" value="<%response.write(cestado)%>" size="18" readonly> 
				</div></td>
			</tr>
			<table width="901">
				<td> <div align="right"><input type="submit" value="Guardar Fecha"></div></td>
			</table>
			<input type="hidden" name="MM_update" value="form2">
			<input type="hidden" name="MM_recordId" value="<%response.write(captados("cedula"))%>">
		</form>
	</table>
	<!-- aqui muestra las observaciones ingresadas -->
	<table width="884"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
		<tr valign="top">
			<td width="100" class="rotulos">Fecha Cita</strong></td>
			<td width="820" class="rotulos">Observación Cita</strong></td>
			<td width="27" class="rotulos">Eval</strong></td>
		</tr>
	</table>
	<table width="884"  border="1" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
		<%While ((Repeat1__numRows <> 0) AND (NOT obsseguim.EOF))%>
			<tr valign="top">
				<td width="100" align="center" bgcolor="<%if ( Repeat1__numRows mod 2)=0 then Response.Write color end if%>">
					<font size="2" ><%response.write(obsseguim("obs_fecha"))%></font>
				</td>
				<td width="820" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write color end if%>"><font size="2" >
					<%response.write(obsseguim("obs_observa"))%></font>
				</td>
				<td width="27" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write color end if%>"><font size="2" >
					<%response.write(obsseguim("obs_evalua"))%></font>
              </td>
			</tr>
			<% 
			Repeat1__index=Repeat1__index+1
			Repeat1__numRows=Repeat1__numRows-1
			obsseguim.MoveNext()
		Wend
		%>
	</table>
	<!-- aqui esta para ingresar una nueva observación -->
	<table width="884" border="0" align="center">
		<form name="formEvalua" method="POST" action="<%=MM_editAction%>">
			<%fechao = date()%>
			<tr valign="top"> 
				<td width="100"><input type="text" name="fechao" value="<%response.write(fechao)%>" size="10" readonly>
				<td class="detalle"> <textarea name="observa" cols="90" rows="2"></textarea></td>
				<td class="detalle">
					<select NAME="evalua" SIZE="1" >
						<option value="0">0%</option>
						<option value="10">10%</option>
						<option value="20">20%</option>
						<option value="30">30%</option>
						<option value="40">40%</option>
						<option value="50">50%</option>
						<option value="60">60%</option>
						<option value="70">70%</option>
						<option value="80">80%</option>
						<option value="90">90%</option>
					</select>
				</td>
			</tr>
			<table width="901" ><tr></tr>
			</table>
		    <table width="100%"  align="center">
        		<tr>
		            <td width="14%">				</td>
        		    <td width="53%">
						<%if session("codemi") >0 then
							select case session("cargo")
								case "11"%>
									<div align="Center"><input name="suspender" type="button" id="suspender" 
										onClick="MM_goToURL('parent','suspende_cliente.asp?Id=<% Response.Write(captados("cedula"))%>');return document.MM_returnValue" 
									value="Suspender Prospecto"> </div>
								<%case "00"%>
									<div align="Center"><input name="suspender" type="button" id="suspender" 
										onClick="MM_goToURL('parent','suspende_cliente.asp?Id=<% Response.Write(captados("cedula"))%>');return document.MM_returnValue" 
									value="Suspender Referido"> </div>
								<%end select%>
						<%end if%>
                    </td>
					<td> <div align="right"><input type="submit" value="Agregar Observación" ></div></td>
					<td width="33%">
						<input type="hidden" name="MM_insert" value="formEvalua">
						<input type="hidden" name="hiddenField"  value="<%response.write(captados("cedula"))%>"></td>
					</td>
		      </tr>
		  </table>
		</form>
	</table>
	<!-- aqui regresa a la pagina inicial de control de citas -->
	<table width="901"  border="0" align="center" cellpadding="1" cellspacing="0" >		
			<tr> 
				<td width="61">&nbsp;</td>
				<td width="122"class="presenta1"></td>
				<td width="122"class="presenta1"></td>
				<td width="122"class="presenta1"></td>
			</tr>
	</table>
    <table width="100%"  align="center">
        <tr>
            <td width="14%">				</td>
            <td width="53%">
                <FORM NAME="volver" ACTION="control_de_citas.asp">
                    <input type="submit" name="asigna" value="Control de Citas a Clientes">
                </FORM>
			</td>
			<td width="33%">
                <FORM NAME="volver" ACTION="menu.asp">
                    <input type="submit" name="" value="Regresar">
                </FORM>
			</td>
      </tr>
    </table>
</body>
</html>
<%
captados.Close()
Set captados = Nothing
obsseguim.Close()
Set obsseguim = Nothing
%>
