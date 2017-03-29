<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/usuarios.asp" -->
<%
Dim puntos
Dim puntos_numRows
titulo ="INGRESAR O MODIFICAR TABLA DE PUNTOS DE VENTA"
Set puntos = Server.CreateObject("ADODB.Recordset")
puntos.ActiveConnection = MM_usuarios_STRING
puntos.Source = "SeLECT *  FROM dbo.puntoventa order by 1"
puntos.CursorType = 0
puntos.CursorLocation = 2
puntos.LockType = 1
puntos.Open()

puntos_numRows = 0
%>

<%
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = 15
Repeat1__index = 0
puntos_numRows = puntos_numRows + Repeat1__numRows
%>
<%
Dim puntos_total
Dim puntos_first
Dim puntos_last
puntos_total = puntos.RecordCount
If (puntos_numRows < 0) Then
	puntos_numRows = puntos_total
Elseif (puntos_numRows = 0) Then
	puntos_numRows = 1
End If
puntos_first = 1
puntos_last  = puntos_first + puntos_numRows - 1
If (puntos_total <> -1) Then
	If (puntos_first > puntos_total) Then
		puntos_first = puntos_total
	End If
	If (puntos_last > puntos_total) Then
		puntos_last = puntos_total
	End If
	If (puntos_numRows > puntos_total) Then
		puntos_numRows = puntos_total
	End If
End If
%>
<%
If (puntos_total = -1) Then
	puntos_total=0
	While (Not puntos.EOF)
		puntos_total = puntos_total + 1
		puntos.MoveNext
	Wend
	session("registros")=puntos_total
	If (puntos.CursorType > 0) Then
		puntos.MoveFirst
	Else
		puntos.Requery
	End If
	If (puntos_numRows < 0 Or puntos_numRows > puntos_total) Then
		puntos_numRows = puntos_total
	End If
	puntos_first = 1
	puntos_last = puntos_first + puntos_numRows - 1
	If (puntos_first > puntos_total) Then
		puntos_first = puntos_total
	End If
	If (puntos_last > puntos_total) Then
		puntos_last = puntos_total
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
Set MM_rs    = puntos
MM_rsCount   = puntos_total
MM_size      = puntos_numRows
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
puntos_first = MM_offset + 1
puntos_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
	If (puntos_first > MM_rsCount) Then
		puntos_first = MM_rsCount
	End If
	If (puntos_last > MM_rsCount) Then
		puntos_last = MM_rsCount
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
	<title>Modificación de Tablas</title>
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
	<table width="600" border="1" bgcolor="#CCCCCC" align="center">
		<tr> 
			<td width="90" class="rotulos">Codigo</td>
			<td width="450" class="rotulos">Descripcion</td>
		</tr>
	</table>
	<table width="600" border="1" align="center" cellspacing="0">
		<% 
		While ((Repeat1__numRows <> 0) AND (NOT puntos.EOF)) 
			%>
			<tr> 
				<td width="90" class="presentac" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write(session("color")) end if%>">
					<a href="ingptoventa1.asp?<%= MM_keepBoth & MM_joinChar(MM_keepBoth) & "id_puntoventa=" & puntos.Fields.Item("id_puntoventa").Value %>">
					<%=(puntos.Fields.Item("id_puntoventa").Value)%></a></td>
				<td width="450" class="presental" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write(session("color")) end if%>">
					<%=(puntos.Fields.Item("nom_puntoventa").Value)%> </font></td>
			</tr>
			<% 
			Repeat1__index=Repeat1__index+1
			Repeat1__numRows=Repeat1__numRows-1
			puntos.MoveNext()
		Wend
		%>
	</table>
	<table border="0" width="50%" align="center">
		<tr> 
			<td width="23%" align="center"> <% If MM_offset <> 0 Then %>
				<a href="<%=MM_moveFirst%>"><img src="imagenes/primero.gif" width="38" height="26" border=0></a> 
				<% End If ' end MM_offset <> 0 %> </td>
			<td width="31%" align="center"> <% If MM_offset <> 0 Then %>
				<a href="<%=MM_movePrev%>"><img src="imagenes/atras.gif" width="38" height="26" border=0></a> 
				<% End If ' end MM_offset <> 0 %> </td>
			<td width="23%" align="center"> <% If Not MM_atTotal Then %>
				<a href="<%=MM_moveNext%>"><img src="imagenes/adelante.gif" width="38" height="26" border=0></a> 
				<% End If ' end Not MM_atTotal %> </td>
			<td width="23%" align="center"> <% If Not MM_atTotal Then %>
				<a href="<%=MM_moveLast%>"><img src="imagenes/ultimo.gif" width="38" height="26" border=0></a> 
			<% End If ' end Not MM_atTotal %> </td>
		</tr>
	</table>
	<div align="center"><font size="2">Registros <%=(puntos_first)%> a <%=(puntos_last)%> de <%=(puntos_total)%></font> </div>
    <table width="100%"  align="center">
        <tr>
            <td width="14%">				</td>
            <td width="53%">
                <FORM NAME="volver" ACTION="menu.asp">
                    <input type="submit" name="" value="Regresar">
                </FORM>
            </td>
            <td width="33%">
            </td>
      </tr>
    </table>
</body>
</html>
<%
puntos.Close()
Set puntos = Nothing
%>
