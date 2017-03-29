<%@LANGUAGE="VBSCRIPT"%>
<!-- #include file="connections/usuarios.asp" -->
<% 
   Response.addHeader "pragma", "no-cache"
   Response.CacheControl = "Private"
   Response.Expires = 0
%>

<%
'response.write(session("accion"))
If (Request.QueryString("opcion") <> "") Then 
  session("opcion") = trim(Request.QueryString("opcion"))
End If
If (Request.QueryString("campo") <> "") Then 
  session("campo") = trim(Request.QueryString("campo"))
End If
If (Request.QueryString("campoa") <> "") Then 
  session("campoa") = trim(Request.QueryString("campoa"))
End If
if session("cargo") ="71" then
	select case session("opcion")
		case 1
			strSQL = "SELECT * FROM dbo.captados where nombre like '%"+rtrim(session("campo"))+"%' and idvendedor = '"+  session("USUARIO")+"'"
			titulo = "BUSQUEDA DE CLIENTES POR NOMBRE "+ucase(session("campo"))
		case 2
			strSQL = "SELECT * FROM dbo.captados where nombre like '%"+rtrim(session("campo"))
			strSQL = strSQL + "%' and apellido like '"+rtrim(session("campoa"))+"%'  and idvendedor = '"+  session("USUARIO")+"'"
			titulo = "BUSQUEDA DE CLIENTES POR NOMBRE "+ucase(session("campo"))+" Y APELLIDO "+ucase(session("campoA"))
		case 3
			strSQL = "SELECT * FROM dbo.captados where apellido like '"+rtrim(session("campo"))+"%' and idvendedor = '"+  session("USUARIO")+"'"
			titulo = "BUSQUEDA DE CLIENTES POR APELLIDO "+ucase(session("campo"))
		case 4
			strSQL = "SELECT * FROM dbo.captados where cedula = '"+rtrim(session("campo"))+"' and idvendedor = '"+  session("USUARIO")+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL NUMERO DE CEDULA  "+(session("campo"))
		case 5
			strSQL = "SELECT * FROM dbo.captados where idvendedor = '"+rtrim(session("campo"))+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL CÓDIGO DE VENDEDOR  "+(session("campo"))
		case 6
			strSQL = "SELECT * FROM dbo.captados where MONTH(ingresado) = '"+(session("campo"))+"' and idvendedor = '"+  session("USUARIO")+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL MES DE CAPTACION "+(session("campo"))
		case 7
			strSQL = "SELECT * FROM dbo.captados where YEAR(ingresado) = '"+(session("campo"))+"' and idvendedor = '"+  session("USUARIO")+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL AÑO DE CAPTACION "+(session("campo"))
		case 8
			strSQL = "SELECT * FROM dbo.captados_todos where cotizador = '"+rtrim(session("campo"))+"' and emisor = '"+ session("codemi")+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL NUMERO DE COTIZADOR  "+(session("campo"))
		case 9
			strSQL = "SELECT * FROM dbo.captados_todos where MONTH(fecaptacion)='"+(session("campo"))+"' and YEAR(fecaptacion)='"
			strSQL = strsql + (session("campoa"))+"' and emisor = '"+ session("codemi")+"'"
			titulo = "BUSQUEDA DE CLIENTES POR MES "+(session("campo"))+" Y AÑO "+(session("campoA"))
	end select
else
	' este else es cuando son los usuarios especificos
	select case session("opcion")
		case 1
			strSQL = "SELECT * FROM dbo.captados where nombre like '%"+rtrim(session("campo"))+"%'"
			titulo = "BUSQUEDA DE CLIENTES POR NOMBRE "+ucase(session("campo"))
		case 2
			strSQL = "SELECT * FROM dbo.captados where nombre like '%"+rtrim(session("campo"))+"%' and apellido like '"+rtrim(session("campoa"))+"%'"
			titulo = "BUSQUEDA DE CLIENTES POR NOMBRE "+ucase(session("campo"))+" Y APELLIDO "+ucase(session("campoA"))
		case 3
			strSQL = "SELECT * FROM dbo.captados where apellido like '"+rtrim(session("campo"))+"%'"
			titulo = "BUSQUEDA DE CLIENTES POR APELLIDO "+ucase(session("campo"))
		case 4
			strSQL = "SELECT * FROM dbo.captados where cedula = '"+rtrim(session("campo"))+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL NUMERO DE CEDULA  "+(session("campo"))
		case 5
			strSQL = "SELECT * FROM dbo.captados where idvendedor = '"+rtrim(session("campo"))+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL CÓDIGO DE VENDEDOR  "+(session("campo"))
		case 6
			strSQL = "SELECT * FROM dbo.captados where MONTH(ingresado) = '"+(session("campo"))+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL MES DE CAPTACION "+(session("campo"))
		case 7
			strSQL = "SELECT * FROM dbo.captados where YEAR(ingresado) = '"+(session("campo"))+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL AÑO DE CAPTACION "+(session("campo"))
		case 8
			strSQL = "SELECT * FROM dbo.captados where cotizador = '"+rtrim(session("campo"))+"'"
			titulo = "BUSQUEDA DE CLIENTES POR EL NUMERO DE COTIZADOR  "+(session("campo"))
		case 9
			strSQL = "SELECT * FROM dbo.captados where MONTH(fecaptacion)='"+(session("campo"))+"' and YEAR(fecaptacion)='"+(session("campoa"))+"'"
			titulo = "BUSQUEDA DE CLIENTES POR MES "+(session("campo"))+" Y AÑO "+(session("campoA"))
	end select
end if
%>
<%
Dim captados
Dim captados_numRows,opcion
Dim captados__MMColParam
captados__MMColParam = "1"
Set captados = Server.CreateObject("ADODB.Recordset")
captados.ActiveConnection = MM_usuarios_STRING
captados.Source = strSQL
captados.CursorType = 0
captados.CursorLocation = 2
captados.LockType = 3
captados.Open()
if captados.bof and captados.bof then
	response.redirect("busqueda.asp?mensajeerror=" &"cclientet_nohay")		
end if
captados_numRows = 0
Dim Repeat1__numRows
Dim Repeat1__index
session("frase") = strSQL
session("titulo") = titulo
Repeat1__numRows = 17
Repeat1__index = 0
captados_numRows = captados_numRows + Repeat1__numRows
%>
<%
Dim captados_total
Dim captados_first
Dim captados_last
captados_total = captados.RecordCount
If (captados_numRows < 0) Then
	captados_numRows = captados_total
Elseif (captados_numRows = 0) Then
	captados_numRows = 1
End If
captados_first = 1
captados_last  = captados_first + captados_numRows - 1
If (captados_total <> -1) Then
	If (captados_first > captados_total) Then
		captados_first = captados_total
	End If
	If (captados_last > captados_total) Then
		captados_last = captados_total
	End If
	If (captados_numRows > captados_total) Then
		captados_numRows = captados_total
	End If
End If
%>
<%
If (captados_total = -1) Then
	captados_total=0
	While (Not captados.EOF)
		captados_total = captados_total + 1
		captados.MoveNext
	Wend
	If (captados.CursorType > 0) Then
		captados.MoveFirst
	Else
		captados.Requery
	End If
	If (captados_numRows < 0 Or captados_numRows > captados_total) Then
		captados_numRows = captados_total
	End If
	captados_first = 1
	captados_last = captados_first + captados_numRows - 1
	If (captados_first > captados_total) Then
		captados_first = captados_total
	End If
	If (captados_last > captados_total) Then
		captados_last = captados_total
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
Set MM_rs    = captados
MM_rsCount   = captados_total
MM_size      = captados_numRows
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
captados_first = MM_offset + 1
captados_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
	If (captados_first > MM_rsCount) Then
		captados_first = MM_rsCount
	End If
	If (captados_last > MM_rsCount) Then
		captados_last = MM_rsCount
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
	<title>Consultas de Clientes</title>
	<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
	<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
	<script language="JavaScript" type="text/JavaScript">
		function MM_preloadImages() {
			var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
			var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
			if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
		}
		
		function MM_swapImgRestore() {
			var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
		}
		
		function MM_findObj(n, d) {
			var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
			d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
			if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
			for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
			if(!x && d.getElementById) x=d.getElementById(n); return x;
		}
		
		function MM_swapImage() { //v3.0
			var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
			if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
			}
		
		function MM_openBrWindow(theURL,winName,features) { //v2.0
			window.open(theURL,winName,features);
		}
	</script>
	<script>
	<!--
	function doit(){
		if (!window.print){
		alert("You need NS4.x to use this print button!")
		return
		}
		window.print()
	}
	//-->
	</script>
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
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr> 
			<td class="titulo"><%response.write(titulo)%></td>
		</tr>
		<tr> 
			<td class="subtitulo"><%response.write(titulo1)%></td>
		</tr>
	</table>
	<table width="100%" height="3%" border="0" >
		<tr></tr>
	</table>
	<table width="90%" border="1" bgcolor="#CCCCCC" align="center">
		<tr> 
			<td width="10%" class="rotulos">Cedula</td>
			<td width="20%" class="rotulos">Nombre</td>
			<%if session("cargo")<>"71" then %>
				<td width="20%" class="rotulos">Asesor  Comercial</td>
			<%end if%>
			<td width="20%" class="rotulos">Articulo</td>
			<td width="12%" class="rotulos">Registrado</td>
			<td width="20%" class="rotulos">Lugar Ventas</td>
		</tr>
	</table>
	<table width="90%" border="1" align="center" cellspacing="0">
		<tr> 
			<% While ((Repeat1__numRows <> 0) AND (NOT captados.EOF))
				nombrecli = trim(captados.Fields.Item("apellido").Value)+" "+(captados.Fields.Item("nombre").Value)
'				bien= (captados.Fields.Item("idproducto").Value)+" -- "+captados.Fields.Item("modelo").Value
				select case session("accion")
					case "Busqueda de Prospectos"%>
						<td width="10%" class="presentac" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write(session("color")) end if%>">
			            <a href="muestratodo.asp?<%= MM_keepBoth & MM_joinChar(MM_keepBoth)&"cedula="&(captados.Fields.Item("cedula").Value)%>">
						<%=(captados.Fields.Item("cedula").Value)%></a></font></div></td>
					<%case "Suspender Prospecto" %>
						<td width="10%" class="presentac" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write(session("color")) end if%>">
						<a href="suspende_cliente.asp?<%= MM_keepBoth & MM_joinChar(MM_keepBoth) & "cedula1=" & captados.Fields.Item("cedula").Value %>">
						<%=(captados.Fields.Item("cedula").Value)%></a> </font></div></td>
				<%end select%>
			<td width="20%" class="presental" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write(session("color")) end if%>">
			<%=(response.write(nombrecli))%></td>
			<%if session("cargo")<>"71" then %>
				<td width="20%" class="presental" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write(session("color")) end if%>">
				<%=(captados.Fields.Item("nombrevendedor").Value)%></td>
			<%end if%>
				
			<td width="20%" class="presental" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write(session("color")) end if%>">
			<%=(captados.Fields.Item("producto").Value)%></td>
			<td width="12%" class="presentar" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write(session("color")) end if%>">
			<%=(captados.Fields.Item("Ingresado").Value)%></td>
			<td width="20%" class="presental" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write(session("color")) end if%>">
			<%=(captados.Fields.Item("nomptoventa").Value)%></td>
		</tr>
				<%Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				captados.MoveNext()
			Wend%>
	</table>
	<table border="0" width="50%" align="center">
		<tr> 
			<td width="23%" align="center"> <% If MM_offset <> 0 Then %>
				<a href="<%=MM_moveFirst%>"><img src="imagenes/primero.gif" alt="Inicio" width="38" height="26" border=0></a> 
				<% End If  %> </td>
			<td width="31%" align="center"> <% If MM_offset <> 0 Then %>
				<a href="<%=MM_movePrev%>"><img src="imagenes/atras.gif" alt="Anterior" width="38" height="26" border=0></a> 
				<% End If  %> </td>
			<td width="23%" align="center"> <% If Not MM_atTotal Then %>
				<a href="<%=MM_moveNext%>"><img src="imagenes/adelante.gif" alt="Siguiente" width="38" height="26" border=0></a> 
				<% End If  %> </td>
			<td width="23%" align="center"> <% If Not MM_atTotal Then %>
				<a href="<%=MM_moveLast%>"><img src="imagenes/ultimo.gif" alt="Ultimo" width="38" height="26" border=0></a> 
				<% End If  %> </td>
		</tr>
	</table>
	<td align="center"><div align="center">
    	<strong><font size="2">Registros <%=(captados_first)%> al <%=(captados_last)%> de <%=(captados_total)%></font></strong></div></td>
	<table width="100%"  align="center">
		<tr> 
			<td width="26%" align="center">
				<%select case session("accion")
					case "Busqueda de Prospectos"%>
				        <FORM NAME="actualizar" ACTION="busqueda.asp">
							  <input type="submit" name="asigna" value="Busqueda de Prospectos">
				        </FORM>
					<%case "Suspender Prospecto" %>
					    <FORM NAME="volver" ACTION="busqueda.asp">
							  <input type="submit" name="asigna" value="Suspender Prospecto">
					    </FORM>
	          <%end select%>
			<td width="26%"> 
			</td>
	  </tr>
	</table>
</body>
</html>
