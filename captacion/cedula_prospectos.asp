<%@ LANGUAGE='VBScript'%>
<!--#include file="Connections/usuarios.asp" -->
<%
titulo="PROSPECTOS"
dim mensajeerror
if request.querystring("mensajeerror") <> "" then
	session("dd")=request.querystring("mensajeerror")
end if
Set cierres = CreateObject("ADODB.Recordset")
cierres.ActiveConnection = MM_usuarios_STRING
hoy = date()
dia = day(hoy)
anio = year(hoy)
mes = month(hoy)
hoy = anio&"/"&mes&"/"&dia
if mes = 1 then
	anio= anio-1
	mescierre=12
else
	mescierre= mes-1
end if
equipos=0
for i = 1 to len(trim(session("equipo")))
	if mid(session("equipo"),i,1)="," then
		equipos=equipos+1
	end if
next
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
	<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
	<title>Validaci&oacute;n de documentos</title>
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
    <table width="679" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
            <td align="Left"><font color="#000000" size="3"><b>PRODUCTO</b></font></td>
            <td align="Left"><font color="#000000" size="3"><b>TIPO DE CLIENTE</b></font></td>
            <td align="Left"><font color="#000000" size="3"><b>TIPO DE DOCUMENTO</b></font></td>
        </tr>
    </table>
	<table width="100%" height ="2%"><tr></tr></table>
	<table height="104" align="center" id="anchoDoc">
		<tr>
			<td width="674" class="general">
				<script language='JavaScript' type="text/javascript">
					function init() {
						document.formulario.opcion[0].checked=true;
						document.formulario.texto.value="";
					}
					function nextPage() {
						if (formulario.texto.value.length < 4 || formulario.texto.value=="") {
							alert('Debe ingresar un dato válido.');
							return false;
						}
						if (!(formulario.opcion1[0].checked) && !(formulario.opcion1[1].checked) && !(formulario.opcion1[2].checked) && !(formulario.opcion1[3].checked) && !(formulario.opcion1[4].checked) && !(formulario.opcion1[5].checked))  {
							alert("Debe escoger el tipo de bien.");
							return false;
						}
						if (!(formulario.opcion2[0].checked) && !(formulario.opcion2[1].checked) && !(formulario.opcion2[2].checked))  {
							alert("Debe escoger tipo de Persona");
							return false;
						}
						if (formulario.opcion1[0].checked){
							tipobien = 'AL';
						}
						if (formulario.opcion1[1].checked){
							tipobien = 'TV';
						}
						if (formulario.opcion1[2].checked){
							tipobien = 'FH';
						}
						if (formulario.opcion1[3].checked){
							tipobien = 'FS';
						}
						if (formulario.opcion1[4].checked){
							tipobien = 'MA';
						}
						if (formulario.opcion1[5].checked){
							tipobien = 'AA';
						}
						if (formulario.opcion2[0].checked){
							tipoper = 'N';
						}
						if (formulario.opcion2[1].checked){
							tipoper = 'J';
						}
						if (formulario.opcion2[2].checked){
							tipoper = 'E';
						}
						if ((formulario.opcion2[1].checked) && (formulario.opcion[0].checked)){
							alert("Debe escoger un RUC para persona Jurídica.")
							return false;
						}
						if ((formulario.opcion2[1].checked) && (formulario.opcion[2].checked)){
							alert("Debe escoger un RUC para persona Jurídica.")
							return false;
						}
						if ((formulario.opcion2[2].checked) && (formulario.opcion[1].checked)){
							alert("Debe escoger un Pasaporte para extrajeros.")
							return false;
						}
						if ((formulario.opcion2[2].checked) && (formulario.opcion[0].checked)){
							alert("Debe escoger un Pasaporte para extrajeros.")
							return false;
						}
						if ((formulario.opcion2[0].checked) && (formulario.opcion[2].checked)){
							alert("Debe escoger un Cedula O RUC para persona Natural.")
							return false;
						}
						if (formulario.opcion[0].checked) {
							if (!isNumeric(formulario.texto.value)) {
								alert("La Cedula debe ser numérica.")
								return false;
							}
							if ((formulario.texto.value.length > 10)  || (formulario.texto.value.length < 10)) {
								alert("El Cedula debe tener 10 dígitos.");
								return false;
							}
							if (!validar(formulario.texto.value)) {
								document.formulario.texto.value="";
								return false;
							}
							tipodoc='C';
							formulario.action = "Buscar_prospectos.asp?cedula1="+numero+"&docum="+tipodoc+"&bien="+tipobien+"&perso="+tipoper,"_parent";				
							var respuesta=window.open("Buscar_prospectos.asp?cedula1="+numero+"&docum="+tipodoc+"&bien="+tipobien+"&perso="+tipoper,"_parent");
							return true;
						}
						// opcion 2 de chek list
						if (formulario.opcion[1].checked) {
							if (!isNumeric(formulario.texto.value)) {
								alert("El RUC debe ser numérico.")
								return false;
							}
							if ((formulario.texto.value.length > 13)  || (formulario.texto.value.length < 13)) {
								alert("El RUC debe ser de 13 dígitos.");
								return false;
							}
							if (!validar(formulario.texto.value)) {
								document.formulario.texto.value="";
								return false;
							}
							tipodoc='R';
							formulario.action = "Buscar_prospectos.asp?cedula1="+numero+"&docum="+tipodoc+"&bien="+tipobien+"&perso="+tipoper,"_parent";				
							var respuesta=window.open("Buscar_prospectos.asp?cedula1="+numero+"&docum="+tipodoc+"&bien="+tipobien+"&perso="+tipoper,"_parent");
							return true;
						}
						// opcion 3 de chek list
						if (formulario.opcion[2].checked) {
						//	if (!isNumeric(formulario.texto.value)) {
						//		alert("El Pasaporte debe ser numérico.")
						//		return false;
						//	}
							if ((formulario.texto.value.length > 10)) {
								alert("El Pasaporte debe contener maximo 10 dígitos.");
								return false;
							}
							tipodoc='P';
							numero = (formulario.texto.value);
							formulario.action = "Buscar_prospectos.asp?cedula1="+numero+"&docum="+tipodoc+"&bien="+tipobien+"&perso="+tipoper,"_parent";				
							var respuesta=window.open("Buscar_prospectos.asp?cedula1="+numero+"&docum="+tipodoc+"&bien="+tipobien+"&perso="+tipoper,"_parent");
							return true;
						}
					}
					// funcion para verificar si es numérico
					function isNumeric(valor) {
						var bNumeric=true;
						for (var ii = 0 ; ii < valor.length ; ii++) {
							if(0 > ('0123456789').indexOf(valor.substring(ii, ii+1))) {
								bNumeric=false;
								break;
							}
						}
						return bNumeric;
					}
					// funcion para validar la cedula
					function validar(valor) {
						numero = (valor);
						var suma = 0;      
						var residuo = 0;      
						var pri = false;      
						var pub = false;            
						var nat = false;      
						var numeroProvincias = 25;                  
						var modulo = 11;
						/* Verifico que el campo no contenga letras */                  
						var ok=1;
						for (i=0; i<numero.length && ok==1 ; i++){
							var n = parseInt(numero.charAt(i));
							if (isNaN(n)) ok=0;
						}
						/* Los primeros dos digitos corresponden al codigo de la provincia */
						provincia = numero.substr(0,2);      
						if (provincia < 1 || provincia > numeroProvincias){           
							alert('El código de la provincia (dos primeros dígitos) es inválido');
							return false;       
						}
						/* Aqui almacenamos los digitos de la cedula en variables. */
						d1  = numero.substr(0,1);         
						d2  = numero.substr(1,1);         
						d3  = numero.substr(2,1);         
						d4  = numero.substr(3,1);         
						d5  = numero.substr(4,1);         
						d6  = numero.substr(5,1);         
						d7  = numero.substr(6,1);         
						d8  = numero.substr(7,1);         
						d9  = numero.substr(8,1);         
						d10 = numero.substr(9,1);                
						/* El tercer digito es: */                           
						/* 9 para sociedades privadas y extranjeros   */         
						/* 6 para sociedades publicas */         
						/* menor que 6 (0,1,2,3,4,5) para personas naturales */ 
						if (d3==7 || d3==8){           
							alert('El tercer dígito ingresado es inválido');                     
							return false;
						}         
						/* Solo para personas naturales (modulo 10) */         
						if (d3 < 6){           
							nat = true;            
							p1 = d1 * 2;  if (p1 >= 10) p1 -= 9;
							p2 = d2 * 1;  if (p2 >= 10) p2 -= 9;
							p3 = d3 * 2;  if (p3 >= 10) p3 -= 9;
							p4 = d4 * 1;  if (p4 >= 10) p4 -= 9;
							p5 = d5 * 2;  if (p5 >= 10) p5 -= 9;
							p6 = d6 * 1;  if (p6 >= 10) p6 -= 9; 
							p7 = d7 * 2;  if (p7 >= 10) p7 -= 9;
							p8 = d8 * 1;  if (p8 >= 10) p8 -= 9;
							p9 = d9 * 2;  if (p9 >= 10) p9 -= 9;             
							modulo = 10;
						}         
						/* Solo para sociedades publicas (modulo 11) */                  
						/* Aqui el digito verficador esta en la posicion 9, en las otras 2 en la pos. 10 */
						else if(d3 == 6){           
							pub = true;             
							p1 = d1 * 3;
							p2 = d2 * 2;
							p3 = d3 * 7;
							p4 = d4 * 6;
							p5 = d5 * 5;
							p6 = d6 * 4;
							p7 = d7 * 3;
							p8 = d8 * 2;            
							p9 = 0;            
						}         
						/* Solo para entidades privadas (modulo 11) */         
						else if(d3 == 9) {           
							pri = true;                                   
							p1 = d1 * 4;
							p2 = d2 * 3;
							p3 = d3 * 2;
							p4 = d4 * 7;
							p5 = d5 * 6;
							p6 = d6 * 5;
							p7 = d7 * 4;
							p8 = d8 * 3;
							p9 = d9 * 2;            
						}
						suma = p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9;                
						residuo = suma % modulo;                                         
						/* Si residuo=0, dig.ver.=0, caso contrario 10 - residuo*/
						digitoVerificador = residuo==0 ? 0: modulo - residuo;                
						/* ahora comparamos el elemento de la posicion 10 con el dig. ver.*/                         
						if (pub==true){           
							if (digitoVerificador != d9){                          
								alert('El ruc de la empresa del sector público es incorrecto.');            
								return false;
							}                  
							/* El ruc de las empresas del sector publico terminan con 0001*/         
							if ( numero.substr(9,4) != '0001' ){                    
								alert('El ruc de la empresa del sector público debe terminar con 0001');
								return false;
							}
						}         
						else if(pri == true){         
							if (digitoVerificador != d10){                          
								alert('El ruc de la empresa del sector privado es incorrecto.');
								return false;
							}         
							if ( numero.substr(10,3) != '001' ){                    
								alert('El ruc de la empresa del sector privado debe terminar con 001');
								return false;
							}
						}      
						else if(nat == true){         
							if (digitoVerificador != d10){                          
								alert('El número de cédula de la persona natural es incorrecto.');
								return false;
							}         
							if (numero.length >10 && numero.substr(10,3) != '001' ){                    
								alert('El ruc de la persona natural debe terminar con 001');
								return false;
							}
						}      
						return true;   
					}
				</script>
			  <form name="formulario" method=post onSubmit='return nextPage();' >
					<input type=hidden name="pagina" value="resultado"/>
					  <table width="104%" border="0" align="center" cellpadding="0" cellspacing="0">
							<tr>
								<td width="33%"><input type="radio" name="opcion1" value="1" >Alevines</td>
								<td width="33%"><input type="radio" name="opcion2" value="1" >Persona Natural</td>
								<td width="33%"><input type="radio" name="opcion" value="1" checked="checked"/>C&eacute;dula</td>
							</tr>
							<tr>
								<td width="33%"><input type="radio" name="opcion1" value="2">Trucha viva</td>
								<td width="33%"><input type="radio" name="opcion2" value="1" >Persona Juridica</td>
								<td width="33%"><input type="radio" name="opcion" value="2"/>RUC</td>
							</tr>
							<tr>
	  					        <td width="33%"><input type="radio" name="opcion1" value="3">Filete con Hueso</td>
								<td width="33%"><input type="radio" name="opcion2" value="1" >Extranjero</td>
								<td width="33%"><input type="radio" name="opcion" value="3"/>Pasaporte</td>
							</tr>
							<tr>
	  					        <td width="33%"><input type="radio" name="opcion1" value="4">Filete Sin Hueso</td>
	  					        <td width="33%" align="center">&nbsp;</td>
	  					        <td width="33%" align="center">&nbsp;</td>
							</tr>
							<tr>
	  					        <td width="33%"><input type="radio" name="opcion1" value="5">Filete Mariposa</td>
	  					        <td width="33%" align="center">&nbsp;</td>
	  					        <td width="33%" align="center">&nbsp;</td>
							</tr>
							<tr>
	  					        <td width="33%"><input type="radio" name="opcion1" value="6">Asesor&iacute;a Acuacultura</td>
	  					        <td width="33%" align="center">&nbsp;</td>
	  					        <td width="33%" align="center">&nbsp;</td>
							</tr>
							<tr>
	  					        <td width="33%" align="center">&nbsp;</td>
	  					        <td width="33%" align="center">&nbsp;</td>
								<td width="33%"><input type="text" maxlength='30' size='20' name="texto" value=""></td>
							</tr>
							<tr>
	  					        <td width="33%" align="center">&nbsp;</td>
					  	        <td width="33%" align="center">&nbsp;</td>
								<td> <input type="button" class="boton" name="btnBuscar" value="Buscar" onclick='javascript:nextPage();'  /></td>
							</tr>
					</table>
				</form>
			</td>
		</tr>
	</table>
    <p>
    <table width="75%" height="24" align="center">
        <%
        select case request.querystring("mensajeerror")
            case "cliente_problema"%>
                <tr><td> <div align="Center"><%response.write ("<b>Cliente con algún problema</b>")%></div></td></tr>
                <tr><td> <div align="Center"><%response.write ("<b>Comuniquese con Sistemas</b>")%></div></td></tr>
			<%case "usuario_noes"%>
                <tr><td><div align="Center"><%response.write ("<b>Usted no esta autorizado para ingresarlo</b>")%></div></td></tr>
			<%case "cliente_nohay"%>
                <tr><td> <div align="Center"><%response.write ("<b>Cliente no esta ingresada en la base de datos</b>")%></div></td></tr>
			<%case "cliente_nosuyo"%>
                <tr><td> <div align="Center"><%response.write ("<b>Prospecto de otro vendedor, no puede ingresar</b>")%></div></td></tr>
                <tr><td><div align="Center"></div></td></tr>
			<%case "cliente_problemas"%>
                <tr><td> <div align="Center"><%response.write ("<b>Existe un problema en la base de datos</b>")%></div></td></tr>
                <tr><td><div align="Center"></div></td></tr>
			<%case "vende_inact"%>
                <tr><td> <div align="Center"><%response.write ("<b>Este Prospecto esta ingresado, pero el vendedor</b>")%></div></td></tr>
                <tr><td><div align="Center"><%response.write ("<b>ha salido de la empresa, pida que se le asigne</b>")%></div></td></tr>
			<%case "Cliente_Prohibido"%>
                <tr><td> <div align="Center"><%response.write ("<b>Este prospecto ha tenido problemas</b>")%></div></td></tr>
                <tr><td><div align="Center"><%response.write ("<b>No puede ingresar a nuestra empresa</b>")%></div></td></tr>
		<%end select
        if request.querystring("mensajeerror") <> "" then 
            %>
            <tr><td><div align="Center"></div></td></tr>
            <%
        end if			
        %>
    </table>
    <table width="100%"  align="center">
        <tr>
            <td width="14%"></td>
            <td width="53%">
                <FORM NAME="volver" ACTION="menu.asp">
					<input type="submit" name="" value="Regresar">
				</FORM>
			</td>
			<td width="33%"></td>
      </tr>
    </table>
</body>
</html>
