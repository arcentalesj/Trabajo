function myFunction(id,valor,nombre) {
	var x = id.value;
	var resultado = x*valor;
	document.getElementById(nombre).innerHTML = resultado;
	var lim = document.getElementsByName("datos").length;
	var total= 0;
	for (var i = 0; i < lim; i++) {
		var el = document.getElementsByName('datos')[i];
		var currentNumber = parseFloat(el.innerHTML);
		total+= currentNumber
	}
	document.getElementById('total').innerHTML = total;
}

function valcedula() {
	var cuadro = document.getElementById("cedula");
	var cedula = cuadro.value;
	if (!esNumero(formulario.cedula.value)) {
		cuadro.value = "Cedula Incorrecta";
		return false;
		}
	if ((formulario.cedula.value.length > 10)  || (formulario.cedula.value.length < 10)) {
		cuadro.value =("Cedula Incorrecta");
		return false;
		}
	if (!validar(formulario.cedula.value)) {
		document.formulario.cedula.value="";
		cuadro.value =("Cedula Incorrecta");
		return false;
		}
}

function valcodigo() {
	var ccuadro = document.getElementById("codigo");
	var codigo = ccuadro.value;
	if (!esNumero(formulario.codigo.value)) {
		ccuadro.value = "Codigo Errado";
		return false;
		}
	if ((formulario.codigo.value.length > 2)  || (formulario.codigo.value.length < 1)) {
		ccuadro.value =("Codigo errado");
		return false;
		}
}

function valdatos() {
	var cuadro = document.getElementById("cedula");
	var ccuadro = document.getElementById("codigo");
	var cedula = cuadro.value;
	var codigo = ccuadro.value;
	codigo=(formulario.codigo.value)
	if (!esNumero(formulario.cedula.value)) {
		cuadro.value = "Cedula Incorrecta";
		return false;
		}
	if ((formulario.cedula.value.length > 10)  || (formulario.cedula.value.length < 10)) {
		cuadro.value =("Cedula Incorrecta");
		return false;
		}
	if (!validar(formulario.cedula.value)) {
		document.formulario.cedula.value="";
		cuadro.value =("Cedula Incorrecta");
		return false;
		}
	if (formulario.codigo.value=="") {
		codigo.value = "Ingrese el Código";
		return false;
		}
	if (!esNumero(formulario.codigo.value)) {
		codigo.value = "Codigo Incorrecto";
		return false;
		}
	formulario.action = "Buscar_asesor.asp?cedula1="+numero+"&codigo1="+codigo,"_parent";				
	var respuesta=window.open("Buscar_asesor.asp?cedula1="+numero+"&codigo1="+codigo,"_parent");
	return true;
}
	
function esNumero(valor) {
	var bNumeric=true;
	for (var ii = 0 ; ii < valor.length ; ii++) {
		if(0 > ('0123456789').indexOf(valor.substring(ii, ii+1))) {
			bNumeric=false;
			break;
		}
	}
	return bNumeric;
}

function validar(valor) {
	/* funcion para validar la cedula */
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
			return false;
		}                  
		/* El ruc de las empresas del sector publico terminan con 0001*/         
		if ( numero.substr(9,4) != '0001' ){                    
			return false;
		}
	}         
	else if(pri == true){         
		if (digitoVerificador != d10){                          
			return false;
		}         
		if ( numero.substr(10,3) != '001' ){                    
			return false;
		}
	}      
	else if(nat == true){         
		if (digitoVerificador != d10){                          
			return false;
		}         
		if (numero.length >10 && numero.substr(10,3) != '001' ){                    
			return false;
		}
	}      
	return true;   
}

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
	
function MM_swapImage() {
	var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
	if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
	
function MM_openBrWindow(theURL,winName,features) {
	window.open(theURL,winName,features);
}

function MM_goToURL() {
	var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
	for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function armar()
	{
	if (!esNumero(formulario.cedula.value)) {
		alert("El documento de identificacion debe ser numérico.")
		return false;
	}
	if (!esNumero(formulario.mes.value)) {
		alert("El campo del mes debe ser numérico.")
		return false;
	}
	if (!esNumero(formulario.anio.value)) {
		alert("El campo del año debe ser numérico.")
	return false;
	}
	var campo1=(formulario.nombre.value);
	var campo2=(formulario.apellido.value);
	var campo3=(formulario.cedula.value);
	var campo4=(formulario.mes.value);
	var campo5=(formulario.anio.value);
	var campo6=(formulario.vende1.value);
	if ((campo1 == "") && (campo2 == "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0)) { 
		alert('NO ha ingresado datos para procesar la consulta');
		return false; 
	}
	if ((campo1 !== "") && (campo2 == "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0)) { 
		formulario.action = "consultatotal.asp?opcion=1&campo="+campo1,"_parent";
		var respuesta=window.open("consultatotal.asp?opcion=1&campo="+campo1,"_parent");
		return true;
	}
	if ((campo1 == "") && (campo2 !== "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0)) { 
		formulario.action = "consultatotal.asp?opcion=3&campo="+campo2,"_parent";
		var respuesta=window.open("consultatotal.asp?opcion=3&campo="+campo2,"_parent");
		return true;
	}
	if ((campo1 == "") && (campo2 == "") && (campo3 !== "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0)) { 
		formulario.action = "consultatotal.asp?opcion=4&campo="+campo3,"_parent";
		var respuesta=window.open("consultatotal.asp?opcion=4&campo="+campo3,"_parent");
		return true;
	}
	if ((campo1 == "") && (campo2 == "") && (campo3 == "") && (campo4 !== "") &&(campo5 == "") && (campo6 == 0)) { 
		formulario.action = "consultatotal.asp?opcion=6&campo="+campo4,"_parent";
		var respuesta=window.open("consultatotal.asp?opcion=6&campo="+campo4,"_parent");
		return true;
	}
	if ((campo1 == "") && (campo2 == "") && (campo3 == "") && (campo4 == "") &&(campo5 !== "") && (campo6 == 0)) { 
		formulario.action = "consultatotal.asp?opcion=7&campo="+campo5,"_parent";
		var respuesta=window.open("consultatotal.asp?opcion=7&campo="+campo5,"_parent");
		return true;
	}
	if ((campo1 == "") && (campo2 == "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 !== 0)) { 
		formulario.action = "consultatotal.asp?opcion=5&campo="+campo6,"_parent";
		var respuesta=window.open("consultatotal.asp?opcion=5&campo="+campo6,"_parent");
		return true;
	}
	if ((campo1 !== "") && (campo2 !== "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0)) { 
		formulario.action = "consultatotal.asp?opcion=2&campo="+campo1+"&campoa="+campo2,"_parent";
		var respuesta=window.open("consultatotal.asp?opcion=2&campo="+campo1+"&campoa="+campo2,"_parent");
		return true;
	}
}

function autor() 
	{
	document.write("<div style='color:black;font-size:16px;'>");
	document.write('\&copy;\&#32;\Copyright Juan Carlos Arcentales Burgos 2016');
	document.write("<div style='color:black;font-size:16px;'>");
	document.write('CV Proyectos y Aseor&#237;a Acu&#237;cola');
	document.write("</div>");
	}

function fecha() 
	{
	dows = new Array("Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado");
	months = new Array("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre");
	now = new Date();
	dow = now.getDay();
	d = now.getDate();
	m = now.getMonth();
	h = now.getTime();
	y = now.getFullYear();
	document.write(dows[dow]+" "+d+" de "+months[m]+" de "+y);
}
	
function veriFechas() 
	{
	// funcion para validar rango de fechas
	var fechaini = (formulario.fecini.value);
	var fechafin = (formulario.fecfin.value);
	var fecha1 = fechaini.split('/');
	diaini = parseInt(fecha1[0])*1;
	mesini = parseInt(fecha1[1])*30;
	añoini = parseInt(fecha1[2])*360;
	tfechaini= (parseInt(diaini)+parseInt(mesini)+parseInt(añoini));
	var fecha2 = fechafin.split('/');
	diafin = parseInt(fecha2[0])*1;
	mesfin = parseInt(fecha2[1])*30;
	añofin = parseInt(fecha2[2])*360;
	tfechafin= (parseInt(diafin)+parseInt(mesfin)+parseInt(añofin));
	var dias= (parseInt(tfechafin)-parseInt(tfechaini));
	if (dias >=0 ) {
		formulario.action = "prospec_ingresadosXLS.asp?fecini="+fechaini+"&fecfin="+fechafin,"_parent";				
		var respuesta=window.open("prospec_ingresadosXLS.asp?fecini="+fechaini+"&fecfin="+fechafin,"_parent");
		return true;
		}
	else {
		alert('Fecha Final es menor que Fecha Inicial');
		return false; 
		}
}

function veriFechasC() 
	{
	// funcion para validar rango de fechas
	var fechaini = (formulario.fecini.value);
	var fechafin = (formulario.fecfin.value);
	var fecha1 = fechaini.split('/');
	diaini = parseInt(fecha1[0])*1;
	mesini = parseInt(fecha1[1])*30;
	añoini = parseInt(fecha1[2])*360;
	tfechaini= (parseInt(diaini)+parseInt(mesini)+parseInt(añoini));
	var fecha2 = fechafin.split('/');
	diafin = parseInt(fecha2[0])*1;
	mesfin = parseInt(fecha2[1])*30;
	añofin = parseInt(fecha2[2])*360;
	tfechafin= (parseInt(diafin)+parseInt(mesfin)+parseInt(añofin));
	var dias= (parseInt(tfechafin)-parseInt(tfechaini));
	if (dias >=0 ) {
		formulario.action = "citasXLS.asp?fecini="+fechaini+"&fecfin="+fechafin,"_parent";				
		var respuesta=window.open("citasXLS.asp?fecini="+fechaini+"&fecfin="+fechafin,"_parent");
		return true;
		}
	else {
		alert('Fecha Final es menor que Fecha Inicial');
		return false; 
		}
}

function valprefijo1() {
	var cuadro = document.getElementById("codarea1");
	var ccodarea1 = cuadro.value;
	if (!esNumero(formulario.codarea1.value)) {
		cuadro.value = "";
		alert ("No ha seleccionado Provincia");
		return false;
		}
}
function valestado() {
	var cuadro = document.getElementById("ecivil");
	var ccivil = cuadro.value;
	if (!esNumero(formulario.ecivil.value)) {
		cuadro.value = "";
		alert ("No ha seleccionado el Estado");
		return false;
		}
}
function valpunto() {
	var cuadro = document.getElementById("pto1");
	var cpto1 = cuadro.value;
	if (!esNumero(formulario.pto1.value)) {
		cuadro.value = "";
		alert ("No ha seleccionado Punto Venta");
		return false;
		}
}
