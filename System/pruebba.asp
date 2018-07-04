
<html>
<head>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>

<SCRIPT TYPE="text/javascript">
$(document).ready(function(){
	$('#cambio_color').toggle(function(){
		$('#resultado').css('background-color','red')	
	}, function(){
		$('#resultado').css('background-color','blue')	
	})
})

function suma(){
	var SUMA1 =	$('#SUMA1').val()
	var SUMA2 =	$('#SUMA2').val()

	var parametros ="alea="+Math.random()+"&SUMA1="+SUMA1+"&SUMA2="+SUMA2
	$('#resultado').load('pagina_servidor.asp',parametros, function(){})


}


function limpiar(){
	$('#SUMA1').val("")
	$('#SUMA2').val("")	

	$('#resultado').text("")	

}

</SCRIPT>



<style type="text/css" media="screen">
	.color{
		width:200px; 
		height:300px;	
	}

	.color:hover{
		font-size: 50px;
		color:#ccc;
	}
</style>
</head>
<body>
	<input type="text" name="SUMA1" ID="SUMA1" value="">	
	<input type="text" name="SUMA2" ID="SUMA2" value="">

	<input type="button" name="" value="ingresar" onclick="suma()">
	<input type="button" name="" value="limpiar" onclick="limpiar()">
	<input type="button" name="" id="cambio_color" value="color" onclick="">


	<div id="resultado" class="color"></div>
</body>
</html>


