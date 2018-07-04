<!DOCTYPE html>
<html>
  <head>
    <meta http-equiv="content-type" content="text/html; charset=utf-8"/>
<%

	Response.CodePage=65001
	Response.charset ="utf-8"
	
	direccion =request.querystring("direccion")
%>
    
    <script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="http://maps.google.com/maps?file=api&v=3&key=AIzaSyA3_fBj86leWUs4h3vBEBGm5AazF36EwdU&sensor=false" type='text/javascript'></script>

<script type='text/javascript'>
	function load() {
		var direccion =$('#direccion').val()+', Chile' 

		if (GBrowserIsCompatible()) {
			var map = new GMap2(document.getElementById('map'));
			var geocoder = new GClientGeocoder();
			map.addControl(new GLargeMapControl());
			map.addControl(new GScaleControl());

		      map.setMapType(G_NORMAL_MAP);
		      map.setUIToDefault();

			if (geocoder) {
				var address = direccion;
				geocoder.getLatLng(
					address,
					function(point) 
{						if (!point) {
							alert(address + ' not found');
						} else {
							map.setCenter(point, 17);
							var marker = new GMarker(point);
							document.getElementById('latitud').value = marker.getLatLng().lat();
							document.getElementById('longitud').value = marker.getLatLng().lng();
							map.addOverlay(marker);
							document.title = 'geo';
						}
					}
				);
			}
		}
	}
</script>
</head>
<body onload='load()' onunload='GUnload()'>
<div style="width:100%;text-align:center;"><%=direccion%></div>
<br>
<div id='map' align='center' style='width:600px; height: 600px'></div>

<input type='hidden' name='latitud' id='latitud'/>
<input type='hidden' name='longitud' id='longitud'/>

<input type='hidden' name='direccion' id='direccion' value="<%=direccion%>"/>

</body>