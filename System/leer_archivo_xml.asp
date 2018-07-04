<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

<script>

function loadXMLDoc(url)
{

var xmlhttp;
var txt,x,xx,i;

if (window.XMLHttpRequest){// code for IE7+, Firefox, Chrome, Opera, Safari
	 xmlhttp=new XMLHttpRequest();

}else{// code for IE6, IE5
	xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
}

xmlhttp.onreadystatechange=function(){

	if (xmlhttp.readyState==4 && xmlhttp.status==200)
	{

		txt="<table border='1'><tr><th>Title</th><th>Artist</th></tr>";

		x=xmlhttp.responseXML.documentElement.getElementsByTagName("CD");

		for (i=0;i<x.length;i++)
		{
			txt=txt + "<tr>";
			xx=x[i].getElementsByTagName("subject");
			{
				try
				{
					txt=txt + "<td width='200'>" + xx[0].firstChild.nodeValue + "</td>";
				}
				catch (er)
				{
					txt=txt + "<td> </td>";
				}
			}

			xx=x[i].getElementsByTagName("extract");

			{
				try
				{
					txt=txt + "<td>" + xx[0].firstChild.nodeValue + "</td>";
				}
				catch (er)
				{
					txt=txt + "<td> </td>";
				}
			}

			txt=txt + "</tr>";

		}

		txt=txt + "</table>";

		document.getElementById('txtCDInfo').innerHTML=txt;

	}

}

xmlhttp.open("GET",url,true);

xmlhttp.send();

}

</script>

</head>

<body>



<div id="txtCDInfo">

<button onclick="loadXMLDoc('prueba.xml')">Obtener info del CD</button>

</div>

<%

%>

</body>

</html>

<%
Response.Write ("<html><head><title>Ejemplo fichero XML (RSS)</title>")
Response.Write ("</head><body>")
Dim objHTTP
Dim url
url= "http://www.forosdelweb.com/index.xml"
Set objHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
objHTTP.Open "POST", url, false
objHTTP.send()

Response.Write ("<h2>" & objHTTP.responseXml.SelectSingleNode("rss/channel/title").Text & "</h2>")
Response.Write ("<h3>" &  objHTTP.responseXml.SelectSingleNode("rss/channel/description").Text & "</h3>")
For Each objItem in objHTTP.responseXML.SelectNodes("rss/channel/item")
Response.Write ("<p>")
 Response.Write ("<h5>" & objItem.SelectSingleNode("title").text & "</h5>")
 Response.Write (objItem.SelectSingleNode("description").text & "<br />")
 Response.Write ("<a href=""" & objItem.SelectSingleNode("link").Text & """>más...</a>")
 Response.Write ("</p>")
Next
Set objHTTP = Nothing
Response.Write ("</body></html>")

%>