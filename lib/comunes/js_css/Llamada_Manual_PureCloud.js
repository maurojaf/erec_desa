        
		
		var userIDGlob;

        // Funcion que separa los parametros de la URL cuando se envia el Token
        function getParameterByName(name) {
            name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
            var regex = new RegExp("[\\#&]" +
                name + "=([^&#]*)"),
                results = regex.exec(location.hash);
            return results === null ? "" : decodeURIComponent(results[1].replace(/\+/g,
                " "));
        }

        // Reconoce si en la URL de la web trae el Access_token
        if (window.location.hash) {
            console.log(location.hash);
            var token = getParameterByName('access_token');
            //location.hash = ''
        } else {
            var queryStringData = {
                response_type: "token",
                client_id: "92014f51-547f-4f04-af6d-0649e54ba879", //Client ID perteneciente a Llacruz, no modificar
                redirect_uri: "http://localhost/DiscadoManual/index.html" // Aca se tiene que poner la pagina donde van a estar los numeros telefonicos, para hacer el redirect y capturar el Token, que es necesario para el consumo de la API que realiza la llamada
            }
            console.log(queryStringData);
            console.log(jQuery.param(queryStringData));
            window.location.replace("https://login.mypurecloud.com/oauth/authorize?" + jQuery.param(queryStringData));
        }

        //Metodo que realiza la llamada al momento de llamar a la API

        function getUserID() {
            var UserID;
            var settings = {
                "async": true,
                "crossDomain": true,
                "url": "https://api.mypurecloud.com/api/v2/users/me",
                "method": "GET",
                "headers": {
                    "Authorization": "Bearer " + token,
                    "Cache-Control": "no-cache"
                }
            }

            $.ajax(settings).done(function (response) {
                UserID = response.id;
                userIDGlob = UserID;
                return UserID;
            });

        }

        getUserID();

        function llamar() {
            console.log(userIDGlob);

            var button = $('#Message').val(); //Trae el contenido del Input que esta en este ejemplo, se puede poner un <a href="tel:"></a> donde se obtenga el numero

            var settings = {
                "async": true,
                "crossDomain": true,
                "url": "https://api.mypurecloud.com/api/v2/conversations/callbacks",
                "method": "POST",
                "headers": {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + token,
                    "Cache-Control": "no-cache"
                },
                "processData": false,
                "data": "{\n   \"scriptId\": \"\",\n   \"queueId\": \"\",\n   \"routingData\": {\n      \"queueId\": \"50fec37f-a4ff-45c8-a690-520f4a6c7a48\",\n      \"languageId\": \"\",\n      \"skillIds\": [],\n      \"preferredAgentIds\": [\"" + userIDGlob + "\"]\n   },\n   \"callbackUserName\": \"\",\n   \"callbackNumbers\": [" + button + "],\n   \"callbackScheduledTime\": \"\",\n   \"countryCode\": \"\",\n   \"data\": {}\n}"
            }

            $.ajax(settings).done(function (response) {
                console.log(response);
            });
        }