/// <reference path="../App.js" />

//global variables
window.thisapp = {};

//Please replace this key with your own key for your instance of the Bing Search API. use the key as given on the website, then base 64 encode it.
thisapp.azureServiceKey = "TWFjaGluZUxlYXJuaW5nVGV4dEFuYWx5dGljc1NlcnZpY2VTZW50aW1lbnRBbmFseXNpczp2QUsxMkxkWWNaN21QQnFlQ3VMVFZrQVBIQVZ6MGw1ZkpXbFZpbTVUSEJzPSA=";

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    getSearchResults(result.value);

                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function getSearchResults(text) {
        if (text != "") {
            //https://api.datamarket.azure.com/Bing/Search/v1/Web?Query=%27xbox%27
            //CuErloyPpcm/l83OA+D/ALP0wRJLPRVx24ioEcIOx3s=

            var authorization = "Basic " + thisapp.azureServiceKey;
            var accept = "application/json";
            var apiUrl = "https://api.datamarket.azure.com/Bing/Search/v1/Web?Query=%27" + text +"%27";
            $.support.cors = true;
            $.ajax({
                beforeSend: function (xhr) {
                    xhr.setRequestHeader("Authorization", authorization);
                    xhr.setRequestHeader("ACCEPT", accept);
                },
                url: apiUrl,
                method: 'GET',
                dataType: 'json',
                complete: function (response) {
                    var data = JSON.parse(response.responseText);

                    var resultsDiv = document.getElementById('results'); 

                    data.d.results.forEach(function(entry) {
                        var resultDiv = document.createElement("div");
                        resultDiv.innerHTML = "<a href=\"" + entry.Url + "\"><span>" + entry.Title + "</span></a><br><span>" + entry.Description + "</span><br><br><hr><br>";

                        resultsDiv.appendChild(resultDiv);

                    });

                }
            });
        }
    }





})();