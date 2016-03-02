/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $("#btnDialog").click(function () {
                var url = "https://localhost:44300/app/home/auth.html";

                Office.context.ui.displayDialogAsync(url, { height: 40, width: 40, requireHTTPS: true }, function (result) {
                    var _dlg = result.value;
                    _dlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
                });
            });
        });
    };

    function processMessage(arg) {
        if (arg) {
            var token = JSON.parse(arg.message);
            $.ajax({
                url: "https://graph.microsoft.com/v1.0/me/contacts",
                headers: {
                    "accept": "application/json",
                    "Authorization": "Bearer " + token.accessToken
                },
                success: function (data) {
                    $("#results").html("success");
                },
                error: function (err) {
                    $("#results").html("error");
                }
            });
        }
    }
})();