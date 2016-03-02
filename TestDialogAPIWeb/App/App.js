/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};

    app.clientId = "8e6a163d-a531-49cf-b666-d9ba1d920bca";
    app.tenant = "rzdemos.com";
    app.redirectUri = "https://localhost:44300/app/home/auth.html";

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };
    };

    return app;
})();