/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global console, document, Excel, Office */

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    await init();
  }
});

import $ from "jquery";

function init() {
  $("#loginForm").submit(async function (event) {
    console.log("eddddd");
    event.preventDefault();
    // Code pour se connecter à Zoho Books

    var clientId = $("#clientId").val();
    await authenticate(clientId);
  });
}

function authenticate(clientId) {
  var redirectUri = "https://magical-puffpuff-f652ff.netlify.app/callback.html";
  var authorizationEndpoint = "https://accounts.zoho.com/oauth/v2/auth";
  var scope = "ZohoBooks.fullaccess.all"; // Adjust the scope as needed

  var authorizationUrl = `${authorizationEndpoint}?scope=${scope}&client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}`;

  Office.context.ui.displayDialogAsync(authorizationUrl, { height: 50, width: 50 }, function (asyncResult) {
    var dialog = asyncResult.value;

    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
      var message = JSON.parse(arg.message);
      if (message.status === "success") {
        var token = message.accessToken;

        console.log("Access Token:", token);
        // Utilisez le jeton d'accès pour vos requêtes API

        // Fermer la boîte de dialogue
        dialog.close();
        // Exécuter du code supplémentaire dans le contexte principal
        Office.context.ui.displayDialogAsync("about:blank", { displayInIframe: true });
        console.error("Authentication failed.");
        // Fermer la boîte de dialogue en cas d'échec d'authentification
      }
    });

    dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
      console.log("Dialog closed");
      // Rediriger ou exécuter du code dans le contexte principal après la fermeture de la boîte de dialogue
      // Par exemple, vous pouvez appeler une fonction spécifique ou recharger une partie de votre add-in
      document.location.href = "taskpane.html";

      //document.location.reload(); // Exemple: recharger la page actuelle
    });
  });
}
