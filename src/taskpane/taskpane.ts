/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import axios from "axios";

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("sideload-msg").style.display = "none";
    //  document.getElementById("app-body").style.display = "flex";
    //  document.getElementById("run").onclick = run;

    init();
  }
});

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }
var donneesExemple;
import $ from "jquery";

function init() {
  $("#importFromZoho").submit(async function (event) {
    console.log("eddddd");
    event.preventDefault();
    // Code pour se connecter à Zoho Books
    var datePickerFrom = $("#datePickerFrom").val();
    var datePickerTo = $("#datePickerTo").val();
    console.log("datePickerFrom", datePickerFrom);
    console.log("datePickerTo", $("#datePickerTo").val());

    handleZohoRedirect(datePickerFrom, datePickerTo);

    await insererDonneesDansNouvelleFeuille(donneesExemple);
  });
}
// Function to exchange authorization code for access token
/*

const exchangeCodeForToken = (code) => {
  const tokenParams = {
    grant_type: "refresh_token",
    client_id: zohoConfig.clientId,
    client_secret: zohoConfig.clientSecret,
    redirect_uri: zohoConfig.redirectUri,
    refresh_token: code,
  };

  axios
    .post(zohoConfig.tokenUrl, new URLSearchParams(tokenParams), {})

    .then((data) => {
      console.log("data", data);
      const accessToken = data.data.access_token;
      console.log(accessToken);
      axios
        .get(apiURL, {
          headers: {
            Authorization: `Zoho-oauthtoken ${accessToken}`,
          },
        })
        .then((response) => {
          console.log(response.data);
        });

      // Store access token securely (e.g., localStorage.setItem('access_token', accessToken))
      // Once stored, you can use it to make authenticated requests to Zoho APIs
    })
    .catch((error) => console.error("Error exchanging code for token:", error));
};
*/
// Function to handle redirect from Zoho with authorization code
//http://localhost:8010/proxy/api/v3/reports/metadata?entity_type=trial_balance&include_all_columns=true&is_response_new_flow=false&organization_id=849408688
//https://books.zoho.com/api/v3/reports/trialbalance?usestate=true&is_response_new_flow=true&response_option=1&organization_id=849408688
//lcp --proxyUrl https://books.zoho.com

//"http://localhost:8010/proxy/api/v3/reports/trialbalance?usestate=true&is_response_new_flow=true&response_option=1&organization_id=709668213";
//const apiURLreport = "http://localhost:8010/proxy/api/v3/reports/meta?organization_id=709668213";
//const proxyUrl = "https://cors-anywhere.herokuapp.com/";
async function handleZohoRedirect(datePickerFrom, datePickerTo) {
  const apiURL = `https://books.zoho.com/api/v3/reports/trialbalance?cash_based=false&filter_by=TransactionDate.CustomDate&from_date=${datePickerFrom}&to_date=${datePickerTo}&select_columns=%5B%7B%22field%22%3A%22name%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22account_code%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22net_debit%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22net_credit%22%2C%22group%22%3A%22report%22%7D%5D&is_for_date_range=true&show_rows=non_zero&sort_column=account&sort_order=A&usestate=true&is_response_new_flow=true&response_option=1&organization_id=709668213`;
  console.log("cccc", document.location);
  var hash = document.location.hash.substr(1);
  var result = hash.split("&").reduce(function (res, item) {
    var parts = item.split("=");
    res[parts[0]] = parts[1];
    return res;
  }, {});

  axios
    .get(apiURL, {
      headers: {
        Authorization: `Zoho-oauthtoken ${result.access_token}`,
      },
    })
    .then((response) => {
      console.log(response.data);
      donneesExemple = response.data;
    });
  console.log("bbbbb", result);
}

function insererDonneesDansNouvelleFeuille(donnees) {
  Excel.run(function (context) {
    // Créer une nouvelle feuille
    var nouvelleFeuille = context.workbook.worksheets.add();
    console.log("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa");
    // Définir le nom de la nouvelle feuille
    nouvelleFeuille.name = "zohoData";

    // Extraire les données des comptes
    const accounts = donnees.trialbalance.accounts;

    // Extraire les en-têtes_
    const headers = ["ACCOUNT", "ACCOUNT CODE", "NET DEBIT", "NET CREDIT"];
    const zohoData = [];
    // Transformer les données en tableau 2D
    for (let i = 0; i < accounts.length; i++) {
      const account = accounts[i];
      const accountsSubs1 = account.accounts;

      zohoData.push([
        account.name,
        account.account_code,
        account.net_debit_formatted,
        account.net_credit_sub_account_formatted,
      ]);
      if (accountsSubs1 && accountsSubs1.length > 0)
        for (let j = 0; j < accountsSubs1.length; j++) {
          const accountsSub1 = accountsSubs1[j];
          const accountsSubs2 = accountsSub1.accounts;

          zohoData.push([
            "    " + accountsSub1.name,
            accountsSub1.account_code,
            accountsSub1.values[0].net_debit_formatted,
            accountsSub1.values[0].net_credit_sub_account_formatted,
          ]);
          if (accountsSubs2 && accountsSubs2.length > 0)
            for (let k = 0; k < accountsSubs2.length; k++) {
              const accountsSub2 = accountsSubs2[k];

              zohoData.push([
                "         " + accountsSub2.name,
                accountsSub2.account_code,
                accountsSub2.values[0].net_debit_formatted,
                accountsSub2.values[0].net_credit_sub_account_formatted,
              ]);
            }
          if (accountsSub1.total_label)
            zohoData.push([
              accountsSub1.total_label,
              accountsSub1.account_code,
              accountsSub1.values[0].net_debit_sub_account_formatted,
              accountsSub1.values[0].net_credit_sub_account_formatted,
            ]);
        }
    }
    zohoData.push([
      donnees.trialbalance.total_label,
      donnees.trialbalance.account_code,
      donnees.trialbalance.values[0].net_debit_sub_account_formatted,
      donnees.trialbalance.values[0].net_credit_sub_account_formatted,
    ]);

    // Ajouter les en-têtes au début du tableau des données
    zohoData.unshift(headers);

    // Calculer la plage de cellules pour insérer les données
    const startCell = "A1";
    const endCell = `${String.fromCharCode(65 + headers.length - 1)}${zohoData.length}`;
    const range = nouvelleFeuille.getRange(`${startCell}:${endCell}`);

    // Insérer les données dans la nouvelle feuille
    range.values = zohoData;

    // Auto-fit des colonnes
    range.format.autofitColumns();

    const lastRow = zohoData.length;
    const lastRowRange = nouvelleFeuille.getRange(
      `A${lastRow}:${String.fromCharCode(65 + headers.length - 1)}${lastRow}`
    );

    // Appliquer le formatage à la dernière ligne
    lastRowRange.format.font.bold = true;
    lastRowRange.format.font.size = 12;
    // Exécuter les tâches en attente
    $("#datePickerFrom").val("");
    $("#datePickerTo").val("");
    return context.sync();
  }).catch(function (erreur) {
    console.log(erreur);
  });
}

("https://books.zoho.com/api/v3/reports/trialbalance?cash_based=false&filter_by=TransactionDate.CustomDate&from_date=2024-01-01&to_date=2024-08-31&select_columns=%5B%7B%22field%22%3A%22name%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22account_code%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22net_debit%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22net_credit%22%2C%22group%22%3A%22report%22%7D%5D&is_for_date_range=true&show_rows=non_zero&sort_column=account&sort_order=A&usestate=true&is_response_new_flow=true&response_option=1&organization_id=644595459");
