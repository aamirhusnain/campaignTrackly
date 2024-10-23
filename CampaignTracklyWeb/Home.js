﻿var app = angular.module('myApp', ['ngMaterial'], function ($mdThemingProvider) {

    $mdThemingProvider.theme('default')
        .primaryPalette('green', {
            'default': '500',
        });
});
app.controller('myCtrl', function ($scope, $mdToast, $log, $mdDialog, $element) {


    ProgressLinearActive();



    function generateCodeVerifier() {
        const array = new Uint8Array(32);
        window.crypto.getRandomValues(array);
        return Array.from(array, (byte) => ("0" + byte.toString(16)).slice(-2)).join("");
    }

    function generateCodeChallenge(codeVerifier) {
        return crypto.subtle.digest("SHA-256", new TextEncoder().encode(codeVerifier)).then((hash) => {
            return btoa(String.fromCharCode.apply(null, new Uint8Array(hash)))
                .replace(/\+/g, "-")
                .replace(/\//g, "_")
                .replace(/=+$/, "");
        });
    }

    let logoutDialog;

    $scope.logout = function () {
        
        let logutRedirectURI= "https://aamirhusnain.github.io/campaignTrackly/CampaignTracklyWeb/Templates/logout.html";
        
        let logoutURL = `https://dev-7ye0b6zftkobwhvz.us.auth0.com/v2/logout?client_id=hpbcTZOBJtUGwjR0PjNRSIrGYGGZ6fa4&returnTo=${logutRedirectURI}`;

        Office.context.ui.displayDialogAsync(logoutURL, { height: 50, width: 20 }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log('Dialog opened successfully:', asyncResult);
                logoutDialog = asyncResult.value;
                logoutDialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                    console.log('Message received from dialog:', arg.message);
                    if (arg.message === "Logout successful") {
                        // Handle logout success (e.g., navigate to a different page, update UI, etc.)
                        console.log('Logout was successful.');
                    }
                    logoutDialog.close();
                });
            } else {
                console.error('Error opening dialog:', asyncResult.error);
            }
        });
    };



   

    try {

        /////////// Functino for code refresh everytime perfectly ///////////
        function createGuid() {
            return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
                var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
                return v.toString(16);
            });
        };

        var guid = createGuid();

        console.log("working fine");

        $scope.LoginDiv = true;
        $scope.MainPageDiv = true;
        $scope.NavBarDiv = true;
        $scope.StartedScreen = true;
        $scope.SelectedOption = {};
        var APIToken = null;
        var FirstTime;
        $scope.UsedSheetData = [];
        $scope.result_Links;
        $scope.ChatGPTKey = "";
        var requiredCusTags = [];
        var isSetting = false;
        // $scope.replacedVal = false;

        function decryptAPIKey(encryptedKey, encryptionKey) {
            const decrypted = CryptoJS.AES.decrypt(encryptedKey, encryptionKey);
            return decrypted.toString(CryptoJS.enc.Utf8);
        };


        /////////// check user is logined or not ///////////
        var getFromLocal = window.localStorage.getItem("APIToken");
        if (getFromLocal != null) {
            getFromLocal = JSON.parse(getFromLocal);
            APIToken = getFromLocal.token;
        };


        var BaseURL = "https://devapp.campaigntrackly.com";
        //   var BaseURL = "https://app.campaigntrackly.com";

        /////////// show the started screen to user ///////////
        var checkUser = window.localStorage.getItem("UserVisted");

        if (checkUser === null) {
            $scope.StartedScreen = false;
            $scope.LoginDiv = true;
            $scope.MainPageDiv = true;
            $scope.NavBarDiv = true;
            FirstTime = true;

        } else {
            $scope.StartedScreen = true;
        };



        $scope.StartAddin = function () {
            window.localStorage.setItem("UserVisted", "Visted");
            $scope.StartedScreen = true;
            $scope.LoginDiv = false;
        };


        $scope.goToPlaylist = function () {
            window.open('https://youtube.com/playlist?list=PLK9s3o7S-dOf_NDebrNVG2jLHSXDWhp-D&si=gjzhM0a\_p--h2Rk8', '_blank');

        };


        $scope.clearSheet = function () {
            Excel.run(async (context) => {
                var sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.load("name");

                await context.sync();

                if (sheet.name !== "Settings") {
                    sheet.getRange().clear();
                    await context.sync();
                };
            }).catch(function (error) {
                console.error(error);
            });
        };



        $scope.gptBtn = false;
        $scope.tooltipText = "Ask Chat GPT";
        function startGptLoader() {
            document.getElementById("loaderGpt").style.display = 'block';
            $scope.tooltipText = "Please wait";
            $scope.gptBtn = true;
            if (!$scope.$$phase) {
                $scope.$apply();
            };
        };


        function endGptLoader() {
            document.getElementById("loaderGpt").style.display = 'none';
            $scope.tooltipText = "Ask Chat GPT";
            $scope.gptBtn = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            };
        };




        async function insertExistingChannel(data) {
            try {
                await Excel.run(async (context) => {

                    let sheet = context.workbook.worksheets.getItem("Settings");

                    var mediumColumn = [];
                    var sourceColumn = [];
                    var titleColumn = [];

                    //data.slice(1).forEach(function (item) {
                    data.forEach(function (item) {
                        mediumColumn.push([item.medium]);
                        sourceColumn.push([item.source]);
                        titleColumn.push([item.title]);
                    });

                    var mediumRange = sheet.getRange("AA1").getResizedRange(mediumColumn.length - 1, 0);
                    var sourceRange = sheet.getRange("AB1").getResizedRange(sourceColumn.length - 1, 0);
                    var titleRange = sheet.getRange("AC1").getResizedRange(titleColumn.length - 1, 0);

                    mediumRange.values = mediumColumn;
                    sourceRange.values = sourceColumn;
                    titleRange.values = titleColumn;

                    await context.sync();
                });
            } catch (error) {
                console.error(error);
            }
        }




        var enterEvent = new KeyboardEvent("keydown", {
            key: "Enter",
            code: "Enter",
            keyCode: 13,
            charCode: 13,
            bubbles: true
        });

        let dialog;


        Office.onReady(function () {

            try {




                async function getFirstRowData() {
                    await Excel.run(function (context) {
                        let sheet = context.workbook.worksheets.getActiveWorksheet();

                        let range = sheet.getUsedRange();

                        range.load("values");

                        return context.sync()
                            .then(function () {
                                const firstRowData = range.values[0];
                                
                                const targetHeader = "Campaign Name".toLowerCase();
                                const lowercasedFirstRowData = firstRowData.map(header => header.toLowerCase());

                                if (lowercasedFirstRowData.includes(targetHeader)) {
                                    setNameColumnDropdown();
                                }

                            });
                    }).catch(function (error) {
                        console.error(error);
                    });
                };

                async function setNameColumnDropdown() {
                    await Excel.run(async function (context) {
                        let sheet = context.workbook.worksheets.getActiveWorksheet();
                        let used_Range = sheet.getRange("A1:Z1");

                        used_Range.load("values");

                        await context.sync();

                        //console.log(used_Range.values);
                        //console.log(used_Range.address);

                        let headerRow = used_Range.values[0];
                        let columnIndex;

                        let targetHeader = "campaign name";

                        for (let colIndex = 0; colIndex < headerRow.length; colIndex++) {
                            if (headerRow[colIndex].toLowerCase() === targetHeader) {
                                columnIndex = colIndex;
                                break;
                            }
                        }

                        let columnName = String.fromCharCode(65 + columnIndex);

                        //console.log("Column Name for '" + targetHeader + "': " + columnName);

                        // columnName

                        var campaignRange = sheet.getRange(`${columnName}2:${columnName}3`);

                        //var campaignRange = sheet.getRange("A2:A3");

                        const nameSourceRange = context.workbook.worksheets.getItem("Settings").getRange("A:A");

                        let approvedListRule = {
                            list: {
                                inCellDropdown: true,
                                source: nameSourceRange
                            }
                        };

                        campaignRange.dataValidation.clear();



                        campaignRange.dataValidation.rule = approvedListRule;
                        campaignRange.dataValidation.prompt = {
                            message: "Select existing campaign or click the 'Manually Enter Tags' icon to type a new one.",
                            showPrompt: true,
                            title: "Campaign Name"
                        };



                        await context.sync();
                    });
                };


                //async function setNameColumnDropdown() {
                //    await Excel.run(async function (context) {
                //        let sheet = context.workbook.worksheets.getActiveWorksheet();
                //        let used_Range = sheet.getRange("A1:Z1");

                //        used_Range.load("values");
                //        await context.sync();

                //        let headerRow = used_Range.values[0];
                //        let columnsToValidate = {
                //            "campaign name": null,
                //            "medium": null,
                //            "source": null
                //        };

                //        for (let targetHeader in columnsToValidate) {
                //            for (let colIndex = 0; colIndex < headerRow.length; colIndex++) {
                //                if (headerRow[colIndex].toLowerCase() === targetHeader) {
                //                    let columnIndex = colIndex;
                //                    let columnName = String.fromCharCode(65 + columnIndex);
                //                    columnsToValidate[targetHeader] = columnName;
                //                    break;
                //                }
                //            }
                //        }

                //        console.log("Columns to Validate: ", columnsToValidate);

                //        let settingsSheet = context.workbook.worksheets.getItem("Settings");

                //        for (let targetHeader in columnsToValidate) {
                //            if (columnsToValidate[targetHeader]) {
                //                let columnName = columnsToValidate[targetHeader];
                //                let campaignRange = sheet.getRange(`${columnName}2:${columnName}100`); // Adjust range as necessary

                //                let sourceColumn;
                //                if (targetHeader === "medium") {
                //                    sourceColumn = "AA";
                //                } else if (targetHeader === "campaign name") {
                //                    sourceColumn = "A";
                //                } else if (targetHeader === "source") {
                //                    sourceColumn = "AB";
                //                } else {
                //                    continue;
                //                }

                //                let nameSourceRange = settingsSheet.getRange(`${sourceColumn}:${sourceColumn}`);

                //                let approvedListRule = {
                //                    list: {
                //                        inCellDropdown: true,
                //                        source: nameSourceRange
                //                    }
                //                };

                //                campaignRange.dataValidation.clear();
                //                campaignRange.dataValidation.rule = approvedListRule;
                //            }
                //        }

                //        await context.sync();
                //    });
                //}




                $scope.OpenDialog = function (ev) {

                    $mdDialog.show({
                        scope: $scope.$new(),
                  //      templateUrl: '/Templates/SheetConfirm.html',
                       templateUrl: '/campaignTrackly/CampaignTracklyWeb/Templates/SheetConfirm.html',
                    //      templateUrl: 'https://app.campaigntrackly.com/excel-addin/CampaignTracklyWeb/Templates/SheetConfirm.html',
                        parent: angular.element(document.body),
                        targetEvent: ev,
                        clickOutsideToClose: false,
                        escapeToClose: false,
                        controller: ['$scope', '$mdDialog', function ($scope, $mdDialog) {

                        }]
                    });

                };

                /////////// Ask Question ///////////

                $scope.chatGpt = function () {
                    startGptLoader();


                    Excel.run(function (context) {
                        var sheet = context.workbook.worksheets.getActiveWorksheet();
                        var range = context.workbook.getSelectedRange();
                        range.load("address");
                        range.load("values");

                        return context.sync().then(function () {
                            var cellValueArr = range.values;
                            var cellValue = cellValueArr[0][0];

                            var cellAddress = range.address;

                            var cellAddrArr = cellAddress.split("!");
                            var Cell_Address = cellAddrArr[1];


                            if (cellValue != "") {

                                /////////// Chat with GPT ///////////

                                var settings = {
                                    "url": "https://api.openai.com/v1/chat/completions",
                                    "method": "POST",
                                    "timeout": 0,
                                    "headers": {
                                        "Content-Type": "application/json",
                                        "Authorization": "Bearer " + $scope.ChatGPTKey
                                    },
                                    "data": JSON.stringify({
                                        "model": "gpt-3.5-turbo",
                                        "messages": [
                                            {
                                                "role": "user",
                                                "content": cellValue
                                            }
                                        ],
                                        "max_tokens": 2000,
                                        "temperature": 0
                                    }),
                                };

                                $.ajax(settings).done(function (response) {
                                    var Answer = response.choices[0].message.content.trim();

                                    //  console.log(Answer);

                                    var cellAddrRow = parseInt(Cell_Address.replace(/\D/g, ""));
                                    var alphabets = Cell_Address.match(/[a-zA-Z]/g);
                                    var nextCellRow = cellAddrRow + 1;
                                    var nextRowAdd = alphabets + nextCellRow.toString();

                                    //  console.log(nextRowAdd);


                                    var NewRange = sheet.getRange(nextRowAdd);
                                    NewRange.values = [[Answer]];

                                    NewRange.format.wrapText = true;
                                    NewRange.format.columnWidth = 350;

                                    return context.sync()
                                        .then(function () {
                                            NewRange.format.autofitColumns();
                                            endGptLoader();

                                            if (!$scope.$$phase) {
                                                $scope.$apply();
                                            };

                                        }).catch(function (error) {
                                            //   console.log("Error: " + error);
                                            endGptLoader();
                                        });


                                }).fail(function (error, xhr) {
                                    //  console.log(error);
                                    endGptLoader();
                                    loadToast("Connection Issue. Please contact support@campaigntrackly.com");

                                });



                            } else {
                                endGptLoader();
                                loadToast("Please put data in cell.");

                            };
                        }).catch(function (error) {
                            //   console.log("Error occurred during context sync: " + error);
                            endGptLoader();
                            loadToast("Cannot perform this operation while Excel is in editing mode.");
                        });


                    });



                };


                function getAlphabeticCharacter(number) {
                    if (typeof number !== 'number' || number < 1 || number > 26) {
                        return 'Invalid';
                    }

                    const charCode = number + 96;
                    const character = String.fromCharCode(charCode);
                    return character.toUpperCase();
                };

                function checkWordOrSentence(input) {
                    var trimmedInput = input.trim();

                    var words = trimmedInput.split(" ");

                    if (words.length === 1) {
                        return "word";
                    } else {
                        return "sentence";
                    }
                };


                function removeDots(string) {
                    return string.replace(/\./g, '');
                }

                function SeparatWord(sentance) {
                    var trst = sentance;
                    var splidtedArr = trst.split('"');
                    //  console.log(splidtedArr[1]);
                    return splidtedArr[1];
                };
                function checkSpelling(text, beforeWord, adressOfCell) {



                    var endpoint = 'https://api.openai.com/v1/chat/completions';


                    //  var prompt = "Please spell check the following word: " + text + ". If the word is spelled correctly, return 'correct'. If the word is spelled incorrectly, but I know the correct spelling, return the correct spelling. If I do not know the correct spelling, return unknown."
                    var prompt = `The spelling of the word "` + text + `" is`;
                    // var prompt = `correct spelling '"${text}"' if incorrect then please correct it.`;
                    var data = {
                        'model': 'gpt-3.5-turbo',
                        'messages': [{ 'role': 'user', 'content': prompt }],
                        max_tokens: 500,
                        temperature: 0
                    };


                    $.ajax({
                        url: endpoint,
                        async: false,
                        headers: {
                            'Authorization': 'Bearer ' + $scope.ChatGPTKey,
                            'Content-Type': 'application/json'
                        },
                        method: 'POST',
                        data: JSON.stringify(data),
                        success: function (response) {
                            var reply = response.choices[0].message.content;

                            answerArr = reply.split('.');
                            var unkownAns = answerArr[1].trim();
                            //      console.log(unkownAns);
                            if (unkownAns === 'It is not a recognized word in the English language' || unkownAns === "There is no such word in the English language") {
                                loadToast("I am not sure what this word is, please try again");
                            } else {

                                var CheckSen = checkWordOrSentence(reply);
                                var fullReply = removeDots(reply);
                                if (fullReply != 'correct') {

                                    if (CheckSen === "sentence") {
                                        var onlyWord = SeparatWord(fullReply);

                                    } else {
                                        var onlyWord = fullReply;
                                    };

                                    if (onlyWord != '' && onlyWord != undefined) {

                                        if (onlyWord.toLowerCase() != beforeWord.toLowerCase()) {
                                            showActionToast("Spelling might need to be corrected? Thank you", onlyWord, adressOfCell);
                                        } else {
                                        };
                                    };
                                } else {
                                };



                            }



                        },
                        error: function (error) {
                            //  console.error('Error:', error);


                            if (error.responseJSON) {
                                if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                    RefreshToken(getFromLocal.refresh_token);
                                    if (tokenFreshed) {
                                        $scope.getTagTemplates();
                                    };
                                } else {
                                    loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                }
                            } else {
                                loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                            }

                            ProgressLinearInActive();

                        }
                    });
                };



                $scope.eventResult;
                Excel.run(function (context) {
                    var sheet = context.workbook.worksheets.getActiveWorksheet();
                    $scope.eventResult = sheet.onChanged.add($scope.handleOnChange);
                    return context.sync();
                }).catch(function (error) {
                    //   console.log(error);
                });



                $scope.handleOnChange = function (eventArgs) {

                    var address = eventArgs.address;

                    if (address === "A1" && eventArgs.details.valueAfter === "Campaign Name") {
                    //if (eventArgs.details.valueAfter.toLowerCase() === "campaign name") {
                        setNameColumnDropdown();
                    };

                    //if (["campaign name", "source", "medium"].includes(eventArgs.details.valueAfter.toLowerCase())) {
                    //    setNameColumnDropdown();
                    //}


                    var rowNumber = address.split(":")[0].match(/\d+/)[0];


                    if (rowNumber === "1" && eventArgs.details.valueAfter != '') {
                        // ProgressLinearActive();

                        var wordForCheckSpell = eventArgs.details.valueAfter;

                        var wordForCheck = eventArgs.details.valueAfter;
                        if (wordForCheck.toLowerCase() != eventArgs.details.valueBefore.toLowerCase()) {
                            if (wordForCheck.toLowerCase() != "campaign name" && wordForCheck.toLowerCase() != "url" && wordForCheck.toLowerCase() != "content" &&
                                wordForCheck.toLowerCase() != "term" && wordForCheck.toLowerCase() != "source" && wordForCheck.toLowerCase() != "medium") {

                                checkSpelling(wordForCheckSpell, eventArgs.details.valueBefore, address);
                                //       console.log("Change Value");

                            } else {
                                //        console.log("Header Value");
                                ProgressLinearInActive();
                            };

                        } else {
                            ProgressLinearInActive();
                        };


                    };
                };







                /////////// check token expiration ///////////
                //function isTokenExpired(token) {
                //   // const base64Url = token.split(".")[1];
                //    const base64 = token;
                //    const jsonPayload = decodeURIComponent(
                //        atob(base64)
                //            .split("")
                //            .map(function (c) {
                //                return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
                //            })
                //            .join("")
                //    );

                //    const { exp } = JSON.parse(jsonPayload);
                //    var expnew = exp * 1000;

                //    var ee = new Date(Date.now());
                //    var ef = new Date(expnew);
                //    if (ee > ef) {
                //        expired = true;
                //    }
                //    else {
                //        expired = false;
                //    };

                //    return expired
                //};

                function isTokenExpired(token) {
                    const payload = JSON.parse(atob(token.split('.')[1]));
                    const expirationTimestamp = payload.exp * 1000;

                    const currentTimestamp = Date.now();

                    return expirationTimestamp <= currentTimestamp;
                };

                var tokenFreshed;

                //////////////////////// Refresh token ////////////////////////

                function RefreshToken(refshToken) {
                    var getRefFromLocal = window.localStorage.getItem("APIToken");
                    getRefFromLocal = JSON.parse(getRefFromLocal);
                    var settings = {
                        "url": BaseURL + "/wp-json/campaigntrackly/v1/refresh_token",
                        "method": "POST",
                        "timeout": 0,
                        "async": false,
                        "headers": {
                            "Accept": "application/json",
                            "Content-Type": "application/x-www-form-urlencoded"
                        },
                        "data": {
                            "refresh_token": getRefFromLocal.refresh_token
                        }
                    };


                    $.ajax(settings).done(function (response) {
                        //  console.log(response);

                        if (response.statusCode === 200) {
                            APIToken = response.data.token;
                            window.localStorage.setItem("APIToken", JSON.stringify(response.data));
                            ProgressLinearInActive();;
                            tokenFreshed = true;
                        } else {
                            $scope.LoginDiv = false;
                            $scope.MainPageDiv = true;
                            $scope.NavBarDiv = true;
                            loadToast(response.message)
                            tokenFreshed = false;

                            // ProgressLinearInActive();;
                        };
                        ProgressLinearInActive();;

                    }).fail(function (error) {
                        // console.log(error);

                        ProgressLinearInActive();;
                        loadToast("Connection Issue. Please contact support@campaigntrackly.com");

                    });


                };



                // Encryption function using AES
                function encryptAPIKey(apiKey, encryptionKey) {
                    const encrypted = CryptoJS.AES.encrypt(apiKey, encryptionKey);
                    return encrypted.toString();
                };
                
                
                
                /////////////////// New Auth Code /////////////
                
                
                
                
                
let loginDialog;

$scope.SignIn = async function () {


  try {
    const domain = "dev4-8sn87rh0bu88b2rz.us.auth0.com";
    const clientId = "cVQubOWRbtDVF9Mpsy5Vx0EXTzcdkvHM";
    const redirectUri = "https://aamirhusnain.github.io/campaignTrackly/CampaignTracklyWeb/Templates/RedirectPage.html";

    const responseType = "code";
    const responseMode = "query";
    const scope = "openid profile email";
    const state = btoa(JSON.stringify({ timestamp: Date.now() }));
    const codeVerifier = generateCodeVerifier();
    const codeChallenge = await generateCodeChallenge(codeVerifier);
    const codeChallengeMethod = "S256";
    const auth0Client = btoa(JSON.stringify({ name: "auth0-react", version: "2.2.4" }));

    localStorage.setItem("code_verifier", codeVerifier);

    const dialogeURL = `https://${domain}/authorize?client_id=${clientId}&scope=${encodeURIComponent(scope)}&redirect_uri=${encodeURIComponent(redirectUri)}&response_type=${responseType}&response_mode=${responseMode}&state=${encodeURIComponent(state)}&nonce=${encodeURIComponent(state)}&code_challenge=${codeChallenge}&code_challenge_method=${codeChallengeMethod}&auth0Client=${auth0Client}`;

    console.log(dialogeURL);
    Office.context.ui.displayDialogAsync(dialogeURL, { height: 80, width: 20 }, function (asyncResult) {
      console.log(asyncResult);
      loginDialog = asyncResult.value;
      loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    });

  } catch (error) {
    console.error("Error during the authentication process:", error);
  }



  // const domain = "dev-7ye0b6zftkobwhvz.us.auth0.com";
  // const clientId = "hpbcTZOBJtUGwjR0PjNRSIrGYGGZ6fa4";


  // //const domain = "dev4-8sn87rh0bu88b2rz.us.auth0.com";
  // //const clientId = "cVQubOWRbtDVF9Mpsy5Vx0EXTzcdkvHM";




  // const redirectUri = "https://localhost:44371/templates/RedirectPage.html";




  // const responseType = "code";
  // const responseMode = "query";
  // const scope = "openid profile email";
  // const state = btoa(JSON.stringify({ timestamp: Date.now() }));
  // const codeVerifier = generateCodeVerifier();
  // const codeChallenge = await generateCodeChallenge(codeVerifier);
  // const codeChallengeMethod = "S256";
  // const auth0Client = btoa(JSON.stringify({ name: "auth0-react", version: "2.2.4" }));

  // localStorage.setItem("code_verifier", codeVerifier);

  // const dialogeURL = `https://${domain}/authorize?client_id=${clientId}&scope=${encodeURIComponent(scope)}&redirect_uri=${encodeURIComponent(redirectUri)}&response_type=${responseType}&response_mode=${responseMode}&state=${encodeURIComponent(state)}&nonce=${encodeURIComponent(state)}&code_challenge=${codeChallenge}&code_challenge_method=${codeChallengeMethod}&auth0Client=${auth0Client}`;
  // //const dialogeURL = "https://localhost:44371/templates/RedirectPage.html";

  // console.log(dialogeURL);
  // Office.context.ui.displayDialogAsync(dialogeURL, { height: 50, width: 20 }, function (asyncResult) {
  //     console.log(asyncResult);
  //     loginDialog = asyncResult.value;
  //     loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  // });
};

function processMessage(arg) {
  const message = arg.message;

  if (message) {
    // Close the dialog
    loginDialog.close();

    // Now exchange the authorization code for an access token
    exchangeAuthorizationCodeForToken(message);
  } else {
    console.error("Authorization code not found in the response.");
  }
  // getToken(arg.message).then((tokenResponse) => {
  //     console.log(tokenResponse);
  //     APIToken = tokenResponse.access_token;
  //     window.localStorage.setItem("APIToken", JSON.stringify(tokenResponse));
  //     loginDialog.close();
  // }).catch((error) => {
  //     console.error(error);
  // });
}


async function fetchGPTToken(bearerToken) {
    const url = `${BaseURL}/wp-json/campaigntrackly/v1/gpt_token`;
    const headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Authorization": `Bearer ${bearerToken}`
    };

    try {
        const response = await fetch(url, {
            method: "POST",
            headers: headers,
            timeout: 0 // Fetch does not support timeout directly, so you might need to handle this manually if required.
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        const token = result.data.token;

        $scope.ChatGPTKey = token;

        const encryptedKey = encryptAPIKey($scope.ChatGPTKey, 'ChatGPTKey');
        window.localStorage.setItem('SecretKey', encryptedKey);

    } catch (error) {
        console.error("Connection Issue. Please contact support@campaigntrackly.com", error);
        loadToast("Connection Issue. Please contact support@campaigntrackly.com");
    }
}



async function exchangeAuthorizationCodeForToken(authorizationCode) {
  const domain = "dev4-8sn87rh0bu88b2rz.us.auth0.com";
  const clientId = "cVQubOWRbtDVF9Mpsy5Vx0EXTzcdkvHM";
  const redirectUri = "https://aamirhusnain.github.io/campaignTrackly/CampaignTracklyWeb/Templates/RedirectPage.html";

  const tokenEndpoint = `https://${domain}/oauth/token`;
  const codeVerifier = localStorage.getItem("code_verifier");

  const tokenRequestBody = {
    grant_type: "authorization_code",
    client_id: clientId,
    code: authorizationCode,
    redirect_uri: redirectUri,
    code_verifier: codeVerifier,
  };

  try {
    console.log("Exchanging authorization code for access token...");
    const response = await fetch(tokenEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(tokenRequestBody),
    });

    const tokenResponse = await response.json();

    if (response.ok) {
      console.log("Access Token:", tokenResponse.access_token);

      // Call the UserInfo API to get user information
      const userInfo = await callUserInfoAPI(tokenResponse.access_token);

      // Call the CampaignTrackly token API with the user's email and Auth0 token
      if (userInfo && userInfo.email) {
        const campaignTracklyToken = await callCampaignTracklyTokenAPI(userInfo.email, tokenResponse.access_token);

        // Use the CampaignTrackly token to call the channels API
        if (campaignTracklyToken) {

        //   console.log(campaignTracklyToken);
        await fetchGPTToken(campaignTracklyToken);
        }

      }
    } else {
      console.error("Error exchanging code for token:", tokenResponse);
    }
  } catch (error) {
    console.error("Network error exchanging code for token:", error);
  }
}

async function callUserInfoAPI(accessToken) {
  const domain = "dev4-8sn87rh0bu88b2rz.us.auth0.com";
  const userInfoEndpoint = `https://${domain}/userinfo`;

  try {
    console.log("Calling Auth0 UserInfo API...");
    const response = await fetch(userInfoEndpoint, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    const userInfo = await response.json();

    if (response.ok) {
      console.log("User Info:", userInfo);
      return userInfo;
    } else {
      console.error("Error calling UserInfo API:", userInfo);
    }
  } catch (error) {
    console.error("Network error calling UserInfo API:", error);
  }
  return null;
}




async function callCampaignTracklyTokenAPI(email, auth0Token) {
  const campaignTracklyEndpoint = 'https://devapp.campaigntrackly.com/wp-json/campaigntrackly/v1/login';

  const requestBody = new URLSearchParams();
  requestBody.append('email', email);
  requestBody.append('auth0_token', auth0Token);

  try {
    console.log("Calling CampaignTrackly Token API...");
    const response = await fetch(campaignTracklyEndpoint, {
      method: 'POST',
      headers: {
        'accept': 'application/json',
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: requestBody,
    });

    const campaignTracklyResponse = await response.json();

    if (response.ok) {
      console.log("CampaignTrackly Token Response:", campaignTracklyResponse);
      
      

      $scope.LoginDiv = true;
      $scope.MainPageDiv = false;
      $scope.NavBarDiv = false;
      
          if (campaignTracklyResponse.data.token) {
        APIToken = campaignTracklyResponse.data.token;
        window.localStorage.setItem("APIToken", JSON.stringify(campaignTracklyResponse.data));
        
        
          $scope.LoadSetting();

        $scope.UserName = undefined;
        $scope.UserPassword = undefined;

        $scope.getTagTemplates();
        
          }
      
      
      return campaignTracklyResponse.data.token;
    } else {
      console.error("Error calling CampaignTrackly Token API:", campaignTracklyResponse);
    }
  } catch (error) {
    console.error("Network error calling CampaignTrackly Token API:", error);
  }
  return null;
}
                
                
                
                
                
                ////////////////////////////////////////////////
                
                
                
                



                //////////////////////// Sign In ////////////////////////



                //const auth0 = new Auth0Client({
                //    domain: "dev-sqc8vsqw1qbf38cs.us.auth0.com",
                //    client_id: "lD3jeDPKex8oIFIs4CK4UdRQAW3EnMQO",
                //    redirect_uri: "https://localhost:44371/templates/RedirectPage.html"
                //});


                //const auth0 = new Auth0Client({
                //    domain: "dev-7ye0b6zftkobwhvz.us.auth0.com",
                //    client_id: "hpbcTZOBJtUGwjR0PjNRSIrGYGGZ6fa4",
                //    redirect_uri: "https://localhost:44371/templates/RedirectPage.html"
                //});

                //$scope.SignIn = function () {

                //    auth0.loginWithPopup().then((response) => {
                //        auth0.getUser().then((user) => {
                //            console.log('User:', user);
                //            // Handle user information here
                //        });
                //    }).catch((error) => {
                //        console.error('Login failed:', error);
                //    });
                //};





                //$scope.SignIn = function () {
                //    ProgressLinearActive();


                //    var settings = {
                //        "url": BaseURL + "/wp-json/campaigntrackly/v1/login",
                //        "method": "POST",
                //        "timeout": 0,
                //        "headers": {
                //            "Content-Type": "application/x-www-form-urlencoded",
                //        },
                //        "data": {
                //            "username": $scope.UserName.trim(),
                //            "password": $scope.UserPassword.trim()
                //        }
                //    };

                //    $.ajax(settings).done(function (response) {
                //        //  console.log(response);

                //        if (response.statusCode === 200) {



                //            $scope.LoginDiv = true;
                //            $scope.MainPageDiv = false;
                //            $scope.NavBarDiv = false;
                //            if (response.data.token) {
                //                APIToken = response.data.token;
                //                window.localStorage.setItem("APIToken", JSON.stringify(response.data));




                //                $.ajax({
                //                    url: BaseURL + "/wp-json/campaigntrackly/v1/gpt_token",
                //                    method: "POST",
                //                    "timeout": 0,
                //                    "headers": {
                //                        "Content-Type": "application/x-www-form-urlencoded"
                //                    },
                //                    "data": {
                //                        "username": $scope.UserName.trim(),
                //                        "password": $scope.UserPassword.trim()
                //                    },
                //                    success: function (result) {
                //                        //console.log(result);
                //                        $scope.ChatGPTKey = result.data.token;

                //                        const encryptedKey = encryptAPIKey($scope.ChatGPTKey, 'ChatGPTKey');
                //                        //    console.log(encryptedKey);

                //                        window.localStorage.setItem('SecretKey', encryptedKey);
                //                    },
                //                    error: function (error) {
                //                        //     console.log(error);
                //                        loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                //                    }
                //                });


                //                $scope.LoadSetting();

                //                $scope.UserName = undefined;
                //                $scope.UserPassword = undefined;

                //                $scope.getTagTemplates();

                //            };

                //        } else {
                //            $scope.LoginDiv = false;
                //            $scope.MainPageDiv = true;
                //            $scope.NavBarDiv = true;
                //            // ProgressLinearInActive();;
                //        };



                //    }).fail(function (error) {
                //        //  console.log(error);
                //        if (error.responseJSON.statusCode) {

                //            if (error.responseJSON.statusCode === 403 || error.responseJSON.code === "application_passwords_disabled") {
                //                loadToast(error.responseJSON.message, true);
                //            } else {
                //                loadToast("Login Failed", true);
                //            };
                //        };

                //        ProgressLinearInActive();

                //    });

                //};

                var AllCustomTagName = [];



                window.localStorage.removeItem("LastAddress");


                $scope.LoadSetting = async function () {
                    // Fetch channel data
                    async function fetchChannelData() {
                        return $.ajax({
                            url: BaseURL + '/wp-json/campaigntrackly/v1/channels',
                            method: 'GET',
                            headers: {
                                'Authorization': 'Bearer ' + APIToken
                            }
                        });
                    }

                    // Fetch custom tags
                    async function fetchCustomTags() {
                        return $.ajax({
                            url: BaseURL + '/wp-json/campaigntrackly/v1/custom_tags',
                            method: 'GET',
                            headers: {
                                'accept': 'application/json',
                                'Authorization': 'Bearer ' + APIToken
                            }
                        });
                    }

                    try {
                        // Perform both AJAX calls asynchronously
                        const [channelResponse, customTagsResponse] = await Promise.all([fetchChannelData(), fetchCustomTags()]);

                        const channelData = channelResponse; // Channel data response
                        const customTagsData = customTagsResponse; // Custom tags response

                        AllCustomTagName = customTagsData;

                        await Excel.run(async (context) => {
                            var workbook = context.workbook;
                            var worksheets = workbook.worksheets;
                            var newSheetName = "Settings";

                            var existingSheet = worksheets.getItemOrNullObject(newSheetName);
                            existingSheet.load("name");

                            await context.sync();

                            if (existingSheet.isNullObject) {
                                var newSheet = worksheets.add(newSheetName);

                                newSheet.getRange("B:AZ").clear();
                                await context.sync();

                                await insertExistingChannel(channelData);
                                loadExistingCampaigns();
                                checkSetings();
                            } else {
                                var SettingSheet = context.workbook.worksheets.getItem(newSheetName);
                                SettingSheet.getRange("B:AZ").clear();
                                await context.sync();

                                await insertExistingChannel(channelData);
                                loadExistingCampaigns();
                                checkSetings();
                            }
                        });

                    } catch (error) {
                        console.error("Error fetching data:", error);
                        // Handle the error response here
                        if (error.status !== 200 && error.status !== 500) {
                            if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                RefreshToken(getFromLocal.refresh_token);
                                ProgressLinearActive();
                                // $scope.LoadSetting();
                            } else if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Signature verification failed") {
                                $scope.logOut();
                            } else {
                                loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                            }
                        } else {
                            loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                        }

                        ProgressLinearInActive();
                    }
                };





                function loadExistingCampaigns() {


                    var settings = {
                        "url": BaseURL + "/wp-json/campaigntrackly/v1/campaigns",
                        "method": "GET",
                        "timeout": 0,
                        "headers": {
                            "Authorization": "Bearer " + APIToken
                        },
                    };

                    $.ajax(settings).done(function (response) {

                        var AllExistingCampaigns = response.campaigns.map(x => [x.name]);
                        //console.log(AllExistingCampaigns);
                        //     AllExistingCampaigns.unshift(["Existing Campaigns"]);
                        setCampaigns(AllExistingCampaigns);

                    }).fail(function (error) {

                        // Handle the error response here
                        if (error.status != 200 && error.status != 500) {
                            if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                RefreshToken(getFromLocal.refresh_token);
                                ProgressLinearActive();
                                //   $scope.LoadSetting();
                            }
                            else if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Signature verification failed") {

                                $scope.logOut();
                            }
                            else {
                                loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                            };
                        } else {
                            loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                        };

                        ProgressLinearInActive();

                    });

                };




                function checkSetings() {


                    // Make an AJAX GET request
                    $.ajax({
                        url: BaseURL + '/wp-json/campaigntrackly/v1/settings',
                        type: 'GET',
                        headers: {
                            'accept': 'application/json',
                            "Authorization": "Bearer " + APIToken
                        },
                        success: function (response) {

                            //console.log(response);

                            if (response.campaign_options.mode === "Auto-populate both URL campaign name and utm_campaign tag from custom tags") {
                                isSetting = true;
                            };

                            requiredCusTags = response.campaign_options.custom_tags;


                        },
                        error: function (xhr, status, error) {
                            // Handle the error response here
                            if (error.status != 200 && error.status != 500) {
                                if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                    RefreshToken(getFromLocal.refresh_token);
                                    ProgressLinearActive();
                                    //   $scope.LoadSetting();
                                }
                                else if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Signature verification failed") {

                                    $scope.logOut();
                                }
                                else {
                                    loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                };
                            } else {
                                loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                            };

                            ProgressLinearInActive();


                        }
                    });





                };





                async function setCampaigns(existingCampaigns) {
                    await Excel.run(async (context) => {
                        let sheet = context.workbook.worksheets.getItem("Settings");
                        let range = sheet.getRange("A1:A" + (existingCampaigns.length));
                        range.values = existingCampaigns;
                        range.format.autofitColumns();
                        // Sync to apply changes
                        await context.sync();
                        getFirstRowData();
                    });
                };



                var AllCustoms = [];
                var AllCustomValues = [];

                $scope.onSelectChange = function () {
                    //console.log($scope.SelectedOption);

                    setHeaderChannel();

                    AllCustoms = [];
                    AllCustomValues = [];

                    // You can perform any additional actions here based on the selected value.

                    for (var i = 0; i < $scope.SelectedOption.custom.length; i++) {
                        if ($scope.SelectedOption.custom[i].title != null) {
                            AllCustoms.push($scope.SelectedOption.custom[i].title);

                        }
                    };


                    function searchData(jsonArray, searchParams) {
                        return jsonArray.filter(item => searchParams.includes(item.custom));
                    }

                    const searchResults = searchData(AllCustomTagName, AllCustoms);
                    // console.log(searchResults);

                    AllCustomValues = searchResults;
                    //console.log(AllCustomValues);
                    //console.log(AllCustoms);

                    if (searchResults.length > 0) {

                        AllCustoms.sort((a, b) => a.localeCompare(b));


                        searchResults.sort((a, b) => {
                            const customA = a.custom.toLowerCase();
                            const customB = b.custom.toLowerCase();

                            if (customA < customB) {
                                return -1;
                            }
                            if (customA > customB) {
                                return 1;
                            }
                            return 0;
                        });



                        setData(searchResults);
                        setHeadOnWorkingSheet(AllCustoms);
                    };

                };




                function getNextAlphabet(char) {
                    // Convert the character to uppercase to handle lowercase inputs
                    char = char.toUpperCase();

                    // Check if the input character is a valid alphabet character
                    if (/^[A-Z]$/.test(char)) {
                        // Get the ASCII code of the character
                        const charCode = char.charCodeAt(0);

                        // If the character is 'Z', return 'A'
                        if (charCode === 90) {
                            return 'A';
                        } else {
                            // Increment the ASCII code by 1 and convert it back to a character
                            return String.fromCharCode(charCode + 1);
                        }
                    } else {
                        // If the input is not a valid alphabet character, return null
                        return null;
                    }
                }


                $scope.checkChennel = async function () {
                    try {
                        const result = await Excel.run(async (context) => {
                            var sheet = context.workbook.worksheets.getActiveWorksheet();
                            var range = sheet.getUsedRange();
                            range.load("values");

                            await context.sync();
                            var firstRowData = range.values[0];

                        //    console.log(firstRowData);

                            if (firstRowData.includes("Medium") && firstRowData.includes("Source")) {
                                return true;
                            } else {
                             //   console.log("Channel Header does not exist");
                                return false;
                            }
                        });

                        return result;
                    } catch (error) {
                        console.error(error);
                        return false; // or handle the error as needed
                    }
                };



                async function setHeaderChannel() {
                    try {

                        let channelHeaderExists = await $scope.checkChennel();

                        if (!channelHeaderExists) {


                            await Excel.run(async function (context) {
                                // Get the active worksheet
                                const sheet = context.workbook.worksheets.getActiveWorksheet();

                                // Get the used range of the active sheet
                                const usedRange = sheet.getUsedRange();

                                // Load the address of the used range
                                usedRange.load("address");

                                // Execute the queued commands and return a promise to indicate task completion
                                await context.sync();

                                let onlyRange = usedRange.address.split('!')[1];

                                let lastUsedCol = onlyRange.slice(-2, -1);

                                let sourceCol = getNextAlphabet(lastUsedCol);
                                let mediumCol = getNextAlphabet(sourceCol);

                                let ChannelHeaderRange = sourceCol + "1:" + mediumCol + "1";

                             //   console.log(ChannelHeaderRange);

                                let ChannelHeader = sheet.getRange(ChannelHeaderRange);

                                ChannelHeader.values = [["Medium", "Source"]];

                                let sourceDDLCell = sheet.getRange(sourceCol + "2");
                                let mediumDDLCell = sheet.getRange(mediumCol + "2");

                                // Get the name source range from the "Settings" worksheet
                                const nameSourceRange = context.workbook.worksheets.getItem("Settings").getRange("AA:AA");
                                const nameMediumRange = context.workbook.worksheets.getItem("Settings").getRange("AB:AB");
                                // Define the data validation rule
                                let approvedListRule = {
                                    list: {
                                        inCellDropdown: true,
                                        source: nameSourceRange
                                    }
                                };

                                let approvedListRuleMedium = {
                                    list: {
                                        inCellDropdown: true,
                                        source: nameMediumRange
                                    }
                                };

                                // Apply data validation to the source column range

                               

                                sourceDDLCell.dataValidation.rule = approvedListRule;
                                mediumDDLCell.dataValidation.rule = approvedListRuleMedium;

                           
                                await context.sync();
                            });

                        } else {
                         //   console.log("Channel Header already exists");
                        };

                    } catch (error) {
                        console.error(error);
                    }
                };









                $scope.SetAsManual = function () {
                    Excel.run(function (context) {
                        // Get the selected range
                        var selectedRange = context.workbook.getSelectedRange();

                        // Clear data validation from the selected range
                        selectedRange.dataValidation.clear();

                        // Sync to apply the changes
                        return context.sync();
                    }).catch(function (error) {
                        console.log(error);
                    });
                };


                async function checkIfSettingsSheet() {
                    return new Promise((resolve, reject) => {
                        try {
                            Excel.run(async function (context) {
                                let sheet = context.workbook.worksheets.getActiveWorksheet();
                                sheet.load("name");

                                await context.sync();

                                resolve(sheet.name === 'Settings');
                            });
                        } catch (error) {
                            console.error("Error:", error);
                            reject(error);
                        }
                    });
                };



                function getIndexForStarCusTag() {
                    return new Promise(async (resolve, reject) => {
                        try {
                            await Excel.run(async (context) => {
                                let sheet = context.workbook.worksheets.getActiveWorksheet();

                                let range = sheet.getRange("A1:Z1");
                                range.load("values");
                                await context.sync();

                             //   console.log(range.values);
                                let headersRow = range.values[0];

                                var lastEmptyValueInx = 1;
                                for (let i = 0; i < headersRow.length; i++) {
                                    if (headersRow[i] === "" && i != 0 && headersRow[i - 1] != "") {
                                        lastEmptyValueInx = i;
                                    }
                                };

                                resolve(lastEmptyValueInx + 1);
                            });
                        } catch (error) {
                            reject(error);
                        }
                    });
                }




                async function setHeadOnWorkingSheet(data) {
                    try {

                        Excel.run(async function (context) {
                            let isSettingsSheet = await checkIfSettingsSheet();

                            if (!isSettingsSheet) {
                                let sheet = context.workbook.worksheets.getActiveWorksheet();
                                let fullSheetVal = sheet.getRange();
                                fullSheetVal.load("values");

                                var OldCustom = window.localStorage.getItem("LastAddress");

                                if (OldCustom != null) {
                                    const modifiedString = OldCustom.replace(/\d+$/, "200");
                                    let usedRangeCustom = sheet.getRange(modifiedString);
                                    usedRangeCustom.clear();
                                }

                                let usedRange = sheet.getUsedRange();
                                usedRange.load("values");



                                await context.sync();
                                let StartFromIndx = await getIndexForStarCusTag();

                                let StartFrom = getAlphabeticCharacter(StartFromIndx);

                                let firstRow = usedRange.values[0];

                                var UsedColumn = [];

                                if (firstRow && firstRow.length > 0) {
                                    for (let columnIndex = 0; columnIndex < firstRow.length; columnIndex++) {
                                        let columnName = firstRow[columnIndex];
                                        UsedColumn.push(columnName);
                                    }
                                }

                                if (data.length > 0) {


                                  //  console.log(StartFrom);
                                    //var StartFrom = getAlphabeticCharacter(UsedColumn.length + 1);
                                    var EndTo = getAlphabeticCharacter((data.length + StartFromIndx) - 1);

                                    window.localStorage.setItem("LastAddress", StartFrom + 1 + ":" + EndTo + 1);
                                    var Address = sheet.getRange(StartFrom + 1 + ":" + EndTo + 1);
                                    Address.values = [data];
                                } else {
                                    window.localStorage.removeItem("LastAddress");
                                }

                                for (var i = 0; i < AllCustoms.length; i++) {
                                    var StartDropdown = getAlphabeticCharacter(StartFromIndx + i) + "1";
                                    //var StartDropdown = StartFrom + "1";
                                    var lengthOfCusArr = AllCustomValues[i].values.length;
                                    let updatedAddress = StartDropdown.replace(/\d+$/, 2);
                                    const cellOfCustom = sheet.getRange(updatedAddress);

                                    let dropdownRange = getAlphabeticCharacter(i + 2) + "2:" + getAlphabeticCharacter(i + 2) + (lengthOfCusArr + 1);

                                    const nameSourceRange = context.workbook.worksheets.getItem("Settings").getRange(dropdownRange);

                                    let approvedListRule = {
                                        list: {
                                            inCellDropdown: true,
                                            source: nameSourceRange
                                        }
                                    };
                                    cellOfCustom.dataValidation.clear();
                                    cellOfCustom.dataValidation.rule = approvedListRule;
                                }

                                await context.sync();
                            } else {
                                loadToast("Go to the working sheet!");
                            }
                        });
                    } catch (error) {
                        console.error("Error:", error);
                    }

                };



                //function modifyString(inputString) {
                //    let splitValues = inputString.split('!');
                //    if (splitValues.length === 2) {
                //        if (splitValues[1].startsWith('A')) {
                //            return 'B' + splitValues[1].substring(1);
                //        } else {
                //            return splitValues[0] + '!B' + splitValues[1].substring(1);
                //        }
                //    } else {
                //        return inputString.replace(/A/g, 'B');
                //    };
                //};


                //async function getUsedRange() {
                //    try {
                //        return await Excel.run(async (context) => {
                //            let sheet = context.workbook.worksheets.getItem("Settings");
                //            let usedRange = sheet.getUsedRange();
                //            usedRange.load("address");

                //            await context.sync();

                //            var used_Range = modifyString(usedRange.address);

                //            console.log("Used Range Address:", usedRange.address);
                //            console.log("Used Range Remove A Address:", used_Range);
                //            return usedRange; // Returning the used range object
                //        });
                //    } catch (error) {
                //        console.error("Error:", error);
                //        return null;
                //    };
                //};



                //async function setData(data) {
                //    try {
                //        await Excel.run(async (context) => {
                //            let sheet = context.workbook.worksheets.getItem("Settings");

                //            //getUsedRange();

                //            sheet.getUsedRange("B:AZ").clear();



                //            let startRow = 1;
                //            let startColumn = 1;


                //            const headers = data.map(item => item.custom);
                //            sheet.getRangeByIndexes(startRow - 1, startColumn, 1, headers.length).values = [headers];

                //            // Populate data
                //            const maxValuesCount = Math.max(...data.map(item => item.values.length));

                //            for (let valueIndex = 0; valueIndex < maxValuesCount; valueIndex++) {
                //                const rowData = [];
                //                data.forEach(item => {
                //                    const value = item.values[valueIndex] ? item.values[valueIndex].tag : '';
                //                    rowData.push(value);
                //                });
                //                sheet.getRangeByIndexes(startRow, startColumn, 1, rowData.length).values = [rowData];
                //                startRow++;
                //            };

                //            // Autofit columns (optional)
                //            sheet.getUsedRange().format.autofitColumns();

                //            await context.sync();
                //        });
                //        //  console.log("Data set successfully.");
                //    } catch (error) {
                //        console.error("Error:", error);
                //    }
                //};






                function setData(data) {
                    Excel.run(function (context) {
                        const sheet = context.workbook.worksheets.getItem('Settings');


                        // sheet.getUsedRange("B:AZ").clear();
                        sheet.getRange("B:Z").clear();


                        data.forEach((item, index) => {
                            const header = item.custom;
                            const value = item.values.map(val => [val.tag]);

                            // Starting from the second column (Column B)
                            const range = sheet.getRangeByIndexes(0, index + 1, value.length + 1, 1);
                            range.values = [[header], ...value];

                            range.format.autofitColumns();
                            range.format.autofitRows();
                        });

                        return context.sync();
                    }).catch(error => {
                        console.error(error);
                    });
                };





                //async function setData(data) {
                //    try {
                //        await Excel.run(async (context) => {
                //            let sheet = context.workbook.worksheets.getItem("Settings");

                //            // Clear the existing content of the sheet
                //            sheet.getUsedRange("B:AZ").clear();

                //            // Start row and column for data insertion
                //            let startRow = 1; // Start from the second row
                //            let startColumn = 1; // Start from the first column

                //            // Set headers (excluding "id")
                //            const headers = data.map(item => item.custom);
                //            sheet.getRangeByIndexes(startRow - 1, startColumn, 1, headers.length).values = [headers];

                //            // Populate data
                //            const maxValuesCount = Math.max(...data.map(item => item.values.length));

                //            for (let valueIndex = 0; valueIndex < maxValuesCount; valueIndex++) {
                //                const rowData = [];
                //                data.forEach(item => {
                //                    const value = item.values[valueIndex] ? item.values[valueIndex].tag : '';
                //                    rowData.push(value);
                //                });
                //                sheet.getRangeByIndexes(startRow, startColumn, 1, rowData.length).values = [rowData];
                //                startRow++;
                //            };

                //            // Autofit columns (optional)
                //            sheet.getUsedRange().format.autofitColumns();

                //            await context.sync();
                //        });
                //        //  console.log("Data set successfully.");
                //    } catch (error) {
                //        console.error("Error:", error);
                //    }
                //};




























                ////////////////// All Column Autofill //////////////////
                function AllSheetAutoFill() {
                    Excel.run(function (context) {

                        let myWorkbook = context.workbook;
                        let sheet = myWorkbook.worksheets.getActiveWorksheet();

                        let range = sheet.getUsedRange();
                        range.format.autofitColumns();

                        return context.sync().then(function () {
                            // console.log("Autofill");
                        }).catch(function (error) {
                            // Handle any errors that occur during context.sync()
                            // console.log("Error: " + error);
                            ProgressLinearInActive();
                            if (error instanceof OfficeExtension.Error && error.code === "InvalidOperationInCellEditMode") {
                                loadToast("Cannot perform this operation while Excel is in editing mode.");
                            } else {
                                loadToast("An error occurred. Please try again later.");
                            };
                        });
                    });
                };

                AllSheetAutoFill();


                //////////////////////// Get tag_templates for dropdown ////////////////////////
                $scope.getTagTemplates = function () {
                    ProgressLinearActive();

                    $.ajax({
                        url: BaseURL + "/wp-json/campaigntrackly/v1/tag_templates",
                        method: "GET",
                        headers: {
                            "accept": "application/json",
                            "Authorization": "Bearer " + APIToken
                        },
                        success: function (response) {
                            //console.log(response);

                            $scope.Tag_TemplatesArr = response;
                            $scope.SelectedOption = "Dummy";

                            ProgressLinearInActive();;
                            if (!$scope.$$phase) {
                                $scope.$apply();
                            };
                        },
                        error: function (error) {
                            //  console.log(error);
                            ProgressLinearInActive();;

                            if (error.responseJSON) {
                                if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                    RefreshToken(getFromLocal.refresh_token);
                                    if (tokenFreshed) {
                                        $scope.getTagTemplates();
                                    };
                                } else {
                                    loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                }
                            } else {
                                loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                            }
                        }
                    });

                };

                function alphaOnly(a) {
                    var b = '';
                    for (var i = 0; i < a.length; i++) {
                        if (a[i] >= 'A' && a[i] <= 'z') b += a[i];
                    }
                    return b;
                };

                function nextLetter(s) {
                    return s.replace(/([a-zA-Z])[^a-zA-Z]*$/, function (a) {
                        var c = a.charCodeAt(0);
                        switch (c) {
                            case 90: return 'A';
                            case 122: return 'a';
                            default: return String.fromCharCode(++c);
                        }
                    });
                };



                ///////////////////////////////////// Apply Template////////////////////////////////////////////


                var OtherTags = [];
                var indxOfCampName;
                var indxOfURL;
                var indxOfContentTag;
                var indxOfSource;
                var indxOfMedium;
                var indxOfTerms;
                var Scenario;
                var AllNameUrlArr = [];
                var CamNameURLObj = {};
                var PrepareFinalArr = [];
                var PrepareDataApplyTemplate = {};
                var FinalSheetSet = [];
                var headerList = [];
                var CustomTagAPI = [];
                var SelctedTemTag = [];
                var ActiveSheet;



                function GetAllCustTags() {
                    return new Promise((resolve, reject) => {
                        $.ajax({
                            url: BaseURL + "/wp-json/campaigntrackly/v1/tag_templates",
                            method: "GET",
                            headers: {
                                "accept": "application/json",
                                "Authorization": "Bearer " + APIToken
                            },
                            success: function (response) {
                                //   console.log(response);


                                for (let i = 0; i < response.length; i++) {
                                    if (response[i].id === $scope.SelectedOption.id) {
                                        SelctedTemTag = response[i].custom;
                                    };
                                };

                                for (var m = 0; m < SelctedTemTag.length; m++) {
                                    if (SelctedTemTag[m].title != null) {
                                        CustomTagAPI.push(SelctedTemTag[m].title.toLowerCase());

                                    }
                                };

                                //   console.log(CustomTagAPI);
                                resolve(response);
                            },
                            error: function (error) {
                                if (error.status != 200 && error.status != 500) {
                                    if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                        RefreshToken(getFromLocal.refresh_token);
                                        ProgressLinearActive();
                                        $scope.ApplyTemplate();
                                    }
                                    else {
                                        loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                    };
                                } else {
                                    loadToast("Connection Issue. Please contact support@campaigntrackly.com");

                                }
                                ProgressLinearInActive();
                            }
                        });

                    });
                }

                function limitStringLength(str) {
                    if (str.length > 20) {
                        return str.slice(0, 20); // Return the first 20 characters
                    } else {
                        return str; // Return the original string if it's already 20 characters or less
                    }
                }

                function getLastItemAfterSplit(inputString) {
                    if (typeof inputString === 'number') {
                        inputString = inputString.toString();
                    }
                    const splitArray = inputString.split('|');
                    if (splitArray.length > 0) {
                        return splitArray[splitArray.length - 1];
                    } else {
                        return '';
                    }
                }


                function closeDialog() {
                    $mdDialog.hide();
                };


                function convertPropertyNamesToLowercase(data) {
                    return data.map(obj => {
                        const newObj = {};
                        for (let key in obj) {
                            newObj[key.toLowerCase()] = obj[key];
                        }
                        return newObj;
                    });
                };

                function removePropertyNamesAndGetValues(data) {
                    return data.map(obj => {
                        return Object.values(obj).map(value => {
                            if (typeof value === 'string' && value.includes('|')) {
                                const splitValues = value.split('|');
                                return splitValues[splitValues.length - 1];
                            }
                            return value;
                        });
                    });
                }




                $scope.ApplyTemplate = async function () {
                    try {
                        ProgressLinearActive(); // Start the loader before making the API call
                        CustomTagAPI = [];
                        SelctedTemTag = [];
                        $scope.UsedSheetValues = [];
                        await GetAllCustTags(); // Wait for the API call to complete
                        // Handle the API response

                        Excel.run(function (context) {
                            let sheetActCall = context.workbook.worksheets.getActiveWorksheet();
                            sheetActCall.load("name");

                            return context.sync().then(function () {
                                // console.log(sheetActCall.name);
                                var ResNameSplit = sheetActCall.name.split("_");
                                if (ResNameSplit[0] === "Result") {
                                    var workingSheetName = ResNameSplit[1];
                                    ProgressLinearInActive();;
                                    loadToast("Please return to " + workingSheetName + " to create new links.");

                                } else {

                                    AllNameUrlArr = [];
                                    OnlyNameArr = [];
                                    CamNameURLObj = {};
                                    PrepareFinalArr = [];
                                    PrepareDataApplyTemplate = {};
                                    FinalSheetSet = [];
                                    AllTagData = [];
                                    LinksOfSncdSca = [];
                                    checkRes = false;
                                    let missingValues = [];
                                    let mappedTagsToColumns;
                                    Excel.run(function (context) {

                                        let myWorkbook = context.workbook;
                                        let sheet = myWorkbook.worksheets.getActiveWorksheet();

                                        let range = sheet.getUsedRange();
                                        //    range.numberFormat = "m/d/yyyy h:mm";
                                        //  range.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

                                        return context.sync().then(function () {
                                            var DataResults = range.load("values");

                                            return context.sync().then(function () {
                                                // console.log(DataResults.values);

                                                allData = DataResults.values;

                                                // console.log(allData);



                                                if (allData[0][0] === '') {
                                                    ProgressLinearInActive();
                                                    return loadToast("Please add Campaign Name and URL labels and data");
                                                };

                                                for (let m = 0; m < allData.length; m++) {

                                                    const allEmptyOrNewline = allData[m].every(item => item === "" || item === "\n");

                                                    if (!allEmptyOrNewline) {
                                                        $scope.UsedSheetValues.push(allData[m]);
                                                    } else {
                                                        //  console.log("All items in the array are equal to empty strings or newline characters.");
                                                    };
                                                };



                                                function isColumnEmpty(columnIndex) {
                                                    for (var i = 1; i < $scope.UsedSheetValues.length; i++) {
                                                        if ($scope.UsedSheetValues[i][columnIndex] !== "") {
                                                            return false;
                                                        }
                                                    }
                                                    return true;
                                                }

                                                // Find and store indices of columns with all empty values
                                                var emptyColumnIndices = [];
                                                for (var j = 0; j < $scope.UsedSheetValues[0].length; j++) {
                                                    if (isColumnEmpty(j)) {
                                                        emptyColumnIndices.push(j);
                                                    }
                                                }



                                                var checkRequiredTagVal = true;


                                                var isEmptyValueFound = false;




                                                function checkValues(usedSheetValues, allCustoms, requiredCustomTags) {
                                                  
                                                    for (let i = 1; i < usedSheetValues.length; i++) {
                                                        const row = usedSheetValues[i];
                                                        for (let j = 0; j < allCustoms.length; j++) {
                                                            const header = allCustoms[j];

                                                            const requiredTag = requiredCustomTags.find(tag => tag.name === header);
                                                            if (requiredTag) {
                                                                const columnIndex = usedSheetValues[0].indexOf(header);
                                                                const value = row[columnIndex];
                                                                if (value === "") {
                                                                    isEmptyValueFound = true;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                        if (isEmptyValueFound) {
                                                            break;
                                                        }
                                                    }

                                                    return isEmptyValueFound;
                                                };


                                                if (isSetting) {
                                                    checkValues($scope.UsedSheetValues, AllCustoms, requiredCusTags);


                                                } else {


                                                    

                                                    for (var j = 0; j < AllCustoms.length; j++) {
                                                        var header = AllCustoms[j];
                                                        isEmptyValueFound = false; // Reset for each column
                                                        for (var i = 1; i < $scope.UsedSheetValues.length; i++) {
                                                            var row = $scope.UsedSheetValues[i];
                                                            var value = row[$scope.UsedSheetValues[0].indexOf(header)];
                                                            if (value === "") {
                                                                isEmptyValueFound = true;
                                                                break; 
                                                            }
                                                        }
                                                        if (!isEmptyValueFound) {
                                                            break;
                                                        }
                                                    }

                                                };






                                                //console.log(requiredCusTags);
                                                //console.log(isEmptyValueFound);

                                                //console.log($scope.UsedSheetValues);
                                                //console.log(AllCustoms);



                                                function mapCustomTagsToColumns(customTagNames, sheetData) {
                                                    const headerRow = sheetData[0];
                                                    const tagMap = {};

                                                    customTagNames.forEach((tag) => {
                                                        const index = headerRow.indexOf(tag.name);
                                                        if (index !== -1) {
                                                            tagMap[tag.id] = index;
                                                        }
                                                    });

                                                    missingValues = [];

                                                    customTagNames.forEach((tag) => {
                                                        const columnIndex = tagMap[tag.id];
                                                        if (columnIndex !== undefined) {
                                                            const tagName = tag.name;
                                                            const hasMissingValue = sheetData.slice(1).some((row) => {
                                                                const cellValue = row[columnIndex];
                                                                const valueToCheck = typeof cellValue === 'number' ? cellValue.toString() : cellValue;
                                                                return valueToCheck.trim ? valueToCheck.trim() === '' : false;
                                                            });

                                                            if (hasMissingValue) {
                                                                missingValues.push(tagName);
                                                            }
                                                        }
                                                    });


                                                    if (missingValues.length > 0) {
                                                        //console.log(`Missing values for tags: ${missingValues.join(', ')}`);
                                                        // Set checkRequiredTagVal to false or perform actions based on your requirements
                                                        checkRequiredTagVal = false;
                                                    }

                                                    const mappedData = sheetData.slice(1).map((row) => {
                                                        const mappedRow = {};

                                                        customTagNames.forEach((tag) => {
                                                            const columnIndex = tagMap[tag.id];
                                                            if (columnIndex !== undefined && columnIndex < row.length) {
                                                                mappedRow[tag.name] = row[columnIndex];
                                                            }
                                                        });

                                                        return mappedRow;
                                                    });

                                                    return mappedData;
                                                }


                                                if (isSetting) {
                                                    mappedTagsToColumns = mapCustomTagsToColumns(requiredCusTags, $scope.UsedSheetValues);
                                                    //   console.log(mappedTagsToColumns);
                                                } else {
                                                    checkRequiredTagVal = true;
                                                };






                                                var lowerCaseHeadArr = $scope.UsedSheetValues[0];

                                                var headerListLow = lowerCaseHeadArr.map(item => item.toLowerCase());

                                                function replaceMultipleSpaces(str) {
                                                    return str.replace(/\s{2,}/g, ' ');
                                                };

                                                function replaceMultipleSpacesInArray(array) {
                                                    return array.map(function (iteme) {
                                                        return replaceMultipleSpaces(iteme);
                                                    });
                                                };

                                                const headerList = replaceMultipleSpacesInArray(headerListLow);


                                                var isCustom = false;


                                                if (isSetting && isEmptyValueFound === false) {
                                                    isCustom = true;

                                                } else {


                                                    if (isEmptyValueFound == false) {
                                                        AllCustoms.forEach(function (item1, index) {
                                                            // Convert item1 and headerList items to lowercase (you can also use toUpperCase)
                                                            const lowerCaseItem1 = item1.toLowerCase();
                                                            //  const lowerCaseHeaderList = headerList.map(item => item.toLowerCase());

                                                            if (headerList.includes(lowerCaseItem1)) {
                                                                isCustom = true;
                                                                return;
                                                            }
                                                        });
                                                    }




                                                }


                                              //  console.log(isCustom);

                                                //if (isSetting) {
                                                //    isCustom = true;
                                                //};




                                                //////////////////////// Check Scenario ////////////////////////

                                                if (headerList.includes("campaign name") && headerList.includes("url") && !headerList.includes('') && !headerList.includes("content") && !headerList.includes("term") && !headerList.includes("source") && !headerList.includes("medium") && isCustom == false) {

                                                    Scenario = "First Scenario";


                                                    for (let i = 0; i < headerList.length; i++) {

                                                        if (headerList[i] === "campaign name") {
                                                            indxOfCampName = i;
                                                        };
                                                        if (headerList[i] === "url") {
                                                            indxOfURL = i;
                                                        };
                                                    };

                                                }
                                                else {
                                                    Scenario = "Secound Scenario";

                                                    OtherTags = [];


                                                    var checkCountCampName = [];
                                                    var objToCamName = {};
                                                    const itemToCheck = "campaign name";

                                                    for (let m = 0; m < headerList.length; m++) {
                                                        if (headerList[m] === itemToCheck) {
                                                            objToCamName = {
                                                                "headName": headerList[m],
                                                                "CampIndx": m
                                                            };
                                                            checkCountCampName.push(objToCamName);
                                                            objToCamName = {};
                                                        };
                                                    };


                                                    for (let i = 0; i < headerList.length; i++) {
                                                        if (headerList[i] === "campaign name") {
                                                            if (checkCountCampName.length === 1) {
                                                                indxOfCampName = i;
                                                            };
                                                            if (i === 0) {
                                                                indxOfCampName = i;
                                                            };
                                                        } else if (headerList[i] === "url") {
                                                            indxOfURL = i;
                                                        } else if (headerList[i] === "content") {
                                                            indxOfContentTag = i;
                                                        } else if (headerList[i] === "medium") {
                                                            indxOfMedium = i;
                                                        } else if (headerList[i] === "term") {
                                                            indxOfTerms = i;
                                                        } else if (headerList[i] === "source") {
                                                            indxOfSource = i;
                                                        } else {
                                                            // if (headerList[i] != "result" && headerList[i] != "short links" && headerList[i] != "date") {
                                                            if (CustomTagAPI.includes(headerList[i])) {
                                                                var CustomTagObj = {
                                                                    "TagName": headerList[i],
                                                                    "TagIndex": i
                                                                };
                                                                OtherTags.push(CustomTagObj);
                                                                CustomTagObj = {};
                                                            };

                                                        };
                                                    };


                                                };

                                                //console.log(OtherTags);

                                                //////////////////////// First Scenario ////////////////////////


                                                if (checkRequiredTagVal != false) {

                                                    //console.log("Everything Okay");
                                                    //ProgressLinearInActive();

                                                    if (Scenario === "First Scenario") {

                                                        for (var n = 1; n < $scope.UsedSheetValues.length; n++) {
                                                            if ($scope.UsedSheetValues[n][indxOfCampName] != "" || $scope.UsedSheetValues[n][indxOfURL] != "") {
                                                                CamNameURLObj = {
                                                                    "CampaignName": $scope.UsedSheetValues[n][indxOfCampName],
                                                                    "CampaignURL": $scope.UsedSheetValues[n][indxOfURL]
                                                                };
                                                                AllNameUrlArr.push(CamNameURLObj);
                                                                CamNameURLObj = {};
                                                            };
                                                        };

                                                        for (let i = 0; i < AllNameUrlArr.length; i++) {
                                                            PrepareDataApplyTemplate = {
                                                                "template_id": $scope.SelectedOption.id,
                                                                "campaign_name": AllNameUrlArr[i].CampaignName,
                                                                "links": [
                                                                    AllNameUrlArr[i].CampaignURL
                                                                ]
                                                            };
                                                            PrepareFinalArr.push(PrepareDataApplyTemplate);
                                                            PrepareDataApplyTemplate = {};
                                                        };



                                                        $.ajax({
                                                            url: BaseURL + "/wp-json/campaigntrackly/v1/apply_template",
                                                            method: "POST",
                                                            headers: {
                                                                "accept": "application/json",
                                                                "Authorization": "Bearer " + APIToken
                                                            },
                                                            data: JSON.stringify(PrepareFinalArr),
                                                            success: function (response) {
                                                                //   console.log(response);



                                                                if (!response[0].hasOwnProperty("existed_links")) {





                                                                    if (response.code) {
                                                                        if (response.code === "401") {
                                                                            ProgressLinearInActive();
                                                                            loadToast(response.response);

                                                                        };
                                                                    };

                                                                    if (response.code != "401") {




                                                                        var dateIndexs = [];
                                                                        var Headers = $scope.UsedSheetValues[0];

                                                                        for (var i = 0; i < Headers.length; i++) {

                                                                            if (Headers[i].toLowerCase().includes("date")) {
                                                                                dateIndexs.push(i);
                                                                            };


                                                                        };
                                                                        //console.log(dateIndexs);
                                                                        //console.log($scope.UsedSheetValues);



                                                                        function getJsDateTimeFromExcel(excelDateValue) {
                                                                            // Convert Excel date to milliseconds since January 1, 1970 (Unix epoch)
                                                                            const msSinceUnixEpoch = (excelDateValue - (25567 + 2)) * 86400 * 1000;

                                                                            // Create a new JavaScript Date object from milliseconds
                                                                            const jsDate = new Date(msSinceUnixEpoch);

                                                                            // Get the hours in UTC (Coordinated Universal Time)
                                                                            const hoursInUTC = jsDate.getUTCHours();

                                                                            // Get the minutes
                                                                            const minutes = jsDate.getMinutes();

                                                                            // Convert hours to your local time zone (modify the offset as needed)
                                                                            const timeZoneOffsetInHours = 0; // Replace with your local time zone offset in hours
                                                                            const hours = (hoursInUTC + timeZoneOffsetInHours) % 24;

                                                                            // Convert hours and minutes to 24-hour format strings
                                                                            const hoursString = hours < 10 ? `0${hours}` : `${hours}`;
                                                                            const minutesString = minutes < 10 ? `0${minutes}` : `${minutes}`;

                                                                            // Format the date and time as "MM/DD/YYYY HH:mm" and return the string
                                                                            return `${jsDate.getMonth() + 1}/${jsDate.getDate()}/${jsDate.getFullYear()} ${hoursString}:${minutesString}`;
                                                                        }



                                                                        function findDateColumnIndex(headerRow) {
                                                                            const dateColumnKeywords = ["DATE", "TIME", "Start Date", "End Date", "Event Date"]; // Add more variations if needed

                                                                            for (let i = 0; i < headerRow.length; i++) {
                                                                                const header = headerRow[i].toUpperCase().trim();
                                                                                if (dateColumnKeywords.some(keyword => header.includes(keyword.toUpperCase()))) {
                                                                                    return i;
                                                                                }
                                                                            }

                                                                            return -1; // Return -1 if the date column is not found
                                                                        }


                                                                        function convertDateColumnToJSDate(dataArray) {
                                                                            const headerRow = dataArray[0];

                                                                            // Find the index of the "DATE" column
                                                                            // const dateColumnIndex = headerRow.findIndex((header) => header === "DATE");


                                                                            const dateColumnIndex = findDateColumnIndex(headerRow);
                                                                            //console.log(dateColumnIndex); // 

                                                                            if (dateColumnIndex === -1) {
                                                                                //      console.error("Date column not found in the data.");
                                                                                return dataArray; // Return the original array if the "DATE" column is not found
                                                                            }

                                                                            // Loop through the array starting from the second row (skipping the header)
                                                                            for (let i = 1; i < dataArray.length; i++) {
                                                                                const excelDate = dataArray[i][dateColumnIndex];

                                                                                // Convert the Excel date to JavaScript Date format
                                                                                const jsDate = getJsDateTimeFromExcel(excelDate);

                                                                                // Replace the Excel date with the JavaScript Date object in the array
                                                                                dataArray[i][dateColumnIndex] = jsDate;
                                                                            }

                                                                            return dataArray;
                                                                        }

                                                                        // Convert the date column to JavaScript Date format
                                                                        const dataArrayWithJSDate = convertDateColumnToJSDate($scope.UsedSheetValues);

                                                                        // console.log(dataArrayWithJSDate);







                                                                        $scope.result_Links = response;

                                                                        if ($scope.result_Links[0].links.length > 0) {

                                                                            FinalSheetSet = [];
                                                                            var UrlItem = [];


                                                                            for (var i = 0; i < $scope.UsedSheetValues.length;) {
                                                                                if (i != 0) {
                                                                                    for (var m = 0; m < $scope.result_Links.length; m++) {
                                                                                        if ($scope.result_Links[m].links.length > 0) {
                                                                                            for (var n = 0; n < $scope.result_Links[m].links.length; n++) {
                                                                                                FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                                                            };
                                                                                            i++;
                                                                                        } else {
                                                                                            FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                                                        };
                                                                                    };
                                                                                } else {
                                                                                    FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                                                    i++
                                                                                };
                                                                            };

                                                                            for (var m = 0; m < $scope.result_Links.length; m++) {
                                                                                if ($scope.result_Links[m].links.length > 0) {
                                                                                    for (var n = 0; n < $scope.result_Links[m].links.length; n++) {
                                                                                        UrlItem.push([$scope.result_Links[m].links[n], $scope.result_Links[m].short_links[n], $scope.result_Links[m].date, $scope.result_Links[m].short_links[n] + "/qr"])
                                                                                    };
                                                                                } else {
                                                                                    UrlItem.push(['', '', $scope.result_Links[m].date]);
                                                                                };
                                                                            };

                                                                            var lastColName = "";
                                                                            HeadNames = $scope.UsedSheetValues[0];
                                                                            var markers = [];

                                                                            for (var n = 0; n < HeadNames.length; n++) {
                                                                                var Aplhabet = (n + 10).toString(36).toUpperCase();
                                                                                markers[i] = sheet.getRange(Aplhabet + 1);
                                                                                markers[i].values = HeadNames[n];
                                                                                if (n < HeadNames.length) {
                                                                                    if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "QR Code") {
                                                                                        lastColName = Aplhabet;
                                                                                    };
                                                                                };
                                                                            };






                                                                            Excel.run(function (context) {
                                                                                let Actsheet = context.workbook.worksheets.getActiveWorksheet();
                                                                                Actsheet.load("name");

                                                                                let sheets = context.workbook.worksheets;
                                                                                sheets.load("items/name");

                                                                                return context.sync().then(function () {

                                                                                    var checkRes;
                                                                                    for (var i = 0; i < sheets.items.length; i++) {
                                                                                        ActiveSheet = Actsheet.name;
                                                                                        ActiveSheet = limitStringLength(ActiveSheet);
                                                                                        var activeSheetRes = "Result_" + ActiveSheet;
                                                                                        if (sheets.items[i].name === activeSheetRes) {
                                                                                            checkRes = true;
                                                                                            break;
                                                                                        } else {
                                                                                            checkRes = false;
                                                                                        };
                                                                                    };


                                                                                    if (checkRes === true) {

                                                                                        let ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);
                                                                                        ProgressLinearInActive();
                                                                                        $scope.OpenDialog();

                                                                                        $scope.SelectMet = function () {

                                                                                            if (Scenario == "First Scenario") {

                                                                                                var argsmessage = $scope.$$childTail.selectMethod;


                                                                                                if (argsmessage === 'Replace') {

                                                                                                    var UsdRangeRes = ResultSheet.getUsedRange();
                                                                                                    context.load(UsdRangeRes);
                                                                                                    UsdRangeRes.clear();

                                                                                                    return context.sync().then(function () {

                                                                                                        Excel.run(function (context) {

                                                                                                            var ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);
                                                                                                            return context.sync().then(function () {

                                                                                                                var NextColumnForResult = nextLetter(lastColName);
                                                                                                                var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                                                                var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                                                                var NextColumnForQr = nextLetter(NextColumnForDate);
                                                                                                                var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForQr + 1);
                                                                                                                rangeForResHead.values = [["Result", "Short Links", "Date", "QR Code"]];

                                                                                                                var rangeForDate = sheet.getRange(NextColumnForDate + ":" + NextColumnForDate); // Replace "A:A" with your desired column range
                                                                                                                rangeForDate.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];


                                                                                                                var toRangeLink = UrlItem.length + 1;
                                                                                                                var range_Link = NextColumnForResult + 2 + ":" + NextColumnForQr + toRangeLink;
                                                                                                                var rangeForResLink = ResultSheet.getRange(range_Link);


                                                                                                                let data = FinalSheetSet;
                                                                                                                var FROM = 1;
                                                                                                                var TO = FROM + data.length - 1;
                                                                                                                var RANEG = "A" + FROM.toString() + ":" + Aplhabet + TO.toString();
                                                                                                                let range = ResultSheet.getRange(RANEG);
                                                                                                                range.formulas = data;
                                                                                                                range.format.autofitColumns();

                                                                                                                var range_LinksRes = NextColumnForResult + 2 + ":" + NextColumnForResult + toRangeLink;
                                                                                                                var rangeValOfLinks = ResultSheet.getRange(range_LinksRes);

                                                                                                                rangeValOfLinks.format.wrapText = true;
                                                                                                                rangeValOfLinks.format.columnWidth = 250;

                                                                                                                ResultSheet.activate();

                                                                                                                return context.sync().then(function () {
                                                                                                                    rangeForResLink.values = UrlItem;
                                                                                                                    rangeForResLink.format.autofitColumns();
                                                                                                                    closeDialog();
                                                                                                                    ProgressLinearInActive();

                                                                                                                });


                                                                                                            });

                                                                                                        });
                                                                                                    });
                                                                                                };
                                                                                                if (argsmessage === 'Merged') {


                                                                                                    Excel.run(function (context) {

                                                                                                        var ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);
                                                                                                        var usedRange = ResultSheet.getUsedRange();

                                                                                                        // Execute the request
                                                                                                        context.load(usedRange);
                                                                                                        return context.sync().then(function () {
                                                                                                            var rowCount = usedRange.rowCount;

                                                                                                            var HeadNames = $scope.UsedSheetValues[0];
                                                                                                            var markers = [];
                                                                                                            var lastColName;
                                                                                                            for (var n = 0; n < HeadNames.length; n++) {
                                                                                                                var Aplhabet = (n + 10).toString(36).toUpperCase();
                                                                                                                markers[i] = Actsheet.getRange(Aplhabet + 1);
                                                                                                                markers[i].values = HeadNames[n];
                                                                                                                if (n < HeadNames.length) {
                                                                                                                    if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "QR Code") {
                                                                                                                        lastColName = Aplhabet;
                                                                                                                    }
                                                                                                                };
                                                                                                            };



                                                                                                            var NextColumnForResult = nextLetter(lastColName);
                                                                                                            var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                                                            var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                                                            var NextColumnForQr = nextLetter(NextColumnForDate);

                                                                                                            var fromRangeLink = rowCount + 1;
                                                                                                            var toRangeLink = fromRangeLink + UrlItem.length - 1;
                                                                                                            var range_Link = NextColumnForResult + fromRangeLink + ":" + NextColumnForQr + toRangeLink;
                                                                                                            var rangeForResLink = ResultSheet.getRange(range_Link);

                                                                                                            var rangeForDate = sheet.getRange(NextColumnForDate + ":" + NextColumnForDate); // Replace "A:A" with your desired column range
                                                                                                            rangeForDate.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];



                                                                                                            FinalSheetSet.shift();
                                                                                                            let data = FinalSheetSet;
                                                                                                            var FROM = rowCount + 1;
                                                                                                            var TO = FROM + data.length - 1;
                                                                                                            var RANEG = "A" + FROM.toString() + ":" + Aplhabet + TO.toString();
                                                                                                            let range = ResultSheet.getRange(RANEG);
                                                                                                            range.formulas = data;
                                                                                                            range.format.autofitColumns();

                                                                                                            var range_LinksRes = NextColumnForResult + fromRangeLink + ":" + NextColumnForResult + toRangeLink;
                                                                                                            var rangeValOfLinks = ResultSheet.getRange(range_LinksRes);

                                                                                                            rangeValOfLinks.format.wrapText = true;
                                                                                                            rangeValOfLinks.format.columnWidth = 250;

                                                                                                            ResultSheet.activate();

                                                                                                            return context.sync().then(function () {
                                                                                                                rangeForResLink.values = UrlItem;
                                                                                                                rangeForResLink.format.autofitColumns();
                                                                                                                closeDialog();
                                                                                                                ProgressLinearInActive();;

                                                                                                            });
                                                                                                        });
                                                                                                    });
                                                                                                };
                                                                                            };
                                                                                        };
                                                                                    } else {
                                                                                        Excel.run(function (context) {

                                                                                            let sheets = context.workbook.worksheets;

                                                                                            let sheet = sheets.add("Result_" + ActiveSheet);
                                                                                            sheet.load("name, position");

                                                                                            return context.sync().then(function () {

                                                                                                let ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);



                                                                                                var NextColumnForResult = nextLetter(lastColName);
                                                                                                var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                                                var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                                                var NextColumnForQr = nextLetter(NextColumnForDate);
                                                                                                var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForQr + 1);
                                                                                                rangeForResHead.values = [["Result", "Short Links", "Date", "QR Code"]];

                                                                                                var rangeForDate = sheet.getRange(NextColumnForDate + ":" + NextColumnForDate); // Replace "A:A" with your desired column range
                                                                                                rangeForDate.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];



                                                                                                var toRangeLink = UrlItem.length + 1;
                                                                                                var range_Link = NextColumnForResult + 2 + ":" + NextColumnForQr + toRangeLink;
                                                                                                var rangeForResLink = ResultSheet.getRange(range_Link);


                                                                                                let data = FinalSheetSet;
                                                                                                var FROM = 1;
                                                                                                var TO = FROM + data.length - 1;
                                                                                                var RANEG = "A" + FROM.toString() + ":" + Aplhabet + TO.toString();
                                                                                                let range = ResultSheet.getRange(RANEG);
                                                                                                range.formulas = data;
                                                                                                range.format.autofitColumns();

                                                                                                var range_LinksRes = NextColumnForResult + 2 + ":" + NextColumnForResult + toRangeLink;
                                                                                                var rangeValOfLinks = ResultSheet.getRange(range_LinksRes);

                                                                                                rangeValOfLinks.format.wrapText = true;
                                                                                                rangeValOfLinks.format.columnWidth = 250;

                                                                                                ResultSheet.activate();

                                                                                                return context.sync().then(function () {
                                                                                                    rangeForResLink.values = UrlItem;
                                                                                                    rangeForResLink.format.autofitColumns();
                                                                                                    ProgressLinearInActive();;

                                                                                                });

                                                                                            });
                                                                                        });
                                                                                    };


                                                                                }).catch(function (error) {
                                                                                    //     console.log(error);

                                                                                });

                                                                            });


                                                                        } else {
                                                                            ProgressLinearInActive();;
                                                                            loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                                        };


                                                                    } else {
                                                                        ProgressLinearInActive();;
                                                                        loadToast(response.response);
                                                                    };

                                                                    if (!$scope.$$phase) {
                                                                        $scope.$apply();
                                                                    };

                                                                } else {

                                                                    ProgressLinearInActive();
                                                                    loadToast("UTM link exists");
                                                                };


                                                            },
                                                            error: function (error) {
                                                                if (error.status != 200 && error.status != 500) {
                                                                    if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                                                        RefreshToken(getFromLocal.refresh_token);
                                                                        ProgressLinearActive();
                                                                        $scope.ApplyTemplate();
                                                                    }
                                                                    else {
                                                                        loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                                    };
                                                                } else {
                                                                    loadToast("Connection Issue. Please contact support@campaigntrackly.com");

                                                                }

                                                                ProgressLinearInActive();;
                                                            }
                                                        });
                                                    };





                                                    //////////////////////// Second Scenario ////////////////////////

                                                    if (Scenario === "Secound Scenario") {


                                                        for (var n = 1; n < $scope.UsedSheetValues.length; n++) {

                                                            if ($scope.UsedSheetValues[n][indxOfCampName] != "" || $scope.UsedSheetValues[n][indxOfURL] != "") {

                                                                CamNameURLObj = {};
                                                                CamNameURLObj["CampaignName"] = ($scope.UsedSheetValues[n][indxOfCampName] ? $scope.UsedSheetValues[n][indxOfCampName] : '');
                                                                CamNameURLObj["CampaignURL"] = ($scope.UsedSheetValues[n][indxOfURL] ? $scope.UsedSheetValues[n][indxOfURL] : '');
                                                                CamNameURLObj["ContentTag"] = ($scope.UsedSheetValues[n][indxOfContentTag] ? $scope.UsedSheetValues[n][indxOfContentTag] : '');
                                                                CamNameURLObj["UtmMedium"] = ($scope.UsedSheetValues[n][indxOfMedium] ? $scope.UsedSheetValues[n][indxOfMedium] : '');
                                                                CamNameURLObj["UtmTerm"] = ($scope.UsedSheetValues[n][indxOfTerms] ? $scope.UsedSheetValues[n][indxOfTerms] : '');
                                                                CamNameURLObj["UtmSource"] = ($scope.UsedSheetValues[n][indxOfSource] ? $scope.UsedSheetValues[n][indxOfSource] : '');

                                                                AllNameUrlArr.push(CamNameURLObj);
                                                                CamNameURLObj = {};
                                                            };
                                                        };

                                                        var OtherTagValArr = [];

                                                        for (var l = 0; l < OtherTags.length; l++) {
                                                            for (var i = 1; i < $scope.UsedSheetValues.length; i++) {
                                                                var OtherTagVal = $scope.UsedSheetValues[i][OtherTags[l].TagIndex];
                                                                var ObjOfOther = {};



                                                                if (OtherTagValArr.length > 0) {
                                                                    if (OtherTags[l].TagName != Object.keys(OtherTagValArr[OtherTagValArr.length - 1])) {
                                                                        ObjOfOther[OtherTags[l].TagName] = [OtherTagVal]
                                                                    } else {
                                                                        var lastIndexTagName = Object.keys(OtherTagValArr[OtherTagValArr.length - 1]);
                                                                        OtherTagValArr[OtherTagValArr.length - 1][lastIndexTagName[0]].push(OtherTagVal);
                                                                        lastIndexTagName = [];
                                                                        ObjOfOther = null;
                                                                    };
                                                                } else {
                                                                    ObjOfOther[OtherTags[l].TagName] = [OtherTagVal]
                                                                };

                                                                if (ObjOfOther != null) {
                                                                    OtherTagValArr.push(ObjOfOther);
                                                                };
                                                            };
                                                        };


                                                        var PreCustTagForSet = [];
                                                        var CustTagForSet = [];


                                                        for (let i = 0; i < OtherTagValArr.length; i++) {
                                                            var keyOfObj = Object.keys(OtherTagValArr[i]);
                                                            var ArrOfTagItem = OtherTagValArr[i][keyOfObj[0]];
                                                            for (let m = 0; m < ArrOfTagItem.length; m++) {
                                                                keyOfObj = Object.keys(OtherTagValArr[i]);
                                                                PreCustTagForSet.push(OtherTagValArr[i][keyOfObj[0]][m]);
                                                                keyOfObj = "";
                                                            };
                                                            CustTagForSet.push(PreCustTagForSet);
                                                            PreCustTagForSet = [];

                                                        };

                                                        var custArr = [];

                                                        // console.log(CustTagForSet);

                                                        //  console.log(requiredCusTagsVal);


                                                        for (let i = 0; i < AllNameUrlArr.length; i++) {


                                                            for (let m = 0; m < CustTagForSet.length; m++) {
                                                                var CusHeadName = [OtherTags[m].TagName];
                                                                if (!CusHeadName[0].includes("date")) {
                                                                    custArr.push({ [OtherTags[m].TagName]: [getLastItemAfterSplit(CustTagForSet[m][i])] });
                                                                } else {
                                                                    var ChangeFormate = CustTagForSet[m][i];
                                                                    custArr.push({ [OtherTags[m].TagName]: [getJsDateFromExcel(ChangeFormate)] });
                                                                };
                                                            };

                                                            //console.log(custArr);
                                                            //console.log(requiredCusTags);



                                                            //var reqCusTagVal = convertPropertyNamesToLowercase(mappedTagsToColumns);
                                                            //console.log(reqCusTagVal);


                                                            let requiredCusTagsVal = [];

                                                            if (isSetting) {
                                                                if (mappedTagsToColumns.length > 0) {
                                                                    requiredCusTagsVal = removePropertyNamesAndGetValues(mappedTagsToColumns);

                                                                };


                                                                PrepareDataApplyTemplate = {
                                                                    "template_id": $scope.SelectedOption.id,
                                                                    "campaign_name": AllNameUrlArr[i].CampaignName,
                                                                    "links": [{
                                                                        "link": AllNameUrlArr[i].CampaignURL,
                                                                        "channels": {
                                                                            "source": AllNameUrlArr[i].UtmSource,
                                                                            "medium": AllNameUrlArr[i].UtmMedium,
                                                                            "terms":
                                                                                (AllNameUrlArr[i].UtmTerm === "" ? [] : [AllNameUrlArr[i].UtmTerm])
                                                                        },
                                                                        "content": AllNameUrlArr[i].ContentTag,
                                                                        "custom": custArr
                                                                    }],
                                                                    "utm_campaign": {
                                                                        "separate": "_",
                                                                        "add_as_campaign_name": "yes",
                                                                        "campaign_convention_id": 0,
                                                                        "tags": requiredCusTagsVal[i]
                                                                    }
                                                                };

                                                            } else {


                                                                PrepareDataApplyTemplate = {
                                                                    "template_id": $scope.SelectedOption.id,
                                                                    "campaign_name": AllNameUrlArr[i].CampaignName,
                                                                    "links": [{
                                                                        "link": AllNameUrlArr[i].CampaignURL,
                                                                        "channels": {
                                                                            "source": AllNameUrlArr[i].UtmSource,
                                                                            "medium": AllNameUrlArr[i].UtmMedium,
                                                                            "terms":
                                                                                (AllNameUrlArr[i].UtmTerm === "" ? [] : [AllNameUrlArr[i].UtmTerm])
                                                                        },
                                                                        "content": AllNameUrlArr[i].ContentTag,
                                                                        "custom": custArr
                                                                    }]
                                                                };

                                                            }



                                                            custArr = [];
                                                            PrepareFinalArr.push(PrepareDataApplyTemplate);
                                                            PrepareDataApplyTemplate = {};
                                                        };


                                                        //console.log(PrepareFinalArr);

                                                        var settings = {
                                                            "url": BaseURL + "/wp-json/campaigntrackly/v1/apply_template_new_tags",
                                                            "method": "POST",
                                                            "timeout": 0,
                                                            "headers": {
                                                                "Accept": "application/json",
                                                                "Content-Type": "application/json",
                                                                "Authorization": "Bearer " + APIToken
                                                            },
                                                            "data": JSON.stringify(PrepareFinalArr),
                                                        };

                                                        $.ajax(settings).done(function (result) {
                                                            // console.log(result);

                                                            if (result.code) {
                                                                if (result.code === "401") {
                                                                    ProgressLinearInActive();
                                                                    loadToast(result.response);

                                                                };
                                                            };


                                                            if (!result[0].hasOwnProperty("existed_links")) {

                                                                

                                                                var dateIndexs = [];
                                                                var Headers = $scope.UsedSheetValues[0];

                                                                for (var i = 0; i < Headers.length; i++) {

                                                                    if (Headers[i].toLowerCase().includes("date")) {
                                                                        dateIndexs.push(i);
                                                                    };


                                                                };

                                                                function getJsDateTimeFromExcel(excelDateValue) {
                                                                    const msSinceUnixEpoch = (excelDateValue - (25567 + 2)) * 86400 * 1000;

                                                                    if (isNaN(msSinceUnixEpoch)) {
                                                                        return excelDateValue;
                                                                    }

                                                                    const jsDate = new Date(msSinceUnixEpoch);

                                                                    const hoursInUTC = jsDate.getUTCHours();


                                                                    const minutes = jsDate.getMinutes();

                                                                    const timeZoneOffsetInHours = 0; // Replace with your local time zone offset in hours
                                                                    const hours = (hoursInUTC + timeZoneOffsetInHours) % 24;

                                                                    const hoursString = hours < 10 ? `0${hours}` : `${hours}`;
                                                                    const minutesString = minutes < 10 ? `0${minutes}` : `${minutes}`;

                                                                    return `${jsDate.getMonth() + 1}/${jsDate.getDate()}/${jsDate.getFullYear()} ${hoursString}:${minutesString}`;
                                                                }



                                                                function findDateColumnIndex(headerRow) {
                                                                    const dateColumnKeywords = ["DATE", "TIME", "Start Date", "End Date", "Event Date"]; // Add more variations if needed

                                                                    for (let i = 0; i < headerRow.length; i++) {
                                                                        const header = headerRow[i].toUpperCase().trim();
                                                                        if (dateColumnKeywords.some(keyword => header.includes(keyword.toUpperCase()))) {
                                                                            return i;
                                                                        }
                                                                    }

                                                                    return -1;
                                                                }


                                                                function convertDateColumnToJSDate(dataArray) {
                                                                    const headerRow = dataArray[0];



                                                                    const dateColumnIndex = findDateColumnIndex(headerRow);

                                                                    if (dateColumnIndex === -1) {
                                                                        //console.error("Date column not found in the data.");
                                                                        return dataArray; // Return the original array if the "DATE" column is not found
                                                                    }

                                                                    for (let i = 1; i < dataArray.length; i++) {
                                                                        const excelDate = dataArray[i][dateColumnIndex];

                                                                        const jsDate = getJsDateTimeFromExcel(excelDate);
                                                                        dataArray[i][dateColumnIndex] = jsDate;
                                                                    }

                                                                    return dataArray;
                                                                }

                                                                const dataArrayWithJSDate = convertDateColumnToJSDate($scope.UsedSheetValues);





                                                                $scope.result_Links = result;

                                                                if (result.code != "401") {

                                                                    if ($scope.result_Links[0].links.length > 0) {

                                                                        FinalSheetSet = [];

                                                                        var UrlItem = [];
                                                                        OnlyNameArr = [];

                                                                        for (var i = 0; i < $scope.UsedSheetValues.length;) {
                                                                            if (i != 0) {
                                                                                for (var m = 0; m < $scope.result_Links.length; m++) {
                                                                                    if ($scope.result_Links[m].links.length > 0) {
                                                                                        //   console.log("Date");
                                                                                        for (var n = 0; n < $scope.result_Links[m].links.length; n++) {
                                                                                            FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                                                        };
                                                                                        i++;

                                                                                    } else {
                                                                                        FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                                                    };
                                                                                };
                                                                            } else {
                                                                                FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                                                i++
                                                                            };
                                                                        };

                                                                        for (var m = 0; m < $scope.result_Links.length; m++) {
                                                                            if ($scope.result_Links[m].links.length > 0) {
                                                                                for (var n = 0; n < $scope.result_Links[m].links.length; n++) {
                                                                                    UrlItem.push([$scope.result_Links[m].links[n], $scope.result_Links[m].short_links[n], $scope.result_Links[m].date, $scope.result_Links[m].short_links[n] + "/qr"])
                                                                                };
                                                                            } else {
                                                                                UrlItem.push(['', '', $scope.result_Links[m].date]);

                                                                            };

                                                                        };


                                                                        //console.log(FinalSheetSet);
                                                                        //console.log(UrlItem);


                                                                        Excel.run(function (context) {
                                                                            var Actsheet = context.workbook.worksheets.getActiveWorksheet();
                                                                            context.load(Actsheet, "name");

                                                                            var sheets = context.workbook.worksheets;
                                                                            context.load(sheets, "items/name");

                                                                            return context.sync().then(function () {
                                                                                var checkRes;
                                                                                for (var i = 0; i < sheets.items.length; i++) {
                                                                                    var ActiveSheet = Actsheet.name;
                                                                                    ActiveSheet = limitStringLength(ActiveSheet);
                                                                                    var activeSheetRes = "Result_" + ActiveSheet;
                                                                                    if (sheets.items[i].name === activeSheetRes) {
                                                                                        checkRes = true;
                                                                                        break;
                                                                                    } else {
                                                                                        checkRes = false;
                                                                                    };
                                                                                };

                                                                                if (checkRes === true) {
                                                                                    var ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);
                                                                                    context.load(ResultSheet);
                                                                                    ProgressLinearInActive();
                                                                                    $scope.OpenDialog();

                                                                                    $scope.SelectMet = function () {
                                                                                        if (Scenario === "Secound Scenario") {
                                                                                            var argsmessage = $scope.$$childTail.selectMethod;

                                                                                            if (argsmessage === 'Replace') {
                                                                                                //  console.log("Replace Button is clicked");
                                                                                                var UsdRangeRes = ResultSheet.getUsedRange();
                                                                                                context.load(UsdRangeRes);
                                                                                                UsdRangeRes.clear();

                                                                                                return context.sync().then(function () {

                                                                                                    Excel.run(function (context) {

                                                                                                        var ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);

                                                                                                        var HeadNames = $scope.UsedSheetValues[0];
                                                                                                        var markers = [];
                                                                                                        var lastColName;
                                                                                                        for (var n = 0; n < HeadNames.length; n++) {
                                                                                                            var Aplhabet = (n + 10).toString(36).toUpperCase();
                                                                                                            markers[i] = Actsheet.getRange(Aplhabet + 1);
                                                                                                            markers[i].values = HeadNames[n];
                                                                                                            if (n < HeadNames.length) {
                                                                                                                if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "QR Code") {
                                                                                                                    lastColName = Aplhabet;
                                                                                                                }
                                                                                                            };
                                                                                                        };

                                                                                                        var NextColumnForResult = nextLetter(lastColName);
                                                                                                        var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                                                        var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                                                        var NextColumnForQr = nextLetter(NextColumnForDate);
                                                                                                        var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForQr + 1);

                                                                                                        var rangeForDate = sheet.getRange(NextColumnForDate + ":" + NextColumnForDate); // Replace "A:A" with your desired column range
                                                                                                        rangeForDate.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

                                                                                                        rangeForResHead.values = [["Result", "Short Links", "Date", "QR Code"]];
                                                                                                        var toRangeLink = UrlItem.length + 1;
                                                                                                        var range_Link = NextColumnForResult + 2 + ":" + NextColumnForQr + toRangeLink;
                                                                                                        var rangeForResLink = ResultSheet.getRange(range_Link);

                                                                                                        let data = FinalSheetSet;
                                                                                                        var FROM = 1;
                                                                                                        var TO = FROM + data.length - 1;
                                                                                                        var RANEG = "A" + FROM.toString() + ":" + Aplhabet + TO.toString();
                                                                                                        let range = ResultSheet.getRange(RANEG);
                                                                                                        range.formulas = data;
                                                                                                        range.format.autofitColumns();

                                                                                                        var range_LinksRes = NextColumnForResult + 2 + ":" + NextColumnForResult + toRangeLink;
                                                                                                        var rangeValOfLinks = ResultSheet.getRange(range_LinksRes);

                                                                                                        rangeValOfLinks.format.wrapText = true;
                                                                                                        rangeValOfLinks.format.columnWidth = 250;

                                                                                                        ResultSheet.activate();

                                                                                                        return context.sync().then(function () {
                                                                                                            rangeForResLink.values = UrlItem;
                                                                                                            rangeForResLink.format.autofitColumns();
                                                                                                            closeDialog();
                                                                                                            ProgressLinearInActive();
                                                                                                        });
                                                                                                    });
                                                                                                });
                                                                                            }
                                                                                            if (argsmessage === 'Merged') {
                                                                                                //    console.log("Merged Button is clicked");

                                                                                                Excel.run(function (context) {

                                                                                                    var ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);
                                                                                                    var usedRange = ResultSheet.getUsedRange();

                                                                                                    // Execute the request
                                                                                                    context.load(usedRange);
                                                                                                    return context.sync().then(function () {
                                                                                                        // Access the used range properties
                                                                                                        var rowCount = usedRange.rowCount;
                                                                                                        // var columnCount = usedRange.columnCount;

                                                                                                        var HeadNames = $scope.UsedSheetValues[0];
                                                                                                        var markers = [];
                                                                                                        var lastColName;
                                                                                                        for (var n = 0; n < HeadNames.length; n++) {
                                                                                                            var Aplhabet = (n + 10).toString(36).toUpperCase();
                                                                                                            markers[i] = Actsheet.getRange(Aplhabet + 1);
                                                                                                            markers[i].values = HeadNames[n];
                                                                                                            if (n < HeadNames.length) {
                                                                                                                if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "QR Code") {
                                                                                                                    lastColName = Aplhabet;
                                                                                                                }
                                                                                                            };
                                                                                                        };

                                                                                                        var NextColumnForResult = nextLetter(lastColName);
                                                                                                        var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                                                        var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                                                        var NextColumnForQr = nextLetter(NextColumnForDate);
                                                                                                        //var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForDate + 1);
                                                                                                        //rangeForResHead.values = [["Result", "Short Links", "Date"]];


                                                                                                        var fromRangeLink = rowCount + 1;
                                                                                                        var toRangeLink = fromRangeLink + UrlItem.length - 1;

                                                                                                        var range_Link = NextColumnForResult + fromRangeLink + ":" + NextColumnForQr + toRangeLink;
                                                                                                        var rangeForResLink = ResultSheet.getRange(range_Link);

                                                                                                        var rangeForDate = sheet.getRange(NextColumnForDate + ":" + NextColumnForDate); // Replace "A:A" with your desired column range
                                                                                                        rangeForDate.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

                                                                                                        FinalSheetSet.shift();

                                                                                                        let data = FinalSheetSet;
                                                                                                        var FROM = rowCount + 1;
                                                                                                        var TO = FROM + data.length - 1;
                                                                                                        var RANEG = "A" + FROM.toString() + ":" + Aplhabet + TO.toString();
                                                                                                        let range = ResultSheet.getRange(RANEG);
                                                                                                        range.formulas = data;
                                                                                                        range.format.autofitColumns();

                                                                                                        var range_LinksRes = NextColumnForResult + fromRangeLink + ":" + NextColumnForResult + toRangeLink;
                                                                                                        var rangeValOfLinks = ResultSheet.getRange(range_LinksRes);

                                                                                                        rangeValOfLinks.format.wrapText = true;
                                                                                                        rangeValOfLinks.format.columnWidth = 250;

                                                                                                        ResultSheet.activate();



                                                                                                        return context.sync().then(function () {

                                                                                                            rangeForResLink.values = UrlItem;
                                                                                                            rangeForResLink.format.autofitColumns();
                                                                                                            closeDialog();
                                                                                                            ProgressLinearInActive();


                                                                                                        });

                                                                                                    });

                                                                                                });
                                                                                            };
                                                                                        };

                                                                                    };

                                                                                } else {


                                                                                    Excel.run(function (context) {

                                                                                        let sheets = context.workbook.worksheets;

                                                                                        let sheet = sheets.add("Result_" + ActiveSheet);
                                                                                        sheet.load("name, position");

                                                                                        return context.sync().then(function () {

                                                                                            let ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);



                                                                                            var HeadNames = $scope.UsedSheetValues[0];
                                                                                            var markers = [];
                                                                                            var lastColName;
                                                                                            for (var n = 0; n < HeadNames.length; n++) {
                                                                                                var Aplhabet = (n + 10).toString(36).toUpperCase();
                                                                                                markers[i] = sheet.getRange(Aplhabet + 1);
                                                                                                markers[i].values = HeadNames[n];
                                                                                                if (n < HeadNames.length) {
                                                                                                    if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "QR Code") {
                                                                                                        lastColName = Aplhabet;
                                                                                                    };
                                                                                                };
                                                                                            };



                                                                                            var NextColumnForResult = nextLetter(lastColName);
                                                                                            var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                                            var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                                            var NextColumnForQr = nextLetter(NextColumnForDate);
                                                                                            var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForQr + 1);

                                                                                            var rangeForDate = sheet.getRange(NextColumnForDate + ":" + NextColumnForDate); // Replace "A:A" with your desired column range
                                                                                            //    rangeForDate.numberFormat = "dd/mm/yyyy";
                                                                                            rangeForDate.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];



                                                                                            rangeForResHead.values = [["Result", "Short Links", "Date", "QR Code"]];
                                                                                            var toRangeLink = UrlItem.length + 1;
                                                                                            var range_Link = NextColumnForResult + 2 + ":" + NextColumnForQr + toRangeLink;
                                                                                            var rangeForResLink = ResultSheet.getRange(range_Link);


                                                                                            let data = FinalSheetSet;
                                                                                            var FROM = 1;
                                                                                            var TO = FROM + data.length - 1;
                                                                                            var RANEG = "A" + FROM.toString() + ":" + Aplhabet + TO.toString();
                                                                                            let range = ResultSheet.getRange(RANEG);
                                                                                            range.formulas = data;
                                                                                            range.format.autofitColumns();

                                                                                            var range_LinksRes = NextColumnForResult + 2 + ":" + NextColumnForResult + toRangeLink;
                                                                                            var rangeValOfLinks = ResultSheet.getRange(range_LinksRes);

                                                                                            rangeValOfLinks.format.wrapText = true;
                                                                                            rangeValOfLinks.format.columnWidth = 250;

                                                                                            ResultSheet.activate();

                                                                                            return context.sync()
                                                                                                .then(function () {
                                                                                                    rangeForResLink.values = UrlItem;
                                                                                                    rangeForResLink.format.autofitColumns();

                                                                                                    //  AllSheetAutoFill();
                                                                                                    ProgressLinearInActive();;
                                                                                                }).catch(function (error) {
                                                                                                    // Handle any errors that occur during context.sync()
                                                                                                    // console.log("Error: " + error);
                                                                                                    ProgressLinearInActive();
                                                                                                    if (error instanceof OfficeExtension.Error && error.code === "InvalidOperationInCellEditMode") {
                                                                                                        loadToast("Cannot perform this operation while Excel is in editing mode.");
                                                                                                    } else {
                                                                                                        loadToast("An error occurred. Please try again later.");
                                                                                                    };
                                                                                                });

                                                                                        }).catch(function (error) {
                                                                                            // Handle any errors that occur during context.sync()
                                                                                            // console.log("Error: " + error);
                                                                                            ProgressLinearInActive();
                                                                                            if (error instanceof OfficeExtension.Error && error.code === "InvalidOperationInCellEditMode") {
                                                                                                loadToast("Cannot perform this operation while Excel is in editing mode.");
                                                                                            } else {
                                                                                                loadToast("An error occurred. Please try again later.");
                                                                                            };
                                                                                        });

                                                                                    });

                                                                                };

                                                                            });

                                                                        });




                                                                    } else {
                                                                        ProgressLinearInActive();
                                                                        loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                                    };


                                                                };



                                                            } else {
                                                                ProgressLinearInActive();
                                                                loadToast("UTM link exists");
                                                            };



                                                        }).fail(function (error) {
                                                            ProgressLinearInActive();
                                                            // console.log(error);
                                                            if (error.status != 200 && error.status != 500) {
                                                                if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                                                    RefreshToken(getFromLocal.refresh_token);
                                                                    ProgressLinearActive();
                                                                    $scope.ApplyTemplate();
                                                                }
                                                                else {
                                                                    loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                                };
                                                            } else {
                                                                loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                            }

                                                            ProgressLinearInActive();
                                                            // console.log(error);
                                                        });







                                                    }



                                                } else {



                                                    loadToast("These tag's required " + missingValues.join(",") + ".");
                                                    ProgressLinearInActive();


                                                };

                                            });

                                        });

                                    });




                                }

                            })
                        }).catch(function (error) {
                            // Handle any errors that occur during context.sync()
                            // console.log("Error: " + error);
                            ProgressLinearInActive();
                            if (error instanceof OfficeExtension.Error && error.code === "InvalidOperationInCellEditMode") {
                                loadToast("Cannot perform this operation while Excel is in editing mode.");
                            } else {
                                loadToast("An error occurred. Please try again later.");
                            };
                        });



                    } catch (error) {
                        // Handle any errors that occurred during the API call
                        //  console.log(error);
                        ProgressLinearInActive();
                    } finally {
                        //   console.log("Finally")
                        //  ProgressLinearInActive(); // Stop the loader after the API call is complete
                    }
                };


                ////////////////////////change Date formate of excel ////////////////////////
                function getJsDateFromExcel(excelDateValue) {

                    var d = new Date((excelDateValue - (25567 + 2)) * 86400 * 1000);
                    month = '' + (d.getMonth() + 1),
                        day = '' + d.getDate(),
                        year = d.getFullYear();

                    if (month.length < 2)
                        month = '0' + month;
                    if (day.length < 2)
                        day = '0' + day;

                    return [month, day, year].join('/');

                };




                ///////////////////////////////// Clear All Sheet /////////////////////////////////
                function ClearSheet() {
                    Excel.run(function (context) {
                        var worksheet = context.workbook.worksheets.getActiveWorksheet();
                        var UsedFormularange = worksheet.getUsedRange();
                        UsedFormularange.clear();
                        return context.sync()
                            .then(function () {
                                // console.log("Clear Sheet")
                            })
                    });
                };

                //////////////////////// Contact to support ////////////////////////

                $scope.ContactSupport = function (ev) {


                    $mdDialog.show(
                        $mdDialog.alert()
                            .parent(angular.element(document.querySelector('#popupContainer')))
                            .clickOutsideToClose(true)
                            .title('Please contact support here')
                            .textContent('support@campaigntrackly.com')
                            .ariaLabel('Alert Dialog Demo')
                            .ok('Got it!')
                            .targetEvent(ev)
                    );
                };




                //////////////////////// Cehck user is logined or not ////////////////////////

                if (APIToken != null) {
                    $scope.LoginDiv = true;
                    $scope.MainPageDiv = false;
                    $scope.NavBarDiv = false;
                    $scope.StartedScreen = true;

                    var getGptToken = window.localStorage.getItem('SecretKey');
                    if (getGptToken != null) {
                        const decryptedKey = decryptAPIKey(getGptToken, 'ChatGPTKey');
                        $scope.ChatGPTKey = decryptedKey;
                    };

                    var isTokenExp = isTokenExpired(APIToken);

                    if (isTokenExp) {
                        ProgressLinearActive();
                        // console.log("Sesion Expired");
                        RefreshToken(getFromLocal.refresh_token);
                        ProgressLinearActive();
                        $scope.getTagTemplates();

                        $scope.LoadSetting();

                    } else {
                        $scope.getTagTemplates();
                        $scope.LoadSetting();
                    };



                    if (!$scope.$$phase) {
                        $scope.$apply();
                    };

                } else {
                    if (FirstTime) {
                        $scope.LoginDiv = true;
                    } else {
                        $scope.LoginDiv = false;
                    };
                    $scope.MainPageDiv = true;
                    $scope.NavBarDiv = true;
                    ProgressLinearInActive();;


                    if (!$scope.$$phase) {
                        $scope.$apply();
                    };
                };


                //////////////////////// Refresh App ////////////////////////
                $scope.RefreshApp = function () {
                    window.location.reload();
                };

                //////////////////////// logout ////////////////////////

                $scope.logOut = function () {
                    $scope.LoginDiv = false;
                    $scope.MainPageDiv = true;
                    $scope.NavBarDiv = true;

                    window.localStorage.removeItem("APIToken");
                    window.localStorage.removeItem("SecretKey");
                };
            } catch (error) {
                console.log(error);
                $scope.LoginDiv = true;
                $scope.StartedScreen = true;
                $scope.MainPageDiv = true;
                ProgressLinearInActive();
                loadToast("Please run in excel");
            };

        });


        $scope.searchTerm = '';
        $scope.clearSearchTerm = function () {
            $scope.searchTerm = '';
        };
        // The md-select directive eats keydown events for some quick select
        // logic. Since we have a search input here, we don't need that logic.
        $element.find('input').on('keydown', function (ev) {
            ev.stopPropagation();
        });



        showActionToast = function (text, word, position) {

            $scope.toastVisible = true;
            $mdToast.show({
                //    scope: $scope.$new(),
                template:

                    '<md-toast><span class="md-caption" flex>' + text + '</span>' +
                    '<md-button class="md-highlight md-caption" style="text-transform:unset;" ng-click="ignore()">Ignore</md-button>' +
                    '<md-button class="md-highlight md-captions" style="text-transform:unset;" ng-click="replace()">Replace</md-button></md-toast>',

                //'<md-toast><span class="md-caption" flex>' + text + '</span>' +
                //'<md-button class="md-highlight"><md-icon class="material-icons">disabled_by_default</md-icon></md-button>' +
                //'<md-button class="md-highlight"><md-icon class="material-icons">sync_alt</md-icon></md-button></md-toast>',

                hideDelay: 7000, // Duration in milliseconds
                position: "bottom right",
                controller: function ($scope, $mdToast) {
                    $scope.ignore = function () {
                        // console.log("Ignored the toast message");
                        $mdToast.hide();
                    };

                    $scope.replace = function () {
                        //   console.log("Replaced the toast message");
                        $mdToast.hide();

                        // $scope.replacedVal = true;

                        Excel.run($scope.$parent.$$childHead.eventResult.context, function (context) {

                            $scope.$parent.$$childHead.eventResult.remove();
                            $scope.$parent.$$childHead.eventResult = null;
                            var sheet = context.workbook.worksheets.getActiveWorksheet();

                            let range = sheet.getRange(position);
                            range.values = [[word]];
                            //range.format.autofitColumns();

                            // Execute the query to load the range values
                            return context.sync()
                                .then(function () {
                                    //  console.log("Value replaced successfully.");
                                    $scope.$parent.$$childHead.eventResult = sheet.onChanged.add($scope.$parent.$$childHead.handleOnChange);
                                    loadToast("All set. Thank you.");
                                    if (!$scope.$$phase) {
                                        $scope.$apply();
                                    }
                                });
                        }).catch(function (error) {
                            // Handle any errors that occur during context.sync()
                            // console.log("Error: " + error);
                            ProgressLinearInActive();
                            if (error instanceof OfficeExtension.Error && error.code === "InvalidOperationInCellEditMode") {
                                loadToast("Cannot perform this operation while Excel is in editing mode.");
                            } else {
                                loadToast("An error occurred. Please try again later.");
                            };
                        });


                    };
                },
            });
        };


    } catch (error) {
        //  console.log(error);
        ProgressLinearActive();
        loadToast("Connection Issue. Please contact support@campaigntrackly.com");
    };



    // showActionToast("We think compeny spelling is incorrect and might need to be fixed, thank you");


    /////////////////// <---- Progress Linear -----> ///////////////////

    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {

            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };
    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };


    /////////////////// <---- Show Message -----> ///////////////////

    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

});