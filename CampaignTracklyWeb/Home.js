var app = angular.module('myApp', ['ngMaterial'], function ($mdThemingProvider) {

    $mdThemingProvider.theme('default')
        .primaryPalette('green', {
            'default': '500',
        });
});
app.controller('myCtrl', function ($scope, $mdToast, $log, $mdDialog, $element) {


    ProgressLinearActive();

    try {

        /////////// Functino for code refresh everytime perfectly ///////////
        function createGuid() {
            return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
                var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
                return v.toString(16);
            });
        };

        var guid = createGuid();


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
        // $scope.replacedVal = false;


        function decryptAPIKey(encryptedKey, encryptionKey) {
            const decrypted = CryptoJS.AES.decrypt(encryptedKey, encryptionKey);
            return decrypted.toString(CryptoJS.enc.Utf8);
        }

        var BaseURL = "https://devapp.campaigntrackly.com";
     // var BaseURL = "https://app.campaigntrackly.com";

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



        var enterEvent = new KeyboardEvent("keydown", {
            key: "Enter",
            code: "Enter",
            keyCode: 13,
            charCode: 13,
            bubbles: true
        });

     

        Office.onReady(function () {

            try {

                // Attach event listener to detect key presses
                //document.addEventListener('keydown', function (event) {
                //    // Handle key press event here
                //    console.log('Key pressed:', event.key);
                //});


                // Press the enter key programmatically

                $scope.testClick = function () {

                    document.dispatchEvent(enterEvent);

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
                                            console.log("Error: " + error);
                                            endGptLoader();
                                        });


                                }).fail(function (error, xhr) {
                                    console.log(error);
                                    endGptLoader();
                                    loadToast("Connection Issue. Please contact support@campaigntrackly.com");

                                });



                            } else {
                                endGptLoader();
                                loadToast("Please put data in cell.");

                            };
                        }).catch(function (error) {
                            console.log("Error occurred during context sync: " + error);
                            endGptLoader();
                            loadToast("Cannot perform this operation while Excel is in editing mode.");
                        });


                    });



                };




                function checkWordOrSentence(input) {
                    // Remove leading and trailing whitespace from the input
                    var trimmedInput = input.trim();

                    // Split the input by spaces to check the number of words
                    var words = trimmedInput.split(" ");

                    if (words.length === 1) {
                        // If the input contains only one word, it is considered a word
                        return "word";
                    } else {
                        // If the input contains multiple words, it is considered a sentence
                        return "sentence";
                    }
                }


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
                                            // underlineCell(adressOfCell);
                                            //   console.log(onlyWord);
                                            //  ProgressLinearInActive();
                                            showActionToast("Spelling might need to be corrected? Thank you", onlyWord, adressOfCell);
                                         } else {
                                            // ProgressLinearInActive();
                                        };
                                    };
                                } else {
                                    //  ProgressLinearInActive();
                                    //  console.log("Already Correct");
                                };



                            }



                        },
                        error: function (error) {
                            console.error('Error:', error);
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
                    var rowNumber = address.split(":")[0].match(/\d+/)[0];


                    if (rowNumber === "1" && eventArgs.details.valueAfter != '') {
                        // ProgressLinearActive();

                      //  console.log(eventArgs.details.valueAfter);
                      //  console.log(eventArgs);

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





                /////////// check user is logined or not ///////////
                var getFromLocal = window.localStorage.getItem("APIToken");
                if (getFromLocal != null) {
                    getFromLocal = JSON.parse(getFromLocal);
                    APIToken = getFromLocal.token;
                };

                /////////// check token expiration ///////////
                function isTokenExpired(token) {
                    const base64Url = token.split(".")[1];
                    const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
                    const jsonPayload = decodeURIComponent(
                        atob(base64)
                            .split("")
                            .map(function (c) {
                                return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
                            })
                            .join("")
                    );

                    const { exp } = JSON.parse(jsonPayload);
                    var expnew = exp * 1000;

                    var ee = new Date(Date.now());
                    var ef = new Date(expnew);
                    if (ee > ef) {
                        expired = true;
                    }
                    else {
                        expired = false;
                    };

                    return expired
                };

                var tokenFreshed;

                //////////////////////// Refresh token ////////////////////////

                function RefreshToken(refshToken) {

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
                            "refresh_token": refshToken
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

             

                //////////////////////// Sign In ////////////////////////
                $scope.SignIn = function () {
                    ProgressLinearActive();


                    var settings = {
                        "url": BaseURL + "/wp-json/campaigntrackly/v1/login",
                        "method": "POST",
                        "timeout": 0,
                        "headers": {
                            "Content-Type": "application/x-www-form-urlencoded",
                        },
                        "data": {
                            "username": $scope.UserName.trim(),
                            "password": $scope.UserPassword.trim()
                        }
                    };

                    $.ajax(settings).done(function (response) {
                        //  console.log(response);

                        if (response.statusCode === 200) {

                          

                            $scope.LoginDiv = true;
                            $scope.MainPageDiv = false;
                            $scope.NavBarDiv = false;
                            if (response.data.token) {
                                APIToken = response.data.token;
                                window.localStorage.setItem("APIToken", JSON.stringify(response.data));




                                $.ajax({
                                    url: BaseURL + "/wp-json/campaigntrackly/v1/gpt_token",
                                    method: "POST",
                                    "timeout": 0,
                                    "headers": {
                                        "Content-Type": "application/x-www-form-urlencoded"
                                    },
                                    "data": {
                                        "username": $scope.UserName.trim(),
                                        "password": $scope.UserPassword.trim()
                                    },
                                    success: function (result) {
                                        console.log(result);
                                        $scope.ChatGPTKey = result.data.token;

                                        const encryptedKey = encryptAPIKey($scope.ChatGPTKey, 'ChatGPTKey');
                                        console.log(encryptedKey);

                                        window.localStorage.setItem('SecretKey', encryptedKey);
                                    },
                                    error: function (error) {
                                   //     console.log(error);
                                        loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                    }
                                })


                                $scope.UserName = undefined;
                                $scope.UserPassword = undefined;

                                $scope.getTagTemplates();

                            };


                        } else {
                            $scope.LoginDiv = false;
                            $scope.MainPageDiv = true;
                            $scope.NavBarDiv = true;
                            // ProgressLinearInActive();;
                        };



                    }).fail(function (error) {
                      //  console.log(error);
                        if (error.responseJSON.statusCode) {

                            if (error.responseJSON.statusCode === 403 || error.responseJSON.code === "application_passwords_disabled") {
                                loadToast(error.responseJSON.message, true);
                            } else {
                                loadToast("Login Failed", true);
                            };
                        };

                        ProgressLinearInActive();;

                    });

                };




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


                //document.getElementById("selectMenu").addEventListener("click", function () {
                //    var myinput = document.getElementById("searchInput");
                //    myinput.focus();
                //});

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
                             //   console.log(error);
                                reject(error);
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

                                if (sheetActCall.name.includes("Result_")) {
                                    var sheetName = sheetActCall;
                                    let fullName = sheetName.name

                                    var workingSheetName = fullName.split("_");

                                    ProgressLinearInActive();;
                                    loadToast("Please return to " + workingSheetName[1] + " to create new links.");

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

                                    Excel.run(function (context) {

                                        let myWorkbook = context.workbook;
                                        let sheet = myWorkbook.worksheets.getActiveWorksheet();

                                        let range = sheet.getUsedRange();


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


                                                //////////////////////// Check Scenario ////////////////////////

                                                if (headerList.includes("campaign name") && headerList.includes("url") && !headerList.includes('') && !headerList.includes("content") && !headerList.includes("term") && !headerList.includes("source") && !headerList.includes("medium")) {

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
                                                    //  console.log(OtherTags);
                                                };

                                                //////////////////////// First Scenario ////////////////////////

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


                                                            if (response.code) {
                                                                if (response.code === "401") {
                                                                    ProgressLinearInActive();;
                                                                    loadToast(response.response);

                                                                };
                                                            };

                                                            if (response.code != "401") {

                                                                $scope.result_Links = response;

                                                                if ($scope.result_Links[0].links.length > 0) {

                                                                    FinalSheetSet = [];
                                                                    var UrlItem = [];

                                                                    //for (let i = 0; i < $scope.result_Links.length; i++) {
                                                                    //    for (let m = 0; m < $scope.result_Links[i].links.length; m++) {
                                                                    //        var ForSheetSet = [AllNameUrlArr[i].CampaignName, AllNameUrlArr[i].CampaignURL, $scope.result_Links[i].links[m], $scope.result_Links[i].short_links[m], $scope.result_Links[i].date];
                                                                    //        FinalSheetSet.push(ForSheetSet);

                                                                    //    };
                                                                    //};


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
                                                                                UrlItem.push([$scope.result_Links[m].links[n], $scope.result_Links[m].short_links[n], $scope.result_Links[m].date])
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
                                                                            if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "Date") {
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

                                                                                var UsdRangeRes = ResultSheet.getUsedRange();
                                                                                UsdRangeRes.clear();

                                                                                return context.sync().then(function () {


                                                                                    var NextColumnForResult = nextLetter(lastColName);
                                                                                    var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                                    var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                                    var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForDate + 1);
                                                                                    rangeForResHead.values = [["Result", "Short Links", "Date"]];
                                                                                    var toRangeLink = UrlItem.length + 1;
                                                                                    var range_Link = NextColumnForResult + 2 + ":" + NextColumnForDate + toRangeLink;
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



                                                                                    //   let sheet = context.workbook.worksheets.getItem("Sheet1");
                                                                                    //   sheet.load("name, position");
                                                                                    ResultSheet.activate();

                                                                                    return context.sync().then(function () {
                                                                                        rangeForResLink.values = UrlItem;
                                                                                        rangeForResLink.format.autofitColumns();
                                                                                        ProgressLinearInActive();;

                                                                                    });


                                                                                });


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
                                                                                        var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForDate + 1);
                                                                                        rangeForResHead.values = [["Result", "Short Links", "Date"]];
                                                                                        var toRangeLink = UrlItem.length + 1;
                                                                                        var range_Link = NextColumnForResult + 2 + ":" + NextColumnForDate + toRangeLink;
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


                                                    for (let i = 0; i < AllNameUrlArr.length; i++) {


                                                        for (let m = 0; m < CustTagForSet.length; m++) {
                                                            var CusHeadName = [OtherTags[m].TagName];
                                                            if (!CusHeadName[0].includes("date")) {
                                                                custArr.push({ [OtherTags[m].TagName]: [CustTagForSet[m][i]] });
                                                            } else {
                                                                var ChangeFormate = CustTagForSet[m][i];
                                                                custArr.push({ [OtherTags[m].TagName]: [getJsDateFromExcel(ChangeFormate)] });
                                                            };
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
                                                            }]
                                                        };




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
                                                                ProgressLinearInActive();;
                                                                loadToast(result.response);

                                                            };
                                                        };


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
                                                                            UrlItem.push([$scope.result_Links[m].links[n], $scope.result_Links[m].short_links[n], $scope.result_Links[m].date])
                                                                        };
                                                                    } else {
                                                                        UrlItem.push(['', '', $scope.result_Links[m].date]);

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

                                                                            var UsdRangeRes = ResultSheet.getUsedRange();
                                                                            UsdRangeRes.clear();


                                                                            var HeadNames = $scope.UsedSheetValues[0];
                                                                            var markers = [];
                                                                            var lastColName;
                                                                            for (var n = 0; n < HeadNames.length; n++) {
                                                                                var Aplhabet = (n + 10).toString(36).toUpperCase();
                                                                                markers[i] = sheet.getRange(Aplhabet + 1);
                                                                                markers[i].values = HeadNames[n];
                                                                                if (n < HeadNames.length) {
                                                                                    if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "Date") {
                                                                                        lastColName = Aplhabet;
                                                                                    };
                                                                                };
                                                                            };



                                                                            var NextColumnForResult = nextLetter(lastColName);
                                                                            var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                            var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                            var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForDate + 1);
                                                                            rangeForResHead.values = [["Result", "Short Links", "Date"]];
                                                                            var toRangeLink = UrlItem.length + 1;
                                                                            var range_Link = NextColumnForResult + 2 + ":" + NextColumnForDate + toRangeLink;
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
                                                                                });


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
                                                                                            if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "Date") {
                                                                                                lastColName = Aplhabet;
                                                                                            };
                                                                                        };
                                                                                    };



                                                                                    var NextColumnForResult = nextLetter(lastColName);
                                                                                    var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                                    var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                                    var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForDate + 1);
                                                                                    rangeForResHead.values = [["Result", "Short Links", "Date"]];
                                                                                    var toRangeLink = UrlItem.length + 1;
                                                                                    var range_Link = NextColumnForResult + 2 + ":" + NextColumnForDate + toRangeLink;
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
                                                                ProgressLinearInActive();;
                                                                loadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                            };


                                                        };



                                                    }).fail(function (error) {
                                                        ProgressLinearInActive();;
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

                                                        ProgressLinearInActive();;
                                                       // console.log(error);
                                                    })

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

                    } else {
                        $scope.getTagTemplates();
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
