(function () {
    "use strict";

    var messageBanner;
    var ssoToken;
    var msgBody;
    var mailMode;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $(".loader").css("display", "block");
            $("#afailure").css("display", "none");
            getAccessToken();

            $("#drpstatus").change(function (event) {
                getCases(ssoToken, this.value);
            });

            $("#drpcases").change(function (event){
                $("#savesection").css("display", "block");
            });

            $("#btnSave").click(function () {
                $("#afailure").css("display", "none");
                var mailRecepients = [{
                    "displayName": "",
                    "emailAddress": "ankerh@emails.itsm360cloud.net"
                }];

                var selectedCase = $("#drpcases").find("option:selected").val();
                var selectedCat = $("#drpcategories").find("option:selected").val();

                if (selectedCase.length <= 0 || selectedCat.length <= 0) {
                    $("#afailure").text("Please select a case and category or a subject is missing").css("display", "block");
                    return false;
                }

                if (msgBody.indexOf("###AHC-REF-ID") >= 0) {
                    $("#afailure").text("Addin text already part of the body").css("display", "block");
                    return false;
                }

                //var newSubject = msgSubject + " ID" + selectedCase + ", Cat" + selectedCat;
                //Office.context.mailbox.item.subject.setAsync(newSubject, function (asyncResult) {
                //    if (asyncResult.status === "failed") {
                //        console.log("Action failed with error: " + asyncResult.error.message);
                //        $("#afailure").text("Case ID not appended to the subject").css("display", "block");
                //    } else {
                //        console.log("Action Subject appended");
                //        Office.context.mailbox.item.bcc.setAsync(mailRecepients, function (result) {
                //            if (result.error) {
                //                console.log(result.error);
                //                $("#afailure").text("Failure while adding the bcc").css("display", "block");
                //            } else {
                //                console.log("Recipients added to the bcc");
                //                Office.context.ui.closeContainer();
                //            }
                //        });
                //    }
                //});

                var bodyhtml = `<div style="margin-left:95%;font-size:1px;display:none"><span hidden>###AHC-REF-ID${selectedCase}-CAT${selectedCat}###</span></div>`;
                var bodytxt = `###AHC-REF-ID${selectedCase}-CAT${selectedCat}###`;
                console.log("add to body3");
                Office.context.mailbox.item.body.getTypeAsync(function(result){
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.log(result.error.message);
                    } else {
                        console.log("email type: ", result.value);
                        if (result.value === 'html') {
                            //this._mail.body.setSelectedDataAsync
                            Office.context.mailbox.item.body.prependAsync(
                                bodyhtml,
                                {
                                    coercionType: Office.CoercionType.Html,
                                    asyncContext: { var3: 1, var4: 2 }
                                },
                                function(asyncResult){
                                    if (asyncResult.status ===
                                        Office.AsyncResultStatus.Failed) {
                                        console.log(asyncResult.error.message);
                                    }
                                    else {
                                        Office.context.mailbox.item.bcc.setAsync(mailRecepients, function(bccresult){
                                            if (bccresult.error) {
                                                console.log(bccresult.error);
                                            } else {
                                                console.log("Recipients added to the bcc");
                                                Office.context.ui.closeContainer();
                                            }
                                        });
                                    }
                                });
                        } else {
                            Office.context.mailbox.item.body.prependAsync(
                                bodytxt,
                                {
                                    coercionType: Office.CoercionType.Text,
                                    asyncContext: { var3: 1, var4: 2 }
                                },
                                function(asyncResult){
                                    if (asyncResult.status ===
                                        Office.AsyncResultStatus.Failed) {
                                        console.log(asyncResult.error.message);
                                    }
                                    else {
                                        this._mail.bcc.setAsync(mailRecepients, function(bccresult){
                                            if (bccresult.error) {
                                                console.log(bccresult.error);
                                            } else {
                                                console.log("Recipients added to the bcc");
                                                Office.context.ui.closeContainer();
                                            }
                                        });
                                    }
                                });
                        }
                    }
                });

            });
            
            var item = Office.context.mailbox.item;
            item.body.getAsync(Office.CoercionType.Html,function (result) {
                msgBody = result.value;
                console.log("mail body ", msgBody);
            });
        });
    };


    function getAccessToken() {
        if (Office.context.auth !== undefined && Office.context.auth.getAccessTokenAsync !== undefined) {
            Office.context.auth.getAccessTokenAsync({ allowConsentPrompt: true },function (result) {
                if (result.status === "succeeded") {
                    console.log("token was fetched ");
                    ssoToken = result.value;
                    getCases(result.value, $("#drpstatus").val());

                } else if (result.error.code === 13007 || result.error.code === 13005) {
                    console.log("fetching token by force consent");
                    Office.context.auth.getAccessTokenAsync({ allowSignInPrompt: true }, function (result) {
                        if (result.status === "succeeded") {
                            console.log("token was fetched");
                            ssoToken = result.value;
                            getCases(result.value, $("#drpstatus").val());
                        }
                        else {
                            console.log("No token was fetched " + result.error.code);
                            //getSiteCollections();
                        }
                    });
                }
                else {
                    console.log("error while fetching access token " + result.error.code);
                    $(".loader").css("display", "none");
                }
            });
        }
    }

    function getCases(token,status) {

        $.ajax({
            type: "GET",
            url: "api/GetCases/?status="+status,
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            console.log("Fetched the Cases data");
            $("#drpcases").html("");
            $("#drpcases").append('<option value="" selected>-Vælg-</option>');
            $.each(data, function (index, value){
                $("#drpcases").append('<option value="' + value.ID + '">' + value.Title + '</option>');
            });
            $(".loader").css("display", "none");
            getCategory(ssoToken);
        }).fail(function (error) {
            console.log("Fail to fetch cases");
            console.log(error);
            $(".loader").css("display", "none");
        });
    }

    function getCategory(token) {
        $(".loader").css("display", "block");
        $.ajax({
            type: "GET",
            url: "api/GetCategory",
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            console.log("Fetched the Cases data");
            $("#drpcategories").html("");
            $("#drpcategories").append('<option value="" selected>-Vælg-</option>');
            $.each(data, function (index, value){
                $("#drpcategories").append('<option value="' + value.ID + '">' + value.Title + '</option>');
            });
            $(".loader").css("display", "none");
        }).fail(function (error) {
            console.log("Fail to fetch cases");
            console.log(error);
            $(".loader").css("display", "none");
        });
    }
})();