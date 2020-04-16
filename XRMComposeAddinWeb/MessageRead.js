(function () {
    "use strict";

    var messageBanner;
    var ssoToken;
    var msgbody;
    var mailMode;
    var caseFolderName;
    var showconfig = false;
    var userStatus = "Igangværende";
    // var userStatus = "-1";
    var userCase;
    var userListID = "-1";
    var userCaseName = "-Vælg-";
    var userCategory = "";
    var userCatName = "-Vælg-";

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $(".loader").css("display", "block");
            getAccessToken();
            checkForInOut();

            $("#drpconfigstatus").change(function (event) {
                //getStatuses(ssoToken)
                $(".loader").css("display", "block");
                getCases(ssoToken, this.value);
            });

            $("#drpstatus").change(function (event) {
                //getStatuses(ssoToken)
                $(".loader").css("display", "block");
                getCases(ssoToken, this.value);
            });

            $("#drpcases").change(function (event) {
                $(".loader").css("display", "block");
                $("#dvSaveEmail").css("display", "block");
                $("#dvSaveAttachments").css("display", "block");
                $("#savesection").css("display", "block");
                $("#chkSaveEmail").prop("checked", true);
                $("#dvcategory").css("display", "block");
                getCaseFolders(ssoToken, 1, "drpfolders");
            });

            $("#drpfolders").change(function (event) {
                $("#drpfolders1").css("display", "block");
                getCaseFolders(ssoToken, 2, "drpfolders1");
            });

            $("#drpfolders1").change(function (event) {
                $("#drpfolders2").css("display", "block");
                getCaseFolders(ssoToken, 3, "drpfolders2");
            });

            $("#drpfolders2").change(function (event) {
                $("#drpfolders3").css("display", "block");
                getCaseFolders(ssoToken, 4, "drpfolders3");
            });

            $("#chkSaveEmail").change(function () {
                $("#drpcategories").val($("#drpcategories option:first").val());
                if ($(this).is(":checked")) {
                    $("#dvcategory").css("display", "block");
                }
                else {
                    $("#dvcategory").css("display", "none");
                }
            });

            $("#chkSaveAttachment").change(function () {
                $("#drpfolders").val($("#drpfolders option:first").val());
                $("#drpfolders1").val($("#drpfolders1 option:first").val());
                $("#drpfolders2").val($("#drpfolders2 option:first").val());
                $("#drpfolders3").val($("#drpfolders3 option:first").val());
                $("#drpfolders1").css("display", "none");
                $("#drpfolders2").css("display", "none");
                $("#drpfolders3").css("display", "none");
                if ($(this).is(":checked")) {
                    $("#dvFolder").css("display", "block");
                }
                else {
                    $("#dvFolder").css("display", "none");
                }
            });

            $("#btnSave").click(function () {
                if ($("#chkSaveEmail").is(":checked")) {
                    saveEmail(ssoToken);
                } else {
                    $("#afailure").text("Please select the Save Email").css("display", "block");
                    $(".loader").css("display", "none");
                }
            });

            $(".btn-light").click(function () {
                if (showconfig) {
                    $("#configcontent").css("display", "none");
                    $("#maincontent").css("display", "block");
                    $("#btnback").css("display", "none");
                    $("#btnconfig").css("display", "block");
                    showconfig = false;
                } else {
                    $("#configcontent").css("display", "block");
                    $("#maincontent").css("display", "none");
                    $("#btnback").css("display", "block");
                    $("#btnconfig").css("display", "none");
                    showconfig = true;
                }
            });


            $("#btnSaveConfig").click(function () {
                if (userListID === "-1") {

                    createUserInfo(ssoToken);
                } else if (userListID > 0) {
                    updateUserInfo(ssoToken);

                } else {
                    $("#afailure").text("Please select valid data").css("display", "block");
                    $(".loader").css("display", "none");
                }
            });

            var item = Office.context.mailbox.item;
            item.body.getAsync('text', function (result) {
                if (result.status === 'succeeded') {
                    msgbody = result.value;
                }
            });

            //authenticator = new OfficeHelpers.Authenticator();
            //authenticator.endpoints.registerMicrosoftAuth(authConfig.clientId, {
            //    redirectUrl: authConfig.redirectUrl,
            //    scope: authConfig.scopes
            //});
        });
    };


    function getAccessToken() {
        if (Office.context.auth !== undefined && Office.context.auth.getAccessTokenAsync !== undefined) {
            Office.context.auth.getAccessTokenAsync({ allowConsentPrompt: true }, function (result) {
                if (result.status === "succeeded") {
                    console.log("token was fetched ");
                    ssoToken = result.value;
                    //getCases(result.value, $("#drpstatus").val());
                    getUserInfo(ssoToken);
                } else if (result.error.code === 13007 || result.error.code === 13005) {
                    console.log("fetching token by force consent");
                    Office.context.auth.getAccessTokenAsync({ allowSignInPrompt: true }, function (result) {
                        if (result.status === "succeeded") {
                            console.log("token was fetched");
                            ssoToken = result.value;
                            //getCases(result.value, $("#drpstatus").val());
                            getUserInfo(ssoToken);
                        }
                        else {
                            console.log("No token was fetched " + result.error.code);
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

    function getUserInfo(token) {
        $(".loader").css("display", "block");
        $.ajax({
            type: "GET",
            url: "api/GetUserDefaultConfig/?useremail=" + Office.context.mailbox.userProfile.emailAddress,
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            console.log("Fetched the User data");
            $.each(data, function (index, value) {
                userStatus = value.StatusID;
                userCase = value.Title;
                userListID = value.ID;
                userCaseName = value.CaseName;
                userCategory = value.Category;
                userCatName = value.CatName;
            });
            if (userListID !== "-1") {

                $("#drpcases").html("");
                $("#drpcases").append('<option value="' + userCase + '" selected>' + userCaseName + '</option>');
                $("#dvSaveEmail").css("display", "block");
                $("#dvSaveAttachments").css("display", "block");
                $("#savesection").css("display", "block");
                $("#chkSaveEmail").prop("checked", true);
                $("#dvcategory").css("display", "block");
                $("#drpstatus").html("");
                $("#drpstatus").append('<option value="' + userStatus + '" selected>' + userStatus + '</option>');
                $("#drpcategories").html("");
                $("#drpcategories").append('<option value="' + userCategory + '" selected>' + userCatName + '</option>');
                getCases(token, userStatus);
                getCaseFolders(ssoToken, 1, "drpfolders");
            } else {
                getCases(token, userStatus);
            }
            getCategory(token);
            getCaseStatuses(token);
            $(".loader").css("display", "none");
        }).fail(function (error) {
            console.log("Fail to fetch cases");
            console.log(error);
            $(".loader").css("display", "none");
        });
    }

    function getCases(token, status) {
        //$(".loader").css("display", "block");
        $.ajax({
            type: "GET",
            url: "api/GetCases/?status=" + status,
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            console.log("Fetched the Cases data");
            $("#drpcases").html("");
            $("#drpcases").append('<option value="" selected>-Vælg-</option>');
            $("#drpconfigcases").html("");
            $("#drpconfigcases").append('<option value="" selected>-Vælg-</option>');
            $.each(data, function (index, value) {
                $("#drpcases").append('<option value="' + value.ID + '">' + value.Title + '</option>');
                $("#drpconfigcases").append('<option value="' + value.ID + '">' + value.Title + '</option>');
            });
            if (userListID !== "-1") {
                $("#drpcases").val(userCase);
                $("#drpconfigcases").val(userCase);
                $("#drpconfigstatus").val(status);
                $("#drpstatus").val(status);
            }
            $(".loader").css("display", "none");
        }).fail(function (error) {
            console.log("Fail to fetch cases");
            console.log(error);
            $(".loader").css("display", "none");
        });
    }

    function getCaseStatuses(token) {
        //$(".loader").css("display", "block");
        $.ajax({
            type: "GET",
            url: "api/GetCaseStatus",
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            console.log("Fetched the Status data");
            $("#drpconfigstatus").html("");
            $("#drpconfigstatus").append('<option value="" selected>-Vælg-</option>');
            $("#drpstatus").html("");
            $("#drpstatus").append('<option value="" selected>-Vælg-</option>');
            $.each(data, function (index, value) {
                $("#drpconfigstatus").append('<option value="' + value + '">' + value + '</option>');
                $("#drpstatus").append('<option value="' + value + '">' + value + '</option>');
            });

            $("#drpconfigstatus").val(userStatus);
            $("#drpstatus").val(userStatus);
        }).fail(function (error) {
            console.log("Fail to fetch cases");
            console.log(error);
            $(".loader").css("display", "none");
        });
    }

    function getCategory(token) {
        //$(".loader").css("display", "block");
        $.ajax({
            type: "GET",
            url: "api/GetCategory",
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            console.log("Fetched the Categories data");
            $("#drpcategories").html("");
            $("#drpcategories").append('<option value="-1" selected>-Vælg-</option>');
            $("#drpconfigcategories").html("");
            $("#drpconfigcategories").append('<option value="-1" selected>-Vælg-</option>');
            $.each(data, function (index, value) {
                $("#drpcategories").append('<option value="' + value.ID + '">' + value.Title + '</option>');
                $("#drpconfigcategories").append('<option value="' + value.ID + '">' + value.Title + '</option>');
            });
            if (userListID !== "-1") {
                $("#drpcategories").val(userCategory);
                $("#drpconfigcategories").val(userCategory);
            }
            $(".loader").css("display", "none");
        }).fail(function (error) {
            console.log("Fail to fetch cases");
            console.log(error);
            $(".loader").css("display", "none");
        });
    }

    function getCaseFolders(token, level, control) {
        console.log("Getting the folders");
        $(".loader").css("display", "block");
        var title = $("#drpcases").find("option:selected").text();
        var id = $("#drpcases").find("option:selected").val();
        var foldername = "";
        if (level === 1) {
            foldername = id;
        }
        else if (level === 2) {
            foldername = caseFolderName + "/" + $("#drpfolders").find("option:selected").text();
        } else if (level === 3) {
            foldername = caseFolderName + "/" + $("#drpfolders").find("option:selected").text() + "/" + $("#drpfolders1").find("option:selected").text();
        } else if (level === 4) {
            foldername = caseFolderName + "/" + $("#drpfolders").find("option:selected").text() + "/" + $("#drpfolders1").find("option:selected").text() + "/" + $("#drpfolders1").find("option:selected").text();
        }

        var caseInfo = {
            Title: $("#drpcases").find("option:selected").text(),
            ID: $("#drpcases").find("option:selected").val(),
            FolderPath: foldername,
            Level: level,
            CaseFolderName: caseFolderName
        };

        $.ajax({
            type: "POST",
            url: "api/GetCaseFolders",
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(caseInfo)
        }).done(function (data) {
            console.log("Fetched the folders");
            //if (level !== 1) {
            //    var select = '<br/><select class="form-control" id="drpfolders' + caseInfo.ID + '"></select>';
            //    $('#dropdown').append(select);
            //    control = "drpfolders"+caseInfo.ID;
            //}
            ////Office.context.ui.closeContainer();
            $("#" + control).html("");
            $("#" + control).append('<option value="" selected>-Vælg-</option>');
            $.each(data, function (index, value) {
                $("#" + control).append('<option value="' + value.Id + '">' + value.Name + '</option>');
                caseFolderName = value.CaseFolderName;
            });
            $(".loader").css("display", "none");
        }).fail(function (error) {
            console.log("Fail to fetch the folders");
            console.log(error);
            $("#" + control).css("display", "none");
            //$("#afailure").text("Failed to fetch the folders").css("display", "block");
            $(".loader").css("display", "none");
        });
    }

    Date.prototype.addHours = function (h) {
        this.setHours(this.getHours() + h);
        return this;
    };

    function saveEmail(token) {
        $(".loader").css("display", "block");
        var item = Office.context.mailbox.item;
        //var gmt = new Date(item.dateTimeCreated).getTimezoneOffset();
        // var datetimecreated = new Date(item.dateTimeCreated).toUTCString();
        var datetimecreated = new Date(item.dateTimeCreated).addHours(9);
        //var datetimecreated = new Date(item.dateTimeModified);
        var emailInfo = {
            Title: item.subject,
            Message: msgbody,
            From: buildEmailAddressString(item.from),
            To: buildEmailAddressesString(item.to),
            CategoryLookupId: $("#drpcategories").find("option:selected").val(),
            RelatedItemListId: "Lists/Cases",
            RelatedItemId: $("#drpcases").find("option:selected").val(),
            Received: datetimecreated,
            ConversationId: item.conversationId,
            ConversationTopic: item.subject,
            InOut: mailMode,
            messageid: Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0)
        };

        $.ajax({
            type: "POST",
            url: "api/SaveEmail",
            headers: {
                "Authorization": "Bearer " + ssoToken
            },
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(emailInfo)
        }).done(function (data) {
            console.log("Saved the Email");
            //Office.context.ui.closeContainer();
            if ($("#chkSaveAttachment").is(":checked")) {
                saveAttachment(ssoToken);
            } else {
                Office.context.ui.closeContainer();
                $(".loader").css("display", "none");
            }
        }).fail(function (error) {
            console.log("Fail to save the email");
            console.log(error);
            $("#afailure").text("Fail to save the email").css("display", "block");
            $(".loader").css("display", "none");
        });


    }

    function saveAttachment(token) {
        var attachments = Office.context.mailbox.item.attachments;
        var attachmentIds = [];
        //if (attachments.length == 0) {
        //    $("#afailure").text("There is no attachment found in the mail.Please unselect the attachment option").css("display", "block");
        //    $(".loader").css("display", "none");
        //    return false;
        //}
        for (var i = 0; i < attachments.length; i++) {
            attachmentIds.push(Office.context.mailbox.convertToRestId(attachments[i].id, Office.MailboxEnums.RestVersion.v2_0));
        }

        var folderpath = $("#drpfolders").find("option:selected").text();
        var level2 = $("#drpfolders1").find("option:selected").val();
        var level3 = $("#drpfolders2").find("option:selected").val();
        var level4 = $("#drpfolders3").find("option:selected").val();

        if (level2.length > 1) {
            folderpath = folderpath + "/" + $("#drpfolders1").find("option:selected").text();
        }

        if (level3.length > 1) {
            folderpath = folderpath + "/" + $("#drpfolders2").find("option:selected").text();
        }

        if (level4.length > 1) {
            folderpath = folderpath + "/" + $("#drpfolders3").find("option:selected").text();
        }

        var attachmentRequest = {
            attachmentIds: attachmentIds,
            messageId: Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0),
            folderName: folderpath,
            caseFolderName: caseFolderName
        };

        $.ajax({
            type: "POST",
            url: "api/SaveAttachment",
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(attachmentRequest)
        }).done(function (data) {
            console.log("Saved the Attachments");
            Office.context.ui.closeContainer();
            $(".loader").css("display", "none");
        }).fail(function (error) {
            console.log("Fail to save the Attachments");
            console.log(error);
            var errormessage = JSON.parse(error.responseText);
            $("#afailure").text(errormessage.Message).css("display", "block");
            $(".loader").css("display", "none");
        });
    }

    // Format an EmailAddressDetails object as
    // GivenName Surname <emailaddress>
    function buildEmailAddressString(address) {
        return address.displayName + ":" + address.emailAddress + ";";
    }

    // Take an array of EmailAddressDetails objects and
    // build a list of formatted strings, separated by a line-break
    function buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                //if (i > 0) {
                //  returnString = returnString;
                //}
                returnString = returnString + buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }

    // Load properties from the Item base object, then load the
    // message-specific properties.
    function loadProps() {
        var item = Office.context.mailbox.item;

        $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
        $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
        $('#itemClass').text(item.itemClass);
        $('#itemId').text(item.itemId);
        $('#itemType').text(item.itemType);

        $('#message-props').show();

        $('#attachments').html(buildAttachmentsString(item.attachments));
        $('#cc').html(buildEmailAddressesString(item.cc));
        $('#conversationId').text(item.conversationId);
        $('#from').html(buildEmailAddressString(item.from));
        $('#internetMessageId').text(item.internetMessageId);
        $('#normalizedSubject').text(item.normalizedSubject);
        $('#sender').html(buildEmailAddressString(item.sender));
        $('#subject').text(item.subject);
        $('#to').html(buildEmailAddressesString(item.to));
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    function checkForInOut() {
        var userprofile = Office.context.mailbox.userProfile;
        var item = Office.context.mailbox.item;
        if (userprofile.emailAddress === item.from.emailAddress) {
            mailMode = "out";
        } else {
            mailMode = "In";
        }
    }


    function createUserInfo(token) {
        // var item = Office.context.mailbox.item;
        //var xcaseid = $("#drpconfigcases").find("option:selected").val();
        //var xcasename = $("#drpconfigcases").find("option:selected").text();
        //var xCategory = $("#drpconfigcategories").find("option:selected").val();
        var xcaseid = $("#drpconfigcases").find("option:selected").val();
        var xcasename = $("#drpconfigcases").find("option:selected").text();
        var xCategory = $("#drpconfigcategories").find("option:selected").val();
        var xCatName = $("#drpconfigcategories").find("option:selected").text();
        var xStatus = $("#drpconfigstatus").find("option:selected").val();
        var userInfo = {
            Title: xcaseid,
            StatusID: xStatus,
            CaseName: xcasename,
            UserMail: Office.context.mailbox.userProfile.emailAddress,
            Category: xCategory,
            CatName: xCatName
        };

        $.ajax({
            type: "POST",
            url: "api/PostUserDefaultConfig",
            headers: {
                "Authorization": "Bearer " + ssoToken
            },
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(userInfo)
        }).done(function (data) {
            userListID = data.ID;
            console.log("Saved the User information");
            //Office.context.ui.closeContainer();

        }).fail(function (error) {
            console.log("Fail to save the user information");
            console.log(error);
            $("#afailure").text("Fail to save the user").css("display", "block");
        });

        $("#configcontent").css("display", "none");
        $("#maincontent").css("display", "block");
        showconfig = false;
        //$("#drpcases").html("");
        //$("#drpcases").append('<option value="' + xcaseid + '" selected>' + xcasename + '</option>');
        $("#drpcases").val(xcaseid);
        $("#dvSaveEmail").css("display", "block");
        $("#dvSaveAttachments").css("display", "block");
        $("#savesection").css("display", "block");
        $("#chkSaveEmail").prop("checked", true);
        //$("#drpstatus").html("");
        //$("#drpstatus").append('<option value="' + xStatus + '" selected>' + xStatus + '</option>');
        $("#drpstatus").val(xStatus);
        getCaseFolders(ssoToken, 1, "drpfolders");
        //$("#drpcategories").html("");
        //$("#drpcategories").append('<option value="' + xCategory + '" selected>' + xCatName + '</option>');
        $("#drpcategories").val(xCategory);
        $("#dvcategory").css("display", "block");
        $("#btnback").css("display", "none");
        $("#btnconfig").css("display", "block");
    }

    function updateUserInfo(token) {
        //$(".loader").css("display", "block");
        // var item = Office.context.mailbox.item;
        var xcaseid = $("#drpconfigcases").find("option:selected").val();
        var xcasename = $("#drpconfigcases").find("option:selected").text();
        var xCategory = $("#drpconfigcategories").find("option:selected").val();
        var xCatName = $("#drpconfigcategories").find("option:selected").text();
        var xStatus = $("#drpconfigstatus").find("option:selected").val();
        var userInfo = {
            Title: xcaseid,
            StatusID: xStatus,
            CaseName: xcasename,
            ID: userListID,
            Category: xCategory,
            CatName: xCatName
        };

        $.ajax({
            type: "PUT",
            url: "api/UpdateUserDefaultConfig",
            headers: {
                "Authorization": "Bearer " + ssoToken
            },
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(userInfo)
        }).done(function (data) {
            console.log("Saved the User information");
            //Office.context.ui.closeContainer();

        }).fail(function (error) {
            console.log("Fail to save the user information");
            console.log(error);
            $("#afailure").text("Fail to save the user").css("display", "block");
            $(".loader").css("display", "none");
        });

        $("#configcontent").css("display", "none");
        $("#maincontent").css("display", "block");
        showconfig = false;
        //$("#drpcases").html("");
        //$("#drpcases").append('<option value="' + xcaseid + '" selected>' + xcasename + '</option>');
        $("#drpcases").val(xcaseid);
        $("#dvSaveEmail").css("display", "block");
        $("#dvSaveAttachments").css("display", "block");
        $("#savesection").css("display", "block");
        $("#chkSaveEmail").prop("checked", true);
        $("#dvcategory").css("display", "block");
        //$("#drpcategories").html("");
        //$("#drpstatus").html("");
        //$("#drpstatus").append('<option value="' + xStatus + '" selected>' + xStatus + '</option>');
        $("#drpstatus").val(xStatus);
        //$("#drpcategories").html("");
        //$("#drpcategories").append('<option value="' + xCategory + '" selected>' + xCatName + '</option>');
        $("#drpcategories").val(xCategory);
        $("#drpcategories").css("display", "block");
        $("#btnback").css("display", "none");
        $("#btnconfig").css("display", "block");
        getCaseFolders(ssoToken, 1, "drpfolders");
    }
})();
