(function () {
  "use strict";

    var messageBanner;
    var ssoToken;
    var msgbody;
    var mailMode;
    var caseFolderName;
    var showconfig = false;
    var userStatus = "Igangværende";
    var userCase;
    var userListID="-1";

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
      $(document).ready(function () {
          $(".loader").css("display", "block");
          getAccessToken();
          checkForInOut();

          $("#drpstatus").change(function (event) {
              getCases(ssoToken, this.value);
          });

        $("#drpcases").change(function(event){
            $("#dvSaveEmail").css("display", "block");
            $("#dvSaveAttachments").css("display", "block");
            $("#savesection").css("display", "block");
            $("#chkSaveEmail").prop("checked", true);
            $("#dvcategory").css("display", "block");
            getCaseFolders(ssoToken, 1,"drpfolders");
          });

          $("#drpfolders").change(function (event){
              $("#drpfolders1").css("display", "block");
              getCaseFolders(ssoToken, 2, "drpfolders1");
          });

          $("#drpfolders1").change(function (event){
              $("#drpfolders2").css("display", "block");
              getCaseFolders(ssoToken, 3, "drpfolders2");
          });

          $("#drpfolders2").change(function (event){
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

        $("#chkSaveAttachment").change(function() {
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

          $("#btnconfig").click(function () {
              if (showconfig) {
                  $("#configcontent").css("display", "none");
                  $("#maincontent").css("display", "block");
                  showconfig = false;
              } else {
                  $("#configcontent").css("display", "block");
                  $("#maincontent").css("display", "none");
                  showconfig = true;
              }
          });
          //saveEmail(ssoToken);

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
            Office.context.auth.getAccessTokenAsync({ allowConsentPrompt:true},function (result) {
                if (result.status === "succeeded") {
                    console.log("token was fetched ");
                    ssoToken = result.value;
                    //getCases(result.value, $("#drpstatus").val());
                    getUserInfo(ssoToken);
                    getCategory(ssoToken);

                } else if (result.error.code === 13007 || result.error.code === 13005) {
                    console.log("fetching token by force consent");
                    Office.context.auth.getAccessTokenAsync({ allowSignInPrompt: true }, function (result) {
                        if (result.status === "succeeded") {
                            console.log("token was fetched");
                            ssoToken = result.value;
                            //getCases(result.value, $("#drpstatus").val());
                            getUserInfo(ssoToken);
                            getCategory(ssoToken);
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

    //function getAccessTokenWithPrompt() {
    //    authenticator
    //        .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft, true)
    //        .then(function (token) {
    //            // Get callback token, which grants read access to the current message
    //            // via the Outlook API
    //            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    //                if (result.status === "succeeded") {
    //                    console.log("token was fetched ");
    //                    ssoToken = result.value;
    //                    getCases(result.value);
    //                } else {
    //                    console.log("error while fetching access token " + result.error.code);
    //                    $(".loader").css("display", "none");
    //                }
    //            });
    //        })
    //        .catch(function (error) {
    //            console.log("error while fetching access token " + result.error.code);
    //            $(".loader").css("display", "none");
    //        });
    //}

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
            $("#drpconfigcases").html("");
            $("#drpcases").append('<option value="" selected>-Vælg-</option>');
            $("#drpconfigcases").append('<option value="" selected>-Vælg-</option>');
            $.each(data, function(index, value){
                $("#drpcases").append('<option value="' + value.ID + '">' + value.Title + '</option>');
                $("#drpconfigcases").append('<option value="' + value.ID + '">' + value.Title + '</option>');
            });
            if (userListID !== "-1") {
                $("#drpcases").val(userCase);
                $("#dvSaveEmail").css("display", "block");
                $("#dvSaveAttachments").css("display", "block");
                $("#savesection").css("display", "block");
                $("#chkSaveEmail").prop("checked", true);
                $("#dvcategory").css("display", "block");
                getCaseFolders(ssoToken, 1, "drpfolders");
            }
            $(".loader").css("display", "none");
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
            $("#drpcategories").append('<option value="-1" selected>-Vælg-</option>');
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

    function getCaseFolders(token,level,control) {
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
            //$("#" + control).html("");
            //$("#" + control).append('<option value="" selected>-Vælg-</option>');

            $.each(data, function (index, value){
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
            Message:msgbody,
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
  function getUserInfo(token) {

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
          getCases(token, userStatus);
          $("#drpstatus").val(userStatus);
      });
      //$(".loader").css("display", "none");
    }).fail(function (error) {
      console.log("Fail to fetch cases");
      console.log(error);
      $(".loader").css("display", "none");
    });
  }

  function createUserInfo(token) {
    $(".loader").css("display", "block");
    // var item = Office.context.mailbox.item;
    var userInfo = {
      // Title: item.subject,
      // CategoryLookupId: $("#drpcategories").find("option:selected").val(),
      Title: $("#drpcases").find("option:selected").val(),
      StatusID: $("#drpstatus").find("option:selected").val()

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
      console.log("Saved the User information");
      //Office.context.ui.closeContainer();

    }).fail(function (error) {
      console.log("Fail to save the user information");
      console.log(error);
      $("#afailure").text("Fail to save the user").css("display", "block");
      $(".loader").css("display", "none");
    });


  }

  function updateUserInfo(token) {
    $(".loader").css("display", "block");
    // var item = Office.context.mailbox.item;
    var userInfo = {
      // Title: item.subject,
      // CategoryLookupId: $("#drpcategories").find("option:selected").val(),
      // ID: $("#drpcases").find("option:selected").val(),
      Title: $("#drpcases").find("option:selected").val(),
      StatusID: $("#drpstatus").find("option:selected").val()

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
  }
})();
