(function () {
  "use strict";

  var messageBanner;
  var ssoToken;
  var msgBody;
  var mailMode;
  var showconfig = false;
  var userStatus = "Igangværende";
  var userCase;
  var userListID = "-1";
  var userCaseName = "-Vælg-";
    var userCategory = "";
    var userCatName = "-Vælg-";
    var outlookclient;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      $(".loader").css("display", "block");
      $("#afailure").css("display", "none");
      getAccessToken();

        $("#drpconfigstatus").change(function (event) {
        $(".loader").css("display", "block");
        getCases(ssoToken, this.value);
        });

        $("#drpstatus").change(function (event) {
            //getStatuses(ssoToken)
            $(".loader").css("display", "block");
            getCases(ssoToken, this.value);
        });

      $("#drpcases").change(function (event) {
        $("#savesection").css("display", "block");
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

        
      $("#btnSave").click(function () {
          $("#afailure").css("display", "none");
          $("#afailure").css("display", "none");
        var mailRecepients = [{
          "displayName": "",
          "emailAddress": "ankerh@emails.itsm360cloud.net"
        }];

        var selectedCase = $("#drpcases").find("option:selected").val();
        var selectedCat = $("#drpcategories").find("option:selected").val();

        if (selectedCase === "-1" || selectedCat ==="-1" || (outlookclient==="Outlook" && msgBody.length===37606)) {
          $("#afailure").text("Please select a case and category or a body is missing").css("display", "block");
          return false;
        }

        if (msgBody.indexOf("###AHC-REF-ID") >= 0) {
          $("#afailure").text("Addin text already part of the body").css("display", "block");
          return false;
        }

        //var newSubject = msgSubject + " ID" + selectedCase + ", Cat" + selectedCat;
        //Office.context.mailbox.item.subject.setAsync(newSubject, function (asyncResult) {
        // if (asyncResult.status === "failed") {
        // console.log("Action failed with error: " + asyncResult.error.message);
        // $("#afailure").text("Case ID not appended to the subject").css("display", "block");
        // } else {
        // console.log("Action Subject appended");
        // Office.context.mailbox.item.bcc.setAsync(mailRecepients, function (result) {
        // if (result.error) {
        // console.log(result.error);
        // $("#afailure").text("Failure while adding the bcc").css("display", "block");
        // } else {
        // console.log("Recipients added to the bcc");
        // Office.context.ui.closeContainer();
        // }
        // });
        // }
        //});

        var bodyhtml = `<div style="margin-left:95%;font-size:1px;display:none"><span hidden>###AHC-REF-ID${selectedCase}-CAT${selectedCat}###</span></div>`;
        var bodytxt = `###AHC-REF-ID${selectedCase}-CAT${selectedCat}###`;
        //console.log("add to body3");
        Office.context.mailbox.item.body.getTypeAsync(function (result) {
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
                function (asyncResult) {
                  if (asyncResult.status ===
                    Office.AsyncResultStatus.Failed) {
                    console.log(asyncResult.error.message);
                  }
                  else {
                    Office.context.mailbox.item.bcc.setAsync(mailRecepients, function (bccresult) {
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
                function (asyncResult) {
                  if (asyncResult.status ===
                    Office.AsyncResultStatus.Failed) {
                    console.log(asyncResult.error.message);
                  }
                  else {
                    this._mail.bcc.setAsync(mailRecepients, function (bccresult) {
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

      $("#btnSaveConfig").click(function () {
        if (userListID === "-1") {

          createUserInfo(ssoToken);
        } else if (userListID > 0) {
          updateUserInfo(ssoToken);

        }
        else {
          $("#afailure").text("Please select valid data").css("display", "block");
          $(".loader").css("display", "none");
        }
      });
        var item = Office.context.mailbox.item;
        outlookclient = Office.context.mailbox.diagnostics.hostName;
      item.body.getAsync(Office.CoercionType.Html, function (result) {
        msgBody = result.value;
          //console.log("mail body ", msgBody);
          //$("#tstdiv").text("test "+msgBody.length);
      });
        
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
                getUserInfo(ssoToken);
                //getCategory(token);
                ////getStatuses(token);
                //getCaseStatuses(token);
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
      $("#drpcases").append('<option value="-1" selected>-Vælg-</option>');
      $("#drpconfigcases").html("");
      $("#drpconfigcases").append('<option value="-1" selected>-Vælg-</option>');
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

  function getStatuses(token) {
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
      $("#drpconfigstatus").append('<option value="-1" selected>-Vælg-</option>');
      $.each(data, function (index, value) {
        $("#drpconfigstatus").append('<option value="' + value + '">' + value + '</option>');
      });
      if (userListID !== "-1") {
        $("#drpconfigstatus").val(status);
      }

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
            //$("#drpconfigstatus").append('<option value="-1" selected>-Vælg-</option>');
            $("#drpstatus").html("");
            //$("#drpstatus").append('<option value="-1" selected>-Vælg-</option>');
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
        $("#savesection").css("display", "block");
          $("#drpstatus").html("");
          $("#drpstatus").append('<option value="' + userStatus + '" selected>' + userStatus + '</option>');
        
        $("#drpconfigstatus").val(userStatus);
        getCases(token, userStatus);
          $("#drpcategories").html("");
          $("#drpcategories").append('<option value="' + userCategory + '" selected>' + userCatName + '</option>');

      } else {
        getCases(token, userStatus);
        }
        getCategory(token);
        //getStatuses(token);
        getCaseStatuses(token);
      $(".loader").css("display", "none");
    }).fail(function (error) {
      console.log("Fail to fetch cases");
      console.log(error);
      $(".loader").css("display", "none");
    });
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
      // getCaseFolders(ssoToken, 1, "drpfolders");
      //$("#drpcategories").html("");
      //$("#drpcategories").append('<option value="' + xCategory + '" selected>' + xCatName + '</option>');
      $("#drpcategories").val(xCategory);
      $("#dvcategory").css("display", "block");
      $("#btnback").css("display", "none");
      $("#btnconfig").css("display", "block");
    // getCaseFolders(ssoToken, 1, "drpfolders");
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
    // getCaseFolders(ssoToken, 1, "drpfolders");
      //$("#drpcategories").html("");
      //$("#drpcategories").append('<option value="' + xCategory + '" selected>' + xCatName + '</option>');
      $("#drpcategories").val(xCategory);
      $("#drpcategories").css("display", "block");
      $("#btnback").css("display", "none");
      $("#btnconfig").css("display", "block");
  }

})();
