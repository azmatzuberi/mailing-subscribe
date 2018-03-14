/*
    Title: ESA Mailing Subscriptions Admin
    Date: Oct. 06, 2017
    Description: This app allows for the management and admin operations of
        some of the mailing lists on the SharePoint platform. A user can
        add/delete a mailing list and users.
    Author: Azmat Zuberi

*/

//SharePoint core intranet URL
var SharePointContext = {
    'SiteCollection': '/sites/cac-maple/',
    'Site': '/sites/cac-maple/collab/esa/'
};


//Array of user groups for inputs into functions
var arrGroups = ["ESA Service News Subscribers", "ESA Service Health Subscribers"];

//Global variable for userLogin
var userLogin;

//Document load functions
$(document).ready(function() {

    //Get current user login name
    userLogin = $().SPServices.SPGetCurrentUser({
        fieldName: "Name",
        debug: false
    });

    //Two main functions for getting the mailing list and users list
    getMailingList();
    getAllUsers();


    //User: Delete user from both groups
    $(".btn-delete").on("click", function(index) {
        var userName = $(this).attr("user-login");
        for (var i = 0; i < arrGroups.length; i++) {
            console.log(userName);
            $().SPServices({
                operation: "RemoveUserFromGroup",
                groupName: arrGroups[i],
                userLoginName: userName,
                completefunc: function(data, status) {
                }
            });
        }
    });


    //Mailing Group: Delete Mailing Group
    $(document).on('click', '.btn-delete-group', function() {
        var Id = $(this).attr("id");
        console.log(Id);

        $().SPServices({
            operation: "UpdateListItems",
            webURL: SharePointContext.Site,
            batchCmd: "Delete",
            listName: "ESA Mailing Group Subscribers",
            ID: Id,
            completefunc: function(xData, Status) {
                console.log(xData);
            }
        })
    });


    //Change User Groups
    $(document).on('change', '.basic-check', function() {
        var userName = $(this).attr("user-login");
        console.log(userName);

        if (this.checked)
        {
            $().SPServices({
                operation: "AddUserToGroup",
                groupName: "ESA Service News Subscribers",
                userLoginName: userName,
                completefunc: function(data, status) {
                    console.log("success");
                }
            })
        } else {
            $().SPServices({
                operation: "RemoveUserFromGroup",
                groupName: "ESA Service News Subscribers",
                userLoginName: userName,
                completefunc: function(data, status) {
                    console.log("success");
                }
            })
        }
    });
});


//User: Change User groups continues
$(document).on('change', '.advanced-check', function() {
    var userName = $(this).attr("user-login");
    var group = $(this).attr("group");
    console.log(userName);
    if (this.checked)
    {

        $().SPServices({
            operation: "AddUserToGroup",
            groupName: "ESA Service Health Subscribers",
            userLoginName: userName,
            completefunc: function(data, status) {
            }
        })
    } else {
        $().SPServices({
            operation: "RemoveUserFromGroup",
            groupName: "ESA Service Health Subscribers",
            userLoginName: userName,
            completefunc: function(data, status) {
            }
        })
    }
});




//Global variables
var serviceNews;
var serviceHealth;


//Get entire mailing list function
function getMailingList() {

    $.ajax({
        url: SharePointContext.Site + "/_api/web/lists/getbytitle('ESA Mailing Group Subscribers')/items?$orderby=ID asc",
        method: "GET",
        headers: {
            "Accept": "application/json; odata=verbose"
        },
        success: function(data) {
            for (var i = 0; i < data.d.results.length; i++) {
                if (data.d.results[i].ServiceNews == true) {
                    serviceNews = '<td><span>Basic</span><input id="' + data.d.results[i].Id + '" type="checkbox" email="' + data.d.results[i].MailingGroupAddress + '" otherSelect="' + data.d.results[i].ServiceHealth + '" class="basic-group-check" name="' + data.d.results[i].Title + '" checked></td>';
                } else {
                    serviceNews = '<td><span>Basic</span><input id="' + data.d.results[i].Id + '" type="checkbox" email="' + data.d.results[i].MailingGroupAddress + '" otherSelect="' + data.d.results[i].ServiceHealth + '" class="basic-group-check" name="' + data.d.results[i].Title + '"></td>';
                }

                if (data.d.results[i].ServiceHealth == true) {
                    serviceHealth = '<td><span>Advanced</span><input id="' + data.d.results[i].Id + '" type="checkbox" otherSelect="' + data.d.results[i].ServiceNews + '" email="' + data.d.results[i].MailingGroupAddress + '" class="advanced-group-check" name="' + data.d.results[i].Title + '"" checked></td>';
                } else {
                    serviceHealth = '<td><span>Advanced</span><input id="' + data.d.results[i].Id + '" type="checkbox" email="' + data.d.results[i].MailingGroupAddress + '" otherSelect="' + data.d.results[i].ServiceNews + '" name="' + data.d.results[i].Title + '" class="advanced-group-check"></td>';
                }
                var liHtml = "<tr><td>" + data.d.results[i].Title + "</td><td>" + data.d.results[i].MailingGroupAddress + "</td>" + serviceNews + serviceHealth + '<td><button type="button" id="' + data.d.results[i].Id + '" class="btn btn-danger btn-sm btn-delete-group">Delete</button></td></tr>"';
                $("#groupList").append(liHtml);
            }
        },
        error: function(data) {
            console.log(data);
        }
    });
}


//Global variables for user operatoins
var emailNews;
var emailHealth;
var nameNews;
var nameHealth;
var descriptionNews = [];
var descriptionHealth = [];
var userDetailNews = [];
var userDetailHealth = [];
var loginName;

//Get all users function
function getAllUsers()
{

    // Function returns all users in respective Groups                      
    getAllUsersFromNews("ESA Service News Subscribers");

    function getAllUsersFromNews(groupName) {

        $().SPServices({
            operation: "GetUserCollectionFromGroup",
            groupName: groupName,
            async: false,
            completefunc: function(xDataUser, Status) {
                $(xDataUser.responseXML).find("User").each(function() {
                    emailNews = $(this).attr("Email");
                    nameNews = $(this).attr("Name");
                    loginName = $(this).attr("LoginName");
                    descriptionNews = '<tr><td>' + nameNews + '</td><td>' + emailNews + '</td><td>' + '<span style="margin-right: 5px;">Basic</span><input type="checkbox" name="basic" checked group="ESA Service News Subscribers" class="basic-check" user-login="' + loginName + '"></td><td><button type="button" class="btn btn-danger btn-sm btn-delete" user-login="' + loginName + '">Delete</button></td></tr>';
                    userDetailNews.push({
                        name: nameNews,
                        email: emailNews,
                        userLogin: loginName
                    })
                });
            }
        });
    }



    //Print out arrays
    getAllUsersFromHealth("ESA Service Health Subscribers");

    function getAllUsersFromHealth(groupName) {

        // alert(groupName);
        $().SPServices({
            operation: "GetUserCollectionFromGroup",
            groupName: groupName,
            async: false,
            completefunc: function(xDataUser, Status) {
                $(xDataUser.responseXML).find("User").each(function() {
                    emailHealth = $(this).attr("Email");
                    nameHealth = $(this).attr("Name");
                    loginName = $(this).attr("LoginName");
                    descriptionHealth = '<tr><td>' + nameHealth + '</td><td>' + emailHealth + '</td><td><span style="margin-right: 5px;">Advanced</span><input type="checkbox" name="advanced" checked group="ESA Service Health Subscribers" class="advanced-check" user-login="' + loginName + '"><td><button type="button" class="btn btn-danger btn-sm btn-delete" user-login="' + loginName + '">Delete</button></td></tr>';
                    userDetailHealth.push({
                        name: nameHealth,
                        email: emailHealth,
                        userLogin: loginName
                    })
                });
            }
        });
    }


    var healthUsersList = _.filter(userDetailHealth, function(obj) {
        return !_.findWhere(userDetailNews, obj);
    });

    var newsUsersList = _.filter(userDetailNews, function(obj) {
        return !_.findWhere(userDetailHealth, obj);
    });

    for (var i = 0; i < userDetailNews.length; i++)
    {

        for (var j = 0; j < userDetailHealth.length; j++)
        {

            if (userDetailNews[i].email == userDetailHealth[j].email)
            {
                $('#userList').append('<tr><td>' + userDetailNews[i].name + '</td><td>' + userDetailNews[i].email + '</td><td><span style="margin-right: 5px;">Basic</span><input type="checkbox" name="basicd" checked group="ESA Service News Subscribers" class="basic-check" user-login="' + userDetailNews[i].userLogin + '"></td><td><span style="margin-right: 5px;">Advanced</span><input type="checkbox" name="advanced" checked group="ESA Service Health Subscribers" class="advanced-check" user-login="' + userDetailNews[i].userLogin + '"><td></td><td><button type="button" class="btn btn-danger btn-sm btn-delete" user-login="' + userDetailNews[i].userLogin + '">Delete</button></td></tr>');
            }
        }
    }

    for (var i = 0; i < newsUsersList.length; i++)
    {
        $('#userList').append('<tr><td>' + newsUsersList[i].name + '</td><td>' + newsUsersList[i].email + '</td><td><span style="margin-right: 5px;">Basic</span><input type="checkbox" name="basicd" checked group="ESA Service News Subscribers" class="basic-check" user-login="' + newsUsersList[i].userLogin + '"></td><td><span style="margin-right: 5px;">Advanced</span><input type="checkbox" name="advanced" group="ESA Service Health Subscribers" class="advanced-check" user-login="' + newsUsersList[i].userLogin + '"></td><td></td><td><button type="button" class="btn btn-danger btn-sm btn-delete" user-login="' + newsUsersList[i].userLogin + '">Delete</button></td></tr>');
    }

    for (var i = 0; i < healthUsersList.length; i++)
    {
        $('#userList').append('<tr><td>' + healthUsersList[i].name + '</td><td>' + healthUsersList[i].email + '</td><td><span style="margin-right: 5px;">Basic</span><input type="checkbox" name="basic" group="ESA Service News Subscribers" class="basic-check" user-login="' + healthUsersList[i].userLogin + '"></td><td><span style="margin-right: 5px;">Advanced</span><input type="checkbox" name="advanced" checked group="ESA Service Health Subscribers" class="advanced-check" user-login="' + healthUsersList[i].userLogin + '"></td><td></td><td><button type="button" class="btn btn-danger btn-sm btn-delete" user-login="' + healthUsersList[i].userLogin + '">Delete</button></td></tr>');
    }
}


//Add User
$(document).on('click', '#add-user', function() {

    $("#add-user-submit").trigger('click');
    var userEmail = $("#user-name-input").val();
    var userName;
    if ($("#addUser")[0].checkValidity()) {

        $().SPServices({
            operation: "GetUserLoginFromEmail",
            emailXml: "<Users><User Email='" + userEmail + "'/></Users>",
            async: false,
            completefunc: function(xData, Status) {
                $(xData.responseText).find("User").each(function() {
                    if (Status == "Success") {
                        userName = $(this).attr("Login");
                        console.log(userName);
                    } else {
                        console.log(xData);
                        userName = $(this).attr("Login");
                        console.log(userName);
                    }
                })
            }
        });

        for (var i = 0; i < arrGroups.length; i++) {

            $().SPServices({
                operation: "AddUserToGroup",
                groupName: arrGroups[i],
                userLoginName: userName,
                completefunc: function(data, status) {
                    console.log("added");
                }
            });
        }
    }
});


//Add Mail Group
$(document).on('click', '#add-group', function() {

    $("#add-group-submit").trigger('click');
    var title = $("input#titleName").val();
    var mga = $("input#mga").val();
    var serviceNews = $("input[type=checkbox][name=basic-news]:checked").val();
    var serviceNewsValue = (serviceNews == "Yes") ? 1 : 0;
    var serviceHealth = $("input[type=checkbox][name=advanced-health]:checked").val();
    var serviceHealthValue = (serviceHealth == "Yes") ? 1 : 0;
    console.log(title);
    console.log(mga);
    console.log(serviceNewsValue);
    console.log(serviceHealthValue);
    console.log(userLogin);
    var Id = $("#groupList tr:last-child button:last-child").attr("id");
    Id++;
    console.log(Id);

    if ($("#addGroup")[0].checkValidity()) {

        $().SPServices({
            operation: "UpdateListItems",
            async: true,
            batchCmd: "New",
            listName: "ESA Mailing Group Subscribers",
            webURL: SharePointContext.Site,
            valuepairs: [
                ["Title", title],
                ["MailingGroupAddress", mga],
                ["ServiceNews", serviceNewsValue],
                ["ServiceHealth", serviceHealthValue]
            ],
            completefunc: function(xData, Status) {
                if (Status == 'success') {
                    console.log(xData);
                } else {
                    console.log(xData);
                }
            }
        })
    }
});



//Change Mailing Group Level
$(document).on('change', '.basic-group-check', function() {

    var title = $(this).attr("name");
    var mga = $(this).attr("email");
    var idNumber = $(this).attr("id");
    var otherSelect = $(this).attr("otherSelect")
    console.log(otherSelect);
    var serviceHealthValue = (otherSelect == "true") ? 1 : 0;
    console.log(title);
    console.log(mga);
    console.log(idNumber);
    console.log(serviceHealthValue);

    if (this.checked)

    {

        $().SPServices({
            operation: "UpdateListItems",
            async: true,
            batchCmd: "Update",
            listName: "ESA Mailing Group Subscribers",
            webURL: SharePointContext.Site,
            ID: idNumber,
            valuepairs: [
                ["Title", title],
                ["MailingGroupAddress", mga],
                ["ServiceNews", 1],
                ["ServiceHealth", serviceHealthValue]
            ],
            completefunc: function(xData, Status) {
                if (Status == 'success') {
                    console.log(xData);
                } else {
                    console.log(xData);
                }
            }
        })
    } else {

        $().SPServices({
            operation: "UpdateListItems",
            async: true,
            batchCmd: "Update",
            listName: "ESA Mailing Group Subscribers",
            webURL: SharePointContext.Site,
            ID: idNumber,
            valuepairs: [
                ["Title", title],
                ["MailingGroupAddress", mga],
                ["ServiceNews", 0],
                ["ServiceHealth", serviceHealthValue]
            ],
            completefunc: function(xData, Status) {
                if (Status == 'success') {
                    console.log(xData);
                } else {
                    console.log(xData);
                }
            }
        })
    }
});



$(document).on('change', '.advanced-group-check', function() {
    var idNumber = $(this).attr("id");
    var title = $(this).attr("name");
    var mga = $(this).attr("email");
    var otherSelect = $(this).attr("otherSelect");
    var serviceNewsValue = (otherSelect == "true") ? 1 : 0;
    console.log(otherSelect);
    console.log(serviceNewsValue);

    if (this.checked)
    {

        $().SPServices({
            operation: "UpdateListItems",
            async: true,
            batchCmd: "Update",
            listName: "ESA Mailing Group Subscribers",
            webURL: SharePointContext.Site,
            ID: idNumber,
            valuepairs: [
                ["Title", title],
                ["MailingGroupAddress", mga],
                ["ServiceNews", serviceNewsValue],
                ["ServiceHealth", 1]
            ],
            completefunc: function(xData, Status) {
                if (Status == 'success') {
                    console.log(xData);
                } else {
                    console.log(xData);
                }
            }
        })
    } else {

        $().SPServices({
            operation: "UpdateListItems",
            async: true,
            batchCmd: "Update",
            listName: "ESA Mailing Group Subscribers",
            webURL: SharePointContext.Site,
            ID: idNumber,
            valuepairs: [
                ["Title", title],
                ["MailingGroupAddress", mga],
                ["ServiceNews", serviceNewsValue],
                ["ServiceHealth", 0]
            ],
            completefunc: function(xData, Status) {
                if (Status == 'success') {
                    console.log(xData);
                } else {
                    console.log(xData);
                }
            }
        })
    }
});


//Save button: Reloads the page
$('#saveBtn').click(function() {
    location.reload();
});