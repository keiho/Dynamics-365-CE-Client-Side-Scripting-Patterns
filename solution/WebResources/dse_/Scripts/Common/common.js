/*
 * This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
 * THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You
 * a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of
 * the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which
 * the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is
 * embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits,
 * including attorneys? fees, that arise or result from the use or distribution of the Sample Code.
 *
 * Please note: None of the conditions outlined in the disclaimer above will supersede the terms and conditions contained within
 * the Premier Customer Services Description.
 *
 * Extended from CRM SDK References:
 */

var Spark30Common = window.Spark30Common || {};


// Don’t make the call ---------------------------------------------------------------------------------------------

/**
 * @function Spark30Common.AntiUserHasRoleName 
 * @description Return true or false if running user has Security Role by Name passed in parameter
 * @roleName String with exact match of Security Role name in question
 * @return bool
 */
 Spark30Common.AntiUserHasRoleName = function (roleName) 
 {
    "use strict";

    //Returns an array of strings that represent the GUID values of each of the security role or teams that the user is associated with.
    const roleIds = Xrm.Utility.getGlobalContext().userSettings.securityRoles;
    
    //initialize return value
    let hasRole = false;

    //for each security roles ids make a Sync XHR call to get the role name
    roleIds.forEach(
        function(roleId){
            //make a Sync XHR call to get the role name
            var result = Spark30Common.executeSyncRequest("/api/data/v9.0/roles(" + roleId + ")");
            
            //Check passed in value for matches retrieved value
            if (result.name === roleName) {
                //match found set return value to true
                hasRole = true;
            } 
        }
    );
    
    return hasRole;
}

/**
 * @function Spark30Common.UserHasRoleName 
 * @description Return true or false if running user has Security Role by Name passed in parameter
 * @roleName String with exact match of Security Role name in question
 * @return bool
 */
 Spark30Common.UserHasRoleName = function (roleName) {
    "use strict";     

     //get the globalContext
     var globalContext = Xrm.Utility.getGlobalContext();

     //get the roles collection(id, name)
     var userRoles = globalContext.userSettings.roles;

     //initialize return value
     var hasRole = false;

     //for each item in collection
     userRoles.forEach(function hasRoleName(item, index) {

         //Check passed in value for role[].name match 
         if (item.name === roleName) {
             //match found set return value to true
             hasRole = true;
         };
     });
     return hasRole;
 }

// Minimize the number of calls -------------------------------------------------------------------------------------

/**
 * @function Spark30Common.AntiGetConfigurationValue
 * @description Gets requested Configuration Value
 * @configName String with exact match of Configuration name in question
 * @return String with config value
 */
 Spark30Common.AntiGetConfigurationValue = function (configName) {
    "use strict";

    //make a Sync XHR call to get the configuration value matching the string value passed in
    var configResponse = Spark30Common.executeSyncRequest(
        "/api/data/v9.1/dse_systemconfigurations?$select=dse_name,dse_value&$filter=dse_name eq '" + configName + "' and statuscode eq 1");
    
    //initialize return value
    var returnValue = null;
    
    //config setting exist with the given name 
    if(configResponse)
    {
        returnValue = configResponse.value[0]["dse_value"];
    }

    // return configuration value matching name or null if nothing is found
    return returnValue;
}

/**
 * @function Spark30Common.GetConfigurationValue
 * @description Gets requested Configuration Value and 
 *              caches the list of Configuration Values in the browser cache if the list is not already cached to limit repeat calls to the api for future request
 * @configName String with exact match of Configuration name in question
 * @return String with config value
 */
 Spark30Common.GetConfigurationValue = async function (configName) {
    "use strict";

    if (typeof (Storage) !== "undefined") {
        var organizationSettings = Xrm.Utility.getGlobalContext().organizationSettings;
        
        // define a unique session storage name
        var sessionStorageId = organizationSettings.organizationId + ".dse_systemconfiguration"
        
        // if session storage with name does not exist populate it 
        if (!sessionStorage.getItem(sessionStorageId)) {
            
            //Await Async query for the active configuration names and values 
            var configs = await Xrm.WebApi.retrieveMultipleRecords(
                "dse_systemconfiguration", "?$select=dse_name,dse_value&$filter=statuscode eq 1");
            
            //populate the cache with returned config values 
            sessionStorage[sessionStorageId] = JSON.stringify(configs);
        }
        
        // get configuration values from browser cache
        var storedConfig = sessionStorage.getItem(sessionStorageId);

        //initialize return value
        var returnValue = null;
        
        // check for passed in configuration name
        if (storedConfig) {
            var configResponse = JSON.parse(storedConfig);
            for (var i = 0; i < configResponse.entities.length; i++) {
                if (configResponse.entities[i]["dse_name"] == configName) {
                    returnValue = configResponse.entities[i]["dse_value"];
                    break;
                };
            };
        }

        // return configuration value matching name or null if nothing is found
        return returnValue;
    }
}

/**
 * @function Spark30Common.AntiUserIsTeamMember
 * @description Returns true if user is a member of owner team name with the name
 * @teamName String with exact match of Owner Team name in question
 * @return bool
 */
 Spark30Common.AntiUserIsTeamMember = function (teamName) {
    "use strict";
    try{
        
        //initialize return value
        var IsUserExists = false;

        // query for the current users team membership. Filtered to owner teams and the passed in team name
        var teamFetchXml =  "<fetch no-lock='true'>" +
                                "<entity name='team'>" +
                                    "<attribute name='name'/>" +
                                    "<filter>" +
                                        "<condition attribute='teamtype' operator='eq' value='0'/>" +
                                        "<condition attribute='name' operator='eq' value='"+ teamName +"'/>" +
                                    "</filter>" +
                                    "<link-entity name='teammembership' from='teamid' to='teamid' link-type='inner'>" +
                                        "<filter>" +
                                            "<condition attribute='systemuserid' operator='eq-userid'/>" +
                                        "</filter>" +
                                    "</link-entity>" +
                                "</entity>" +
                            "</fetch>"

        //Sync XHR query for Owner Teams 
        var userResponse = Spark30Common.executeSyncRequest("/api/data/v9.1/teams?fetchXml=" + teamFetchXml);
        
        //If a record is returned for the specific owener team query return true
        if(userResponse.value.length>0) {
            IsUserExists = true;
        }

        // returns true is the user is a member of the owner team otherwise false
        return IsUserExists;
    } 
    catch(ex){
        Xrm.Navigation.openAlertDialog({ text: ex });
    }
}

/**
 * @function Spark30Common.UserIsTeamMember
 * @description Returns true if user is a member of owner team name passed in and 
 *              caches the list of user teams names in the browser cache if the list is not already cached to limit repeat calls to the api for future request
 * @teamName String with exact match of Owner Team name in question
 * @return bool
 */
 Spark30Common.UserIsTeamMember = async function (teamName) {
    "use strict";
    try{
        if (typeof (Storage) !== "undefined") {
            var userSettings = Xrm.Utility.getGlobalContext().userSettings;

            // define a unique session storage name
            var sessionStorageId = userSettings.userId + ".teams"

            // if session storage with name does not exist populate it 
            if (!sessionStorage.getItem(sessionStorageId)) {

                // query for all the current users owner teams 
                var teamFetchXml =  "<fetch no-lock='true'>" +
                                        "<entity name='team'>" +
                                            "<attribute name='name'/>" +
                                            "<filter>" +
                                                "<condition attribute='teamtype' operator='eq' value='0'/>" +
                                            "</filter>" +
                                            "<link-entity name='teammembership' from='teamid' to='teamid' link-type='inner'>" +
                                                "<filter>" +
                                                    "<condition attribute='systemuserid' operator='eq-userid'/>" +
                                                "</filter>" +
                                            "</link-entity>" +
                                        "</entity>" +
                                    "</fetch>"
                var apiFetchXml = "?fetchXml=" + teamFetchXml
                
                //Await Async query for Owner Teams 
                var userteams = await Xrm.WebApi.retrieveMultipleRecords("team",apiFetchXml);
                
                //populate cache with owner teams response
                sessionStorage[sessionStorageId] = JSON.stringify(userteams);
            }

            // get configuration values from browser cache
            var storedUserTeams = sessionStorage.getItem(sessionStorageId);

            //initialize return value
            var returnValue = false;

            // check for passed in owner team name
            if (storedUserTeams) {                
                var userTeamResponse = JSON.parse(storedUserTeams);
                
                //return true if user is a member of the owner team passed in
                for (var i = 0; i < userTeamResponse.entities.length; i++) {
                    if (userTeamResponse.entities[i]["name"] === teamName) {
                        returnValue = true;
                        break;
                    };
                };
            }
            
            // return true is the user is a member of the owner team otherwise false
            return returnValue;
        }
    } catch(ex){
        Xrm.Navigation.openAlertDialog({ text: ex });
    }
}

/**
 * @function Spark30Common.AntiUserIsAuditor
 * @description Returns true if running users custom field (systemuser.dse_isauditor) is true 
 * @return bool
 */
 Spark30Common.AntiUserIsAuditor = function () {
    "use strict";
    var userSettings = Xrm.Utility.getGlobalContext().userSettings;

    //Sync XHR query for Current User "dse_isauditor" value
    var userResponse = Spark30Common.executeSyncRequest(
        "/api/data/v9.1/systemusers(" + userSettings.userId.slice(1,-1) + ")?$select=dse_isauditor");
    
    //initialize return value
    var returnValue = false;
    if(userResponse)
    {
        //return true if retrived "dse_isauditor" value equals true
        if (userResponse["dse_isauditor"] === true) {
            returnValue = true;
        };
    }
    return returnValue;
}

/**
 * @function Spark30Common.UserIsAuditor
 * @description Returns true if running users custom field (systemuser.dse_isauditor) is true 
 *              caches the user value in the browser cache if the list is not already cached to limit repeat calls to the api for future request
 * @return bool
 */
 Spark30Common.UserIsAuditor = async function () {
    "use strict";
    if (typeof (Storage) !== "undefined") {

        var userSettings = Xrm.Utility.getGlobalContext().userSettings;

        // define a unique session storage name
        var sessionStorageId = userSettings.userId + ".isAuditor"

        // if session storage with name does not exist populate it 
        if (!sessionStorage.getItem(sessionStorageId)) {

            //Await Async query for Current User "dse_isauditor" value
            var user = await Xrm.WebApi.retrieveRecord("systemuser", userSettings.userId, "?$select=dse_isauditor");

            //cache the value in session storage
            sessionStorage[sessionStorageId] = JSON.stringify(user);
        }

        // get the values from browser cache
        var userRecord = sessionStorage.getItem(sessionStorageId);

        //initialize return value
        var returnValue = false;

        if (userRecord) {

            //return true if cached "dse_isauditor" value equals true
            var userResponse = JSON.parse(userRecord);
            if (userResponse["dse_isauditor"] === true) {
                returnValue = true;
            };
        }

        // return "dse_isauditor" value or false if nothing is found
        return returnValue;
    }
}

// Make it Async ----------------------------------------------------------------------------------------------------

/**
 * @function Spark30Common.AntiIsCreditHoldAccount
 * @description "Credit Hold"
 * @accountId String with the accountid of the Account Record to retrieve
 * @return bool
 */
 Spark30Common.AntiIsCreditHoldAccount = function (accountId) {
    "use strict";

    //Sync XHR call to retrieve account for the passed in accountId
    var accountResponse = Spark30Common.executeSyncRequest(
        "/api/data/v9.1/accounts("+ accountId.slice(1,-1) +")?$select=creditonhold");

    //initialize return value
    var returnValue = false;

    // check respone has a value 
    if(accountResponse)
    {
        //return true if the account is on credit hold
        if (accountResponse["creditonhold"] === true) {
            returnValue = true;
        };
    }
    return returnValue;
}

/**
 * @function Spark30Common.IsCreditHoldAccount
 * @description Returns true if passed in accountid record “Credit Hold” (creditonhold) value is true
 * @accountId String with the accountid of the Account Record to retrieve
 * @return bool
 */
 Spark30Common.IsCreditHoldAccount = async function (accountId) {
    "use strict";

    //Await Async call to retrieve account for the passed in accountId
    let accountResponse = await Xrm.WebApi.retrieveRecord("account", accountId, "?$select=creditonhold");
    
    //initialize return value
    var returnValue = false;

    // check respone has a value 
    if (accountResponse) {

        //return true if the account is on credit hold
        if (accountResponse["creditonhold"] === true) {
            returnValue = true;
        };
    }

    // return 
    return returnValue;
}

// Support functions ------------------------------------------------------------------------------------------------

/**
 * @function Spark30Common.ClearCustomBrowserCache
 * @description removes the sessionStorage item used for caching Teams, Configuration values and isAuditor value
 */
 Spark30Common.ClearCustomBrowserCache = function () {
    "use strict";

    if (typeof (Storage) !== "undefined") {
        var organizationSettings = Xrm.Utility.getGlobalContext().organizationSettings;
        var userSettings = Xrm.Utility.getGlobalContext().userSettings;
        
        // define a unique session storage name
        var sessionStorageId = organizationSettings.organizationId + ".dse_systemconfiguration"

        // remove/clear sessionStorage for Testing 
        sessionStorage.removeItem(sessionStorageId);

        // define a unique session storage name
        sessionStorageId = userSettings.userId + ".teams"

        // remove/clear sessionStorage for Testing 
        sessionStorage.removeItem(sessionStorageId);

        // define a unique session storage name
        sessionStorageId = userSettings.userId + ".isAuditor"
        // remove/clear sessionStorage for Testing 
        sessionStorage.removeItem(sessionStorageId);
    }
}

/**
 * @function Spark30Common.executeSyncRequest 
 * @description Makes a Synchronous XHR request to the passed in requestUrl
 * @requestUrl url represent the table data to query I.E "/api/data/v9.1/systemusers?&select=fullname"
 * @return string - JSON parsed response 
 */
 Spark30Common.executeSyncRequest = function (requestUrl) {
    "use strict";

    try{
        //build the XMLHttp request and set header values
        var request = new XMLHttpRequest();
        
        //setting "false" to force Synchronous
        request.open('GET', requestUrl, false);

        request.setRequestHeader("OData-MaxVersion", "4.0");
        request.setRequestHeader("OData-Version", "4.0");
        request.setRequestHeader("Accept", "application/json");
        request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        request.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
               
        //send the request for the given requestUrl. 
        request.send(null)

        if (request.status === 200) {
        
            //json parse the response and return
            return JSON.parse(request.responseText);
        }
    }
    catch(ex)
    {
        Xrm.Navigation.openAlertDialog({ text: ex });
    }
}