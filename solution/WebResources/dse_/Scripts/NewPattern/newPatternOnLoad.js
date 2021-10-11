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

var NewPatternForm = window.NewPatternForm || {};

/**
 * @function NewPatternForm.OnLoad
 * @description Configured in form properties to trigger on form load
 */
NewPatternForm.OnLoad = async function (executionContext) {
    "use strict"
    var formContext = executionContext.getFormContext(); 

    //get the account lookup control from the formcontext 
    var accountControl = formContext.getControl("dse_accountid");

    //makes a non-UCI blocking block call to show "Auditor" Section based on some evaluation of values
    NewPatternForm.CheckIsAuditor(formContext);
    
    //Returns a Promise. determine if the current user is a member of the "Spark" owner team. Caches value in sessionStorage after first retrieval. Cache is used on subsequent calls
    var checkTeamMemberPromise = NewPatternForm.CheckTeamMember(accountControl);
    
    //Returns a Promise. Function checks "checkAccountCreditHold" config value and retrieves the related account "creditonhold". If account on credit hold a notification is displayed on the form
    var checkCreditHoldPromise = NewPatternForm.CheckCreditHold(formContext);

    //Resolve both promises in parallel. Then return out of onload
    return Promise.all([checkTeamMemberPromise, checkCreditHoldPromise]);
}

/**
 * @function NewPatternForm.CheckTeamMember
 * @description checks if user is a member of Spark Team and has Salesperson Role. Disables Related Account Lookup if both are not true
 */
NewPatternForm.CheckTeamMember = async function (accountControl) {
    "use strict";

    //Makes Async call to determine if the "Spark Team" button feature is enabled. Caches value in sessionStorage after first retrieval. Cache is used on subsequent calls
    var isUserOnSparkTeam = await Spark30Common.UserIsTeamMember("Spark");

    //determine if the current user has the "Salesperson" security role. Uses the roles collection(id, name) provided in the Model driven app scripting
    var isUserSalesperson =  Spark30Common.UserHasRoleName("Salesperson");

    //Disables the "Account" lookup field if the current user is not a "Spark" owner team and "Salesperson" security role
    if (isUserOnSparkTeam && isUserSalesperson) {
        accountControl.setDisabled(false);
    } else {
        accountControl.setDisabled(true);
    }
}

/**
 * @function NewPatternForm.CheckCreditHold
 * @description checks "checkAccountCreditHold" config value and retrieves the related account "creditonhold". If on credit hold a notification is displayed on the form
 */
NewPatternForm.CheckCreditHold = async function (formContext) {
    "use strict";

    //get the "Account" lookup attribute value from the formContext
    var accountValue = formContext.getAttribute("dse_accountid").getValue()
    if (accountValue !== null) {

        //Returns Promise - determine if we should retrieve the Related Accounts "Credit Hold" value
        var checkAccountCreditHoldPromise = Spark30Common.GetConfigurationValue("checkAccountCreditHold");
        
        //Returns Promise - retrieve the Related Account "Credit Hold" value
        var isAccountCreditHoldPromise = Spark30Common.IsCreditHoldAccount(accountValue[0].id);

        //Resolve both promises in parallel.
        const [checkAccountCreditHold, isAccountCreditHold] = await Promise.all([checkAccountCreditHoldPromise, isAccountCreditHoldPromise]);

        //display a form notification if the Related Accounts is on credit hold
        if (checkAccountCreditHold === "1") {
            if(isAccountCreditHold === true) {
                formContext.ui.setFormNotification("Related Account is on credit hold!", "WARNING", "accountCreditHoldNotification");
            } else {
                formContext.ui.clearFormNotification("accountCreditHoldNotification");
            }
        }
    }
}

/**
 * @function NewPatternForm.CheckIsAuditor
 * @description checks "checkAccountCreditHold" config value and retrieves the related account "creditonhold". If on credit hold a notification is displayed on the form
 *              
 *              Function method signature is not designated with async so it will not return a promise to block UCI
 */
 NewPatternForm.CheckIsAuditor = function (formContext) {
    "use strict";

    //Returns promise. Retrieve User Table custom column(dse_isauditor) indicating if the user is an Auditor. Caches value in sessionStorage after first retrieval. Cache is used on subsequent calls
    var isAuditorPromise = Spark30Common.UserIsAuditor();

    //Async unhide the Audit section if the Users is an Auditor
    isAuditorPromise.then((CheckIsUserAuditor) => {
        var tabObj = formContext.ui.tabs.get("general_tab");
        var sectionObj = tabObj.sections.get("audit_section");
        if(CheckIsUserAuditor === true) {
            sectionObj.setVisible(true);
        } else {
            sectionObj.setVisible(false);
        }
    });
}