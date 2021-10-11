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

var AntipatternForm = window.AntipatternForm || {};

/**
 * @function AntipatternForm.OnLoad
 * @description Configured in form properties to trigger on form load. 
 */
 AntipatternForm.OnLoad = function (executionContext) {
    "use strict";
    var formContext = executionContext.getFormContext(); 
    
    //get the account lookup control from the formcontext 
    var accountControl = formContext.getControl("dse_accountid");
    
    //Sync XHR call to determine if the current user is a member of the "Spark" owner team
    var isUserOnSparkTeam = Spark30Common.AntiUserIsTeamMember("Spark");
    
    //Sync XHR call to determine if the current user has the "Salesperson" security role
    var isUserSalesperson =  Spark30Common.AntiUserHasRoleName("Salesperson");

    //Disables the "Account" lookup field if the current user is not a "Spark" owner team and "Salesperson" security role
    if (isUserOnSparkTeam && isUserSalesperson) {
        accountControl.setDisabled(false);
    } else {
        accountControl.setDisabled(true);
    }

    //get the "Account" lookup attribute value from the formContext
    var accountValue = formContext.getAttribute("dse_accountid").getValue()
    if (accountValue !== null) {

        //Sync XHR call to determine if we should retrieve the Related Accounts "Credit Hold" value
        var checkAccountCreditHold = Spark30Common.AntiGetConfigurationValue("checkAccountCreditHold");
        if ( checkAccountCreditHold === "1") {

            //Sync XHR call to retrieve the Related Account "Credit Hold" value
            var isAccountCreditHold = Spark30Common.AntiIsCreditHoldAccount(accountValue[0].id);

            //display a form notification if the Related Accounts is on credit hold
            if(isAccountCreditHold === true) {
                formContext.ui.setFormNotification("Related Account is on credit hold!", "WARNING", "accountCreditHoldNotification");
            } else {
                formContext.ui.clearFormNotification("accountCreditHoldNotification");
            }
        }
    }

    //Sync XHR call to retrieve User Table custom column(dse_isauditor) indicating if the user is an Auditor 
    var userIsAuditor = Spark30Common.AntiUserIsAuditor();

    //unhide the Audit section if the Users is an auditor
    var tabObj = formContext.ui.tabs.get("general_tab");
    var sectionObj = tabObj.sections.get("audit_section");

    if(userIsAuditor === true) {
        sectionObj.setVisible(true);
    } else {
        sectionObj.setVisible(false);
    }
}