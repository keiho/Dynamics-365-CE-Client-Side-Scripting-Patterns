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
 * @function AntipatternForm.SparkTeamButtonEnableRule
 * @description Trigger as a Enable Rule. Enables the "Spark Team" button if the feature control is enabled and the user is a member of the "Spark Team"
 * @return bool
 */
 AntipatternForm.SparkTeamButtonEnableRule = function () {
    "use strict";
    
    //Sync XHR call to determine if the "Spark Team" button feature is enabled
    let hasConfig = Spark30Common.AntiGetConfigurationValue("featureSparkTeamButtonIsEnabled");
    
    //Sync XHR call to determine if the current user is a member of the "Spark" owner team
    let hasTeam = Spark30Common.AntiUserIsTeamMember("Spark");

    //return true if the feature is enable and user is a member of the "Spark" owner team
    if (hasConfig === "true" && hasTeam) {
        return true;
    } else {
        return false;
    }
}

/**
 * @function AntipatternForm.SalespersonButtonEnableRule
 * @description Trigger as a Enable Rule. Enables the "Salesperson" button if the feature control is enabled and the user has the "Salesperson" security role
 * @return bool
 */
 AntipatternForm.SalespersonButtonEnableRule = function () {
    "use strict";
    
    //Sync XHR call to determine if the "Salesperson" button feature is enabled
    let hasConfig = Spark30Common.AntiGetConfigurationValue("featureSalespersonButtonIsEnabled");
    
    //Sync XHR call to determine if the current user has the "Salesperson" security role
    let hasRole =   Spark30Common.AntiUserHasRoleName("Salesperson");

    //return true if the feature is enable and user has the "Salesperson" security role
    if (hasConfig === "true" && hasRole) {
        return true;
    } else {
        return false;
    }
}