///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 
///<reference path="typings/d3/d3.d.ts" /> 

'use strict';

//ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage()
{
    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    jQuery(document).ready(function () {
        SP.SOD.executeFunc("sp.requestexecutor.js", "SP.RequestExecutor", () => {
            var helper = new Pzl.OfficeGraph.Insight.SearchHelper();
            //helper.loadActorsByQuery("mikael svenson").then(actors => {
            helper.loadColleagues().then(actors => {
                console.log("actors count: " + actors.length);
                for (var i = 0; i < actors.length; i++) {
                    //alert(actors[i].name);
                }
            });
        });
    });
}

SP.SOD.executeFunc("sp.js", null, () => {
    initializePage();
});
