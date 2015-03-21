///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 
///<reference path="typings/d3/d3.d.ts" /> 

'use strict';

//ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage()
{
    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    jQuery(document).ready(() => {
        SP.SOD.executeFunc("sp.requestexecutor.js", "SP.RequestExecutor", () => {
            var helper = new Pzl.OfficeGraph.Insight.SearchHelper();
            helper.loadAllOfMe().done(me => {
                console.log(me.name + " has " + me.associates.length + " associates and " + me.items.length + " items and " + me.collabItems.length + " collab items");

                for (var i = 0; i < me.associates.length; i++) {
                    var c = me.associates[i];
                    helper.populateActor(c).done(c => {
                        if (c.items.length === 0) return;
                        console.log(c.name + " has " + c.associates.length + " associates and " + c.items.length + " items and " + c.collabItems.length + " collab items");
                    });
                }

            });
        });
    });
}

SP.SOD.executeFunc("sp.js", null, () => {
    initializePage();
});
