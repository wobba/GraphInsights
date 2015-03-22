///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 
///<reference path="typings/d3/d3.d.ts" /> 
'use strict';
//ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");
var Pzl;
(function (Pzl) {
    var OfficeGraph;
    (function (OfficeGraph) {
        var Insight;
        (function (Insight) {
            var mostCollabItems = 0, mostCollabItemsActor, minCollabItems = 200000, minCollabItemsActor, maxCollaborators = 0, maxCollaboratorsActor, maxEditsPerItemAverage = 0, maxEditsPerItemAverageActor, maxEgo = 0, maxEgoActor, maxCreator = 0, maxCreatorActor, maxModifier = 0, maxModifierActor, longestItem, zeroCollaborators = [];
            //maxEditsPerDay = 0,
            //maxEditsPerDayActor;
            function updateStats(actor) {
                try {
                    var currentCollabItemCount = actor.getCollaborationItemCount();
                    if (currentCollabItemCount > mostCollabItems) {
                        mostCollabItems = currentCollabItemCount;
                        mostCollabItemsActor = actor;
                    }
                    if (currentCollabItemCount <= minCollabItems) {
                        minCollabItems = currentCollabItemCount;
                        minCollabItemsActor = actor;
                        if (minCollabItems === 0) {
                            zeroCollaborators.push(actor.name);
                        }
                    }
                    var thisMaxCollaborators = actor.getCollaborationActorCount();
                    if (thisMaxCollaborators > maxCollaborators) {
                        maxCollaborators = thisMaxCollaborators;
                        maxCollaboratorsActor = actor;
                    }
                    var thisMaxEditsPerItemAverage = actor.getItemModificationsAverage();
                    if (thisMaxEditsPerItemAverage > maxEditsPerItemAverage) {
                        maxEditsPerItemAverage = thisMaxEditsPerItemAverage;
                        maxEditsPerItemAverageActor = actor;
                    }
                    var thisMaxEgo = actor.getEgoSaveCount();
                    if (thisMaxEgo > maxEgo) {
                        maxEgo = thisMaxEgo;
                        maxEgoActor = actor;
                    }
                    var thisLongestItem = actor.getLongestLivingItemWithCollab();
                    if (longestItem === undefined || (thisLongestItem !== undefined && thisLongestItem.itemLifeSpanInDays() > longestItem.itemLifeSpanInDays())) {
                        longestItem = thisLongestItem;
                    }
                    var thisMaxCreator = actor.getStarterCount();
                    if (thisMaxCreator > maxCreator) {
                        maxCreator = thisMaxCreator;
                        maxCreatorActor = actor;
                    }
                    var thismaxModifier = actor.getLastSaverCount();
                    if (thismaxModifier > maxModifier) {
                        maxModifier = thismaxModifier;
                        maxModifierActor = actor;
                    }
                    jQuery("#message").empty();
                    jQuery("#message").append("<p>Most active collaborator is <b>" + mostCollabItemsActor.name + "</b> co-authoring on <b>" + mostCollabItems + "</b> items");
                    if (zeroCollaborators.length > 0) {
                        jQuery("#message").append("<p>The bunch of <b>" + zeroCollaborators.join(", ").replace(/,([^,]*)$/, '</b> and <b>$1') + "</b> refuse to collaborate in public");
                    }
                    else {
                        jQuery("#message").append("<p>Most selfish collaborator is <b>" + minCollabItemsActor.name + "</b> with only <b>" + minCollabItems + "</b> items as co-author");
                    }
                    jQuery("#message").append("<p>Most social collaborator is <b>" + maxCollaboratorsActor.name + "</b> with a reach of <b>" + maxCollaborators + "</b> colleagues");
                    jQuery("#message").append("<p>Most active ego content producer is <b>" + maxEgoActor.name + "</b> with <b>" + maxEgo + "</b> items produced all alone (vs. " + maxEgoActor.getCollaborationItemCount() + " collab)");
                    jQuery("#message").append("<p><b>" + maxEditsPerItemAverageActor.name + "</b> is the most frequent saver with an average of <b>" + maxEditsPerItemAverage + "</b> saves per item ");
                    jQuery("#message").append("<p><b>" + longestItem.lastModifiedByName + "</b> can't let go and has kept an item alive for <b>" + longestItem.itemLifeSpanInDays() + "</b> days");
                    jQuery("#message").append("<p>#1 item starter is <b>" + maxCreatorActor.name + "</b> igniting a whopping <b>" + maxCreator + "</b> items");
                    jQuery("#message").append("<p>Last dude on the ball <b>" + maxModifier + "</b> times was <b>" + maxModifierActor.name + "</b>");
                }
                catch (e) {
                    //alert(e);
                    jQuery("#log").prepend(e);
                    console.log(e.message);
                }
            }
            function initializePage() {
                // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
                jQuery(document).ready(function () {
                    SP.SOD.executeFunc("sp.requestexecutor.js", "SP.RequestExecutor", function () {
                        var helper = new Insight.SearchHelper();
                        helper.loadAllOfMe().delay(500).done(function (me) {
                            $("#log").prepend("Processing edges for " + me.name);
                            console.log(me.name + "(" + me.id + ")" + " has " + me.associates.length + " associates and " + me.collabItems.length + " items");
                            //console.log("\tNumber of modifications: " + me.getNumberOfModificationsByYou() + " per day: " + me.getModificationsPerDay());
                            updateStats(me);
                            for (var i = 0; i < me.associates.length; i++) {
                                var c = me.associates[i];
                                helper.populateActor(c).delay(500).done(function (c) {
                                    if (c.collabItems.length === 0)
                                        return;
                                    $("#log").prepend("Processing edges for " + c.name + "<br/>");
                                    console.log(c.name + "(" + c.id + ")" + " has " + c.associates.length + " associates and " + c.collabItems.length + " items");
                                    updateStats(c);
                                });
                            }
                        });
                    });
                });
            }
            SP.SOD.executeFunc("sp.js", null, function () {
                initializePage();
            });
        })(Insight = OfficeGraph.Insight || (OfficeGraph.Insight = {}));
    })(OfficeGraph = Pzl.OfficeGraph || (Pzl.OfficeGraph = {}));
})(Pzl || (Pzl = {}));
//# sourceMappingURL=App.js.map