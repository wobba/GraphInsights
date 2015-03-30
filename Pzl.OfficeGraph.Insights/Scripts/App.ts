///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 
///<reference path="typings/d3/d3.d.ts" /> 

'use strict';

//ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

module Pzl.OfficeGraph.Insight {
    import MyGraph = Graph.MyGraph;
    var searchHelper = new SearchHelper(),
        graphCanvas: MyGraph,
        edgeLength = 300,
        mostCollabItems = 0,
        mostCollabItemsActor,
        minCollabItems = 200000,
        minCollabItemsActor,
        maxCollaborators = 0,
        maxCollaboratorsActor,
        maxEditsPerItemAverage = 0,
        maxEditsPerItemAverageActor,
        maxEgo = 0,
        maxEgoActor,
        maxCreator = 0,
        maxCreatorActor,
        maxModifier = 0,
        maxModifierActor,
        maxSaverPerItem = 0,
        maxSaverPerItemActor,
        longestItem: Item,
        zeroCollaborators: string[] = [];

    function updateStats(actor: Actor) {
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
            if ((longestItem === undefined && thisLongestItem !== undefined) || (thisLongestItem !== undefined && thisLongestItem.itemLifeSpanInDays() > longestItem.itemLifeSpanInDays())) {
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

            var thisMaxSaverPerItem = actor.getHighestItemSaveCount();
            if (thisMaxSaverPerItem > maxSaverPerItem) {
                maxSaverPerItem = thisMaxSaverPerItem;
                maxSaverPerItemActor = actor;
            }

            jQuery("#message").empty();
            if (mostCollabItemsActor) {
                jQuery("#message").append("<p>Most active collaborator is <b>" + mostCollabItemsActor.name + "</b> co-authoring on <b>" + mostCollabItems + "</b> items");
            }

            if (zeroCollaborators.length > 0) {
                jQuery("#message").append("<p>The bunch of <b>" + zeroCollaborators.join(", ").replace(/,([^,]*)$/, '</b> and <b>$1') + "</b> refuse to collaborate in public");
            }
            else {
                jQuery("#message").append("<p>Most selfish collaborator is <b>" + minCollabItemsActor.name + "</b> with only <b>" + minCollabItems + "</b> items as co-author");
            }

            if (maxCollaboratorsActor) {
                jQuery("#message").append("<p>Most social collaborator is <b>" + maxCollaboratorsActor.name + "</b> with a reach of <b>" + maxCollaborators + "</b> colleagues");
            }
            if (maxEgoActor) {
                jQuery("#message").append("<p>Most active ego content producer is <b>" + maxEgoActor.name + "</b> with <b>" + maxEgo + "</b> items produced all alone (vs. " + maxEgoActor.getCollaborationItemCount() + " collab)");
            }

            if (maxEditsPerItemAverageActor) {
                jQuery("#message").append("<p><b>" + maxEditsPerItemAverageActor.name + "</b> is the most frequent saver with an average of <b>" + maxEditsPerItemAverage + "</b> saves per item ");
            }

            if (longestItem) {
                jQuery("#message").append("<p><b>" + longestItem.lastModifiedByName + "</b> refuse to let go and has kept an item alive for <b>" + longestItem.itemLifeSpanInDays() + "</b> days");
            }

            if (maxSaverPerItemActor) {
                jQuery("#message").append("If you're afraid to lose your work, talk to <b>" + maxSaverPerItemActor.name + "</b> who saved a single item a total of <b>" + maxSaverPerItem + "(!)</b> times");
            }

            if (maxCreatorActor) {
                jQuery("#message").append("<p>#1 item starter is <b>" + maxCreatorActor.name + "</b> igniting a whopping <b>" + maxCreator + "</b> items");
            }

            if (maxModifierActor) {
                jQuery("#message").append("<p>Last dude on the ball <b>" + maxModifier + "</b> times was <b>" + maxModifierActor.name + "</b>");
            }

        } catch (e) {
            //alert(e);
            jQuery("#log").prepend(e);
            console.log(e.message);
        }
    }

    //function addNodeAndLink(graphCanvas, src, dest) {
    //    graphCanvas.addNode(src);
    //    graphCanvas.addNode(dest);
    //    graphCanvas.addLink(src, dest, 400);
    //}

    function addNodeAndLink(src, dest, timeout) {
        if (src === dest) return;
        setTimeout(
            () => {
                graphCanvas.addNode(src);
                graphCanvas.addNode(dest);
                graphCanvas.addLink(src, dest, edgeLength);
                Graph.keepNodesOnTop();
            }, timeout);
    }

    function hasEdge(seenEdges: Edge[], edge: Edge) {
        for (var i = 0; i < seenEdges.length; i++) {
            if (seenEdges[i].workid === edge.workid && seenEdges[i].actorId === edge.actorId) {
                return true;
            }
        }
        seenEdges.push(edge);
        return false;
    }

    function graphEdges(actor: Actor, seenEdges: Edge[]) {
        var pause = 0;
        for (var i = 0; i < actor.collabItems.length; i++) {
            var item = actor.collabItems[i];
            if (item.getNumberOfContributors() > 1) {
                pause++;
                for (var edgeCount = 0; edgeCount < item.rawEdges.length; edgeCount++) {
                    if (hasEdge(seenEdges, item.rawEdges[edgeCount])) {
                        //console.log("edge seen");
                        continue;
                    }
                    var name = actor.getAssociateNameById(item.rawEdges[edgeCount].actorId);
                    addNodeAndLink(actor.name, name, 500 * (edgeCount + pause));
                };
            }
        }
        //Graph.keepNodesOnTop();
    }

    export function hideSingleCollab(count:number) {
        graphCanvas.showFilterByCount(count);
    }

    export function initializePage(reach: number) {
        jQuery("#log").empty();
        var seenEdges: Edge[] = [];
        // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
        jQuery(document).ready(() => {
            graphCanvas = Graph.init("forceGraph");

            SP.SOD.executeFunc("sp.requestexecutor.js", "SP.RequestExecutor",() => {
                //var helper = new Insight.SearchHelper();
                var runfunc;
                if (!selectedActor) {
                    runfunc = searchHelper.loadAllOfMe(reach);
                } else {
                    runfunc = searchHelper.populateActor(selectedActor, reach);
                }

                //(function (): Q.IPromise<any> {
                //    var deferred = Q.defer<any>();

                //    return deferred.promise;
                //})().then(c => {

                //});

                runfunc.delay(1000).done(me => {
                    //graphCanvas.addNode(me.name);

                    $("#log").prepend("Processing edges for " + me.name);
                    console.log(me.name + "(" + me.id + ")" + " has " + me.associates.length + " associates and " + me.collabItems.length + " items");
                    //console.log("\tNumber of modifications: " + me.getNumberOfModificationsByYou() + " per day: " + me.getModificationsPerDay());

                    updateStats(me);

                    graphEdges(me, seenEdges);

                    for (var i = 0; i < me.associates.length; i++) {
                        var c = me.associates[i];
                        $("#log").prepend("Processing edges for " + c.name + "<br/>");
                        if (c.name === me.name) {
                            continue;
                        }
                        searchHelper.populateActor(c, reach).delay(500 * i).done(c => {
                            if (c.collabItems.length === 0) {
                                $("#log").prepend("No collaborative edges found for " + c.name + "<br/>");
                                return;
                            }

                            graphEdges(c, seenEdges);

                            console.log(c.name + "(" + c.id + ")" + " has " + c.associates.length + " associates and " + c.collabItems.length + " items");

                            updateStats(c);
                        });
                    }
                });
            });
        });

    }

    SP.SOD.executeFunc("sp.js", null,() => {
        //       initializePage();       
    });
}

//http://www.getcodesamples.com/src/56AF1EC1/BBFD4D7A
$(document).ready(() => {
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext',() => { loadPeoplePicker("peoplePickerDiv"); });
});


interface IPickerWrapper {
    (a: any, b: any, c: any);
}

declare var SPClientPeoplePicker_InitStandaloneControlWrapper: IPickerWrapper;


var selectedActor: Pzl.OfficeGraph.Insight.Actor = null;

//Load the people picker
function loadPeoplePicker(peoplePickerElementId) {
    var EnsurePeoplePickerRefinementInit = function () {
        var schema = new Array();
        schema["PrincipalAccountType"] = "User";
        schema["AllowMultipleValues"] = false;
        schema["Width"] = 300;
        schema["OnUserResolvedClientScript"] = function () {
            var pickerObj = SPClientPeoplePicker.SPClientPeoplePickerDict.peoplePickerDiv_TopSpan;
            var users = pickerObj.GetAllUserInfo();
            //var userInfo = '';
            var person = users[0];

            if (person != null) {
                var query = "accountname:" + person.AutoFillKey;
                var helper = new Pzl.OfficeGraph.Insight.SearchHelper();
                helper.loadActorsByQuery(query).done(actors => {
                    selectedActor = actors[0];
                });
            } else {
                selectedActor = undefined;
            }
        };

        SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates",() => {
            SP.SOD.executeFunc("clientforms.js", "SPClientPeoplePicker_InitStandaloneControlWrapper",() => {
                SPClientPeoplePicker_InitStandaloneControlWrapper(peoplePickerElementId, null, schema);
            });
        });
    }
    //EnsurePeoplePickerRefinementInit("peoplePicker");
    EnsurePeoplePickerRefinementInit();
}
