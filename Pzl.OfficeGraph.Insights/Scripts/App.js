///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 
///<reference path="typings/d3/d3.d.ts" /> 
///<reference path="typings/hashtable/hashtable.d.ts" /> 
'use strict';
var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
//ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");
var Pzl;
(function (Pzl) {
    var OfficeGraph;
    (function (OfficeGraph) {
        var Insight;
        (function (Insight) {
            function log(message) {
                if (message.length === 0) {
                    jQuery("#statusMessageArea").fadeOut();
                }
                else if (!jQuery("#statusMessageArea").is(":visible")) {
                    jQuery("#statusMessageArea").fadeIn();
                }
                jQuery("#statusMessage").fadeOut(function () {
                    $(this).text(message).fadeIn();
                });
                jQuery("#log").prepend(message + "<br/>");
                console.log(message);
            }
            var HighScoreType;
            (function (HighScoreType) {
                HighScoreType[HighScoreType["CollaborationItemCount"] = 0] = "CollaborationItemCount";
                HighScoreType[HighScoreType["CollaborationActorCount"] = 1] = "CollaborationActorCount";
                HighScoreType[HighScoreType["ItemModificationsAverage"] = 2] = "ItemModificationsAverage";
                HighScoreType[HighScoreType["EgoSaveCount"] = 3] = "EgoSaveCount";
                HighScoreType[HighScoreType["LongestLivingItemWithCollab"] = 4] = "LongestLivingItemWithCollab";
                HighScoreType[HighScoreType["HighestItemSaveCount"] = 5] = "HighestItemSaveCount";
                HighScoreType[HighScoreType["StarterCount"] = 6] = "StarterCount";
                HighScoreType[HighScoreType["LastSaverCount"] = 7] = "LastSaverCount";
            })(HighScoreType || (HighScoreType = {}));
            var HighScoreEntry = (function () {
                function HighScoreEntry(actor, value) {
                    this.actor = actor;
                    this.value = value;
                }
                return HighScoreEntry;
            })();
            var HighScoreComparison = (function () {
                function HighScoreComparison(template) {
                    this.template = template;
                    this.rankValues = [];
                }
                HighScoreComparison.prototype.getMetricString = function (benchMarkActor) {
                    if (this.rankValues.length === 0)
                        return "";
                    this.rankValues.sort(function (a, b) {
                        return b.value - a.value;
                    });
                    var topEntry = this.rankValues[0];
                    var compareActor = benchMarkActor;
                    var compareItems = this.rankValues.filter(function (item) { return (item.actor.id === compareActor.id); });
                    if (compareItems.length === 0)
                        return "";
                    var compareEntry = compareItems[0];
                    var index = this.rankValues.indexOf(compareEntry) + 1;
                    return this.template.replace("{topActor}", topEntry.actor.name).replace("{topValue}", topEntry.value.toString()).replace("{benchmarkRank}", index.toString()).replace("{total}", this.rankValues.length.toString()).replace("{compareValue}", compareEntry.value.toString());
                };
                return HighScoreComparison;
            })();
            var ItemCountHighScore = (function (_super) {
                __extends(ItemCountHighScore, _super);
                function ItemCountHighScore() {
                    _super.call(this, "<p>Most active collaborator is <b>{topActor}</b> co-authoring on <b>{topValue}</b> items, while you rank <b>#{benchmarkRank}</b> of </b>{total}</b>");
                }
                return ItemCountHighScore;
            })(HighScoreComparison);
            var ActorCountHighScore = (function (_super) {
                __extends(ActorCountHighScore, _super);
                function ActorCountHighScore() {
                    _super.call(this, "<p>Most social collaborator is <b>{topActor}</b> with a reach of <b>{topValue}</b> colleagues. You rank <b>#{benchmarkRank}</b> of </b>{total}</b> with a reach of <b>{compareValue}</b>");
                }
                return ActorCountHighScore;
            })(HighScoreComparison);
            var ActorLowCollaboratorHighScore = (function (_super) {
                __extends(ActorLowCollaboratorHighScore, _super);
                function ActorLowCollaboratorHighScore() {
                    _super.call(this, "<p>Most selfish collaborator is <b>{topActor}</b> with only <b>{topValue}</b> items as co-author");
                    this.minCollabItems = 200000;
                }
                ActorLowCollaboratorHighScore.prototype.getMetricString = function (benchMarkActor) {
                    if (this.rankValues.length === 0)
                        return "";
                    this.rankValues.sort(function (a, b) {
                        return a.value - b.value;
                    });
                    var zeroCollaborators = this.rankValues.filter(function (item) { return item.value === 0; }).map(function (item) { return item.actor.name; });
                    if (zeroCollaborators.length > 0) {
                        this.template = "<p>The bunch of <b>{zero}</b> refuse to collaborate in public";
                        var nameString = zeroCollaborators.join(", ").replace(/,([^,]*)$/, '</b> and <b>$1');
                        return this.template.replace("{zero}", nameString);
                    }
                    var topEntry = this.rankValues[0];
                    var compareActor = benchMarkActor;
                    var compareItems = this.rankValues.filter(function (item) { return (item.actor.id === compareActor.id); });
                    if (compareItems.length === 0)
                        return "";
                    var compareEntry = compareItems[0];
                    var index = this.rankValues.indexOf(compareEntry) + 1;
                    return this.template.replace("{topActor}", topEntry.actor.name).replace("{topValue}", topEntry.value.toString()).replace("{benchmarkRank}", index.toString()).replace("{total}", this.rankValues.length.toString()).replace("{compareValue}", compareEntry.value.toString());
                };
                return ActorLowCollaboratorHighScore;
            })(HighScoreComparison);
            var EgoHighScore = (function (_super) {
                __extends(EgoHighScore, _super);
                function EgoHighScore() {
                    _super.call(this, "<p>Most active ego content producer is <b>{topActor}</b> with <b>{topValue}</b> items produced all alone (vs. {collab} collab). You rank <b>#{benchmarkRank}</b> of </b>{total}</b> with <b>{compareValue}</b> items");
                }
                EgoHighScore.prototype.getMetricString = function (benchMarkActor) {
                    var metric = _super.prototype.getMetricString.call(this, benchMarkActor);
                    var topEntry = this.rankValues[0];
                    var count = topEntry.actor.getCollaborationItemCount();
                    return metric.replace("{collab}", count.toString());
                };
                return EgoHighScore;
            })(HighScoreComparison);
            var FrequentSaverHighScore = (function (_super) {
                __extends(FrequentSaverHighScore, _super);
                function FrequentSaverHighScore() {
                    _super.call(this, "<p><b>{topActor}</b> is the most frequent saver with an average of <b>{topValue}</b> saves per item. You rank <b>#{benchmarkRank}</b> of </b>{total} with an average of <b>{compareValue}</b> saves");
                }
                return FrequentSaverHighScore;
            })(HighScoreComparison);
            var TopSaverHighScore = (function (_super) {
                __extends(TopSaverHighScore, _super);
                function TopSaverHighScore() {
                    _super.call(this, "<p>If you're afraid to lose your work, talk to <b>{topActor}</b> who saved a single item a total of <b>{topValue}(!)</b> times. You rank <b>#{benchmarkRank}</b> of </b>{total} with a top save count of <b>{compareValue}</b>");
                }
                return TopSaverHighScore;
            })(HighScoreComparison);
            var ItemStarterHighScore = (function (_super) {
                __extends(ItemStarterHighScore, _super);
                function ItemStarterHighScore() {
                    _super.call(this, "<p>#1 item starter is <b>{topActor}</b> igniting a whopping <b>{topValue}</b> items. You rank <b>#{benchmarkRank}</b> of </b>{total} by creating <b>{compareValue}</b> new item(s)");
                }
                return ItemStarterHighScore;
            })(HighScoreComparison);
            var LastModifierHighScore = (function (_super) {
                __extends(LastModifierHighScore, _super);
                function LastModifierHighScore() {
                    _super.call(this, "<p>Last dude on the ball <b>{topValue}</b> times was <b>{topActor}</b>. You rank <b>#{benchmarkRank}</b> of </b>{total} with a measly <b>{compareValue}</b> save(s)");
                }
                return LastModifierHighScore;
            })(HighScoreComparison);
            Insight.searchHelper = new Insight.SearchHelper();
            var graphCanvas, edgeLength = 400, collabItemHighScore = new ItemCountHighScore(), collabActorHighScore = new ActorCountHighScore(), collabMinActorHightScore = new ActorLowCollaboratorHighScore(), collabEgoHighScore = new EgoHighScore(), frequentSaverHighScore = new FrequentSaverHighScore(), topSaverHighScore = new TopSaverHighScore(), itemStarterHighScore = new ItemStarterHighScore, lastModifierHighScore = new LastModifierHighScore, longestItem, benchmarkActor;
            function getActorById(actorId) {
                var vals = Insight.searchHelper.allReachedActors.values();
                for (var i = 0; i < vals.length; i++) {
                    if (vals[i].id === actorId) {
                        return vals[i];
                    }
                }
                return null;
            }
            function setStatsComparisonActor(actor) {
                benchmarkActor = actor;
            }
            function showStats(actorId) {
                var actor;
                var lastBenchMarkActor = benchmarkActor;
                if (actorId === 0) {
                    actor = benchmarkActor;
                }
                else {
                    actor = getActorById(actorId);
                }
                setStatsComparisonActor(actor);
                jQuery(".statsArea").css("background-image", "url('" + actor.pictureUrl.replace("MThumb", "LThumb") + "');");
                jQuery("#message").empty();
                jQuery("#message").append(collabItemHighScore.getMetricString(actor));
                jQuery("#message").append(collabMinActorHightScore.getMetricString(actor));
                jQuery("#message").append(collabActorHighScore.getMetricString(actor));
                jQuery("#message").append(collabEgoHighScore.getMetricString(actor));
                jQuery("#message").append(frequentSaverHighScore.getMetricString(actor));
                jQuery("#message").append(topSaverHighScore.getMetricString(actor));
                if (longestItem) {
                    jQuery("#message").append("<p><b>" + longestItem.lastModifiedByName + "</b> refuse to let go and has kept an item alive for <b>" + longestItem.itemLifeSpanInDays() + "</b> days");
                }
                jQuery("#message").append(itemStarterHighScore.getMetricString(actor));
                jQuery("#message").append(lastModifierHighScore.getMetricString(actor));
                //reset to benchmarkActor
                setStatsComparisonActor(lastBenchMarkActor);
            }
            function updateStats(actor) {
                try {
                    //if (benchMarkAgainstActor) {
                    //    benchmarkActor = actor;
                    //    return;
                    //}
                    if (!actor.collabItems && actor.id !== benchmarkActor.id)
                        return; // ensure benchmark user is included even though there is no collab docs seen
                    var currentCollabItemCount = actor.getCollaborationItemCount();
                    var collabItemEntry = new HighScoreEntry(actor, currentCollabItemCount);
                    collabItemHighScore.rankValues.push(collabItemEntry);
                    var minCollabEntry = new HighScoreEntry(actor, currentCollabItemCount);
                    collabMinActorHightScore.rankValues.push(minCollabEntry);
                    var thisMaxCollaborators = actor.getCollaborationActorCount();
                    var collabActorEntry = new HighScoreEntry(actor, thisMaxCollaborators);
                    collabActorHighScore.rankValues.push(collabActorEntry);
                    var thisMaxEgo = actor.getEgoSaveCount();
                    var egoCollabEntry = new HighScoreEntry(actor, thisMaxEgo);
                    collabEgoHighScore.rankValues.push(egoCollabEntry);
                    var thisMaxEditsPerItemAverage = actor.getItemModificationsAverage();
                    var frequentSaverEntry = new HighScoreEntry(actor, thisMaxEditsPerItemAverage);
                    frequentSaverHighScore.rankValues.push(frequentSaverEntry);
                    var thisLongestItem = actor.getLongestLivingItemWithCollab();
                    if ((longestItem === undefined && thisLongestItem !== undefined) || (thisLongestItem !== undefined && thisLongestItem.itemLifeSpanInDays() > longestItem.itemLifeSpanInDays())) {
                        longestItem = thisLongestItem;
                    }
                    var thisMaxSaverPerItem = actor.getHighestItemSaveCount();
                    var topSaverEntry = new HighScoreEntry(actor, thisMaxSaverPerItem);
                    topSaverHighScore.rankValues.push(topSaverEntry);
                    var thisMaxCreator = actor.getStarterCount();
                    var itemStarterEntry = new HighScoreEntry(actor, thisMaxCreator);
                    itemStarterHighScore.rankValues.push(itemStarterEntry);
                    var thismaxModifier = actor.getLastSaverCount();
                    var lastModifierEntry = new HighScoreEntry(actor, thismaxModifier);
                    lastModifierHighScore.rankValues.push(lastModifierEntry);
                    showStats(0);
                }
                catch (e) {
                    console.log(e.message);
                }
            }
            function updateSlider() {
                var max = graphCanvas.maxCount() - 1;
                var slider = jQuery("#filterSlider");
                var data = jQuery("#steplist");
                var options = jQuery("#steplist option");
                if (options.length < max) {
                    jQuery("#maxValue").text(max);
                    slider.attr("max", max);
                    options.remove();
                    for (var i = 0; i < max; i++) {
                        data.append(jQuery('<option></option>').html(i.toString()));
                    }
                }
            }
            function getAssociateNameById(actorId) {
                var vals = Insight.searchHelper.allReachedActors.values();
                for (var i = 0; i < vals.length; i++) {
                    if (vals[i].id === actorId) {
                        return vals[i].name;
                    }
                }
                return actorId.toString();
            }
            function addNodeAndLink(srcId, destId, timeout) {
                if (srcId === destId)
                    return;
                var srcName = getAssociateNameById(srcId);
                var destName = getAssociateNameById(destId);
                Q.delay(timeout).done(function () {
                    graphCanvas.addNode(srcName, srcId);
                    graphCanvas.addNode(destName, destId);
                    graphCanvas.addLink(srcName, destName, edgeLength);
                    Insight.Graph.keepNodesOnTop();
                    updateSlider();
                });
            }
            var seenEdges = [];
            function hasEdge(edge, actor) {
                var n1 = edge.actorId;
                var n2 = actor.id;
                if (n2 < n1) {
                    var n3 = n1;
                    n1 = n2;
                    n2 = n3;
                }
                for (var i = 0; i < seenEdges.length; i++) {
                    if (seenEdges[i].workId === edge.workid && seenEdges[i].actorId1 === n1 && seenEdges[i].actorId2 === n2) {
                        return true;
                    }
                }
                var ge = new Insight.GraphedEdge(n1, n2, edge.workid);
                seenEdges.push(ge);
                return false;
            }
            function graphEdges(actor, lastActor) {
                //TODO: add promise - add if last then log message
                log("Graphing edges for " + actor.name);
                var pause = 0;
                for (var i = 0; i < actor.collabItems.length; i++) {
                    var item = actor.collabItems[i];
                    if (item.getNumberOfContributors() > 1) {
                        pause++;
                        for (var edgeCount = 0; edgeCount < item.rawEdges.length; edgeCount++) {
                            if (hasEdge(item.rawEdges[edgeCount], actor)) {
                                continue;
                            }
                            //var name = getAssociateNameById(item.rawEdges[edgeCount].actorId);
                            //addNodeAndLink(actor.name, name, 500 * (edgeCount + pause));
                            addNodeAndLink(actor.id, item.rawEdges[edgeCount].actorId, 500 * (edgeCount + pause));
                        }
                        ;
                    }
                }
            }
            var lastfilterCount = 0;
            function hideSingleCollab(count) {
                if (graphCanvas && count !== lastfilterCount) {
                    lastfilterCount = count;
                    graphCanvas.showFilterByCount(count);
                }
            }
            Insight.hideSingleCollab = hideSingleCollab;
            function resetDataAndUI() {
                jQuery("#log").empty();
                jQuery("#message").empty();
                jQuery("#steplist option").remove();
                jQuery("#maxValue").text("1");
                lastfilterCount = 0;
                Insight.searchHelper.allReachedActors.clear();
                seenEdges = [];
                collabItemHighScore = new ItemCountHighScore();
                collabActorHighScore = new ActorCountHighScore();
                collabMinActorHightScore = new ActorLowCollaboratorHighScore();
                collabEgoHighScore = new EgoHighScore();
                frequentSaverHighScore = new FrequentSaverHighScore();
                topSaverHighScore = new TopSaverHighScore();
                itemStarterHighScore = new ItemStarterHighScore;
                lastModifierHighScore = new LastModifierHighScore;
                longestItem = undefined;
                benchmarkActor = undefined;
            }
            function loadColleaguesFor(count, total, associateActor, reach) {
                var deferred = Q.defer();
                Q.delay(500 * count).done(function () {
                    log("Hiring co-actors for " + associateActor.name);
                    Insight.searchHelper.loadColleagues(associateActor, reach).then(function (actors) {
                        for (var j = 0; j < actors.length; j++) {
                            if (!Insight.searchHelper.allReachedActors.containsKey(actors[j].id)) {
                                Insight.searchHelper.allReachedActors.put(actors[j].id, actors[j]);
                            }
                        }
                        if (count === total - 1) {
                            log("Total cast is " + Insight.searchHelper.allReachedActors.size() + " actors");
                            Q.delay(2500).then(function () {
                                deferred.resolve(true); // loaded all actors
                            });
                        }
                        else {
                            log(Insight.searchHelper.allReachedActors.size() + " actors on audition so far");
                            deferred.resolve(false);
                        }
                    });
                });
                return deferred.promise;
            }
            function queueActorForGraph(actor, count) {
                Q.delay(count * 500).then(function () {
                    Insight.searchHelper.loadCollabModifiedItemsForActor(actor).then(function (items) {
                        if (items.length > 0) {
                            actor.collabItems = items;
                            console.log("Collab actors and items:" + actor.name + " : " + actor.getCollaborationActorCount() + ":" + actor.getCollaborationItemCount());
                            graphEdges(actor, count === Insight.searchHelper.allReachedActors.size());
                            updateStats(actor);
                        }
                        if (count === Insight.searchHelper.allReachedActors.size()) {
                            log("All actors on edge have been cast!");
                            Q.delay(3500).then(function () {
                                log("");
                            });
                        }
                    });
                });
            }
            function loadEdgesForAll() {
                Q.delay(1000).then(function () {
                    log("5");
                }).delay(1000).then(function () {
                    log("4");
                }).delay(1000).then(function () {
                    log("3");
                }).delay(1000).then(function () {
                    log("2");
                }).delay(1000).then(function () {
                    log("1");
                }).delay(1000).then(function () {
                    log("GO!");
                }).then(function () {
                    var count = 0;
                    Insight.searchHelper.allReachedActors.each(function (actorId, actor) {
                        count++;
                        queueActorForGraph(actor, count);
                    });
                });
            }
            function initializePage(reach) {
                resetDataAndUI();
                jQuery(document).ready(function () {
                    graphCanvas = Insight.Graph.init("forceGraph", showStats);
                    SP.SOD.executeFunc("sp.requestexecutor.js", "SP.RequestExecutor", function () {
                        var runfunc;
                        if (!peoplePickerActor) {
                            runfunc = Insight.searchHelper.loadAllOfMe(reach);
                        }
                        else {
                            Insight.searchHelper.mainActor = peoplePickerActor;
                            runfunc = Insight.searchHelper.loadColleagues(peoplePickerActor, reach);
                        }
                        runfunc.then(function (associates) {
                            setStatsComparisonActor(Insight.searchHelper.mainActor);
                            //                    updateStats(searchHelper.mainActor, true);
                            Insight.searchHelper.allReachedActors.put(Insight.searchHelper.mainActor.id, Insight.searchHelper.mainActor);
                            for (var i = 0; i < associates.length; i++) {
                                var associate = associates[i];
                                if (!Insight.searchHelper.allReachedActors.containsKey(associate.id)) {
                                    Insight.searchHelper.allReachedActors.put(associate.id, associate);
                                }
                                if (associate.id === Insight.searchHelper.mainActor.id) {
                                    continue;
                                }
                                loadColleaguesFor(i, associates.length, associate, reach).done(function (isLastActor) {
                                    if (isLastActor) {
                                        loadEdgesForAll();
                                    }
                                });
                            }
                        });
                    });
                });
            }
            Insight.initializePage = initializePage;
            SP.SOD.executeFunc("sp.js", null, function () {
                //       initializePage();       
            });
        })(Insight = OfficeGraph.Insight || (OfficeGraph.Insight = {}));
    })(OfficeGraph = Pzl.OfficeGraph || (Pzl.OfficeGraph = {}));
})(Pzl || (Pzl = {}));
//http://www.getcodesamples.com/src/56AF1EC1/BBFD4D7A
$(document).ready(function () {
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
        loadPeoplePicker("peoplePickerDiv");
    });
});
var peoplePickerActor = null;
//Load the people picker
function loadPeoplePicker(peoplePickerElementId) {
    var EnsurePeoplePickerRefinementInit = function () {
        var schema = new Array();
        schema["PrincipalAccountType"] = "User";
        schema["AllowMultipleValues"] = false;
        schema["Width"] = 200;
        schema["OnUserResolvedClientScript"] = function () {
            var pickerObj = SPClientPeoplePicker.SPClientPeoplePickerDict.peoplePickerDiv_TopSpan;
            var users = pickerObj.GetAllUserInfo();
            var person = users[0];
            if (person != null) {
                var query = "accountname:" + person.AutoFillKey;
                var helper = new Pzl.OfficeGraph.Insight.SearchHelper();
                helper.loadActorsByQuery(query).done(function (actors) {
                    peoplePickerActor = actors[0];
                });
            }
            else {
                peoplePickerActor = undefined;
            }
        };
        SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates", function () {
            SP.SOD.executeFunc("clientforms.js", "SPClientPeoplePicker_InitStandaloneControlWrapper", function () {
                SPClientPeoplePicker_InitStandaloneControlWrapper(peoplePickerElementId, null, schema);
            });
        });
    };
    EnsurePeoplePickerRefinementInit();
}
//# sourceMappingURL=App.js.map