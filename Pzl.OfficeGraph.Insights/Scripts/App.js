///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 
///<reference path="typings/d3/d3.d.ts" /> 
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
                    this.rankValues = [];
                    this.template = template;
                }
                HighScoreComparison.prototype.getMetricString = function (benchMarkActor) {
                    if (this.rankValues.length === 0)
                        return "";
                    this.rankValues.sort(function (a, b) {
                        return b.value - a.value;
                    });
                    var topEntry = this.rankValues[0];
                    var compareActor = benchMarkActor;
                    var compareEntry = this.rankValues.filter(function (item) { return (item.actor.id === compareActor.id); })[0];
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
                    var compareEntry = this.rankValues.filter(function (item) { return (item.actor.id === compareActor.id); })[0];
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
            var searchHelper = new Insight.SearchHelper(), graphCanvas, edgeLength = 400, collabItemHighScore = new ItemCountHighScore(), collabActorHighScore = new ActorCountHighScore(), collabMinActorHightScore = new ActorLowCollaboratorHighScore(), collabEgoHighScore = new EgoHighScore(), frequentSaverHighScore = new FrequentSaverHighScore(), topSaverHighScore = new TopSaverHighScore(), itemStarterHighScore = new ItemStarterHighScore, lastModifierHighScore = new LastModifierHighScore, longestItem, benchmarkActor;
            function updateStats(actor, benchMarkAgainstActor) {
                try {
                    if (benchMarkAgainstActor)
                        benchmarkActor = actor;
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
                    jQuery("#message").empty();
                    jQuery("#message").append(collabItemHighScore.getMetricString(benchmarkActor));
                    jQuery("#message").append(collabMinActorHightScore.getMetricString(benchmarkActor));
                    jQuery("#message").append(collabActorHighScore.getMetricString(benchmarkActor));
                    jQuery("#message").append(collabEgoHighScore.getMetricString(benchmarkActor));
                    jQuery("#message").append(frequentSaverHighScore.getMetricString(benchmarkActor));
                    jQuery("#message").append(topSaverHighScore.getMetricString(benchmarkActor));
                    if (longestItem) {
                        jQuery("#message").append("<p><b>" + longestItem.lastModifiedByName + "</b> refuse to let go and has kept an item alive for <b>" + longestItem.itemLifeSpanInDays() + "</b> days");
                    }
                    jQuery("#message").append(itemStarterHighScore.getMetricString(benchmarkActor));
                    jQuery("#message").append(lastModifierHighScore.getMetricString(benchmarkActor));
                }
                catch (e) {
                    log(e.message);
                }
            }
            function updateSlider() {
                var max = graphCanvas.maxCount();
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
            function addNodeAndLink(src, dest, timeout) {
                if (src === dest)
                    return;
                setTimeout(function () {
                    graphCanvas.addNode(src);
                    graphCanvas.addNode(dest);
                    graphCanvas.addLink(src, dest, edgeLength);
                    Insight.Graph.keepNodesOnTop();
                    updateSlider();
                }, timeout);
            }
            function hasEdge(seenEdges, edge) {
                for (var i = 0; i < seenEdges.length; i++) {
                    if (seenEdges[i].workid === edge.workid && seenEdges[i].actorId === edge.actorId) {
                        return true;
                    }
                }
                seenEdges.push(edge);
                return false;
            }
            function graphEdges(actor, seenEdges) {
                var pause = 0;
                for (var i = 0; i < actor.collabItems.length; i++) {
                    var item = actor.collabItems[i];
                    if (item.getNumberOfContributors() > 1) {
                        pause++;
                        for (var edgeCount = 0; edgeCount < item.rawEdges.length; edgeCount++) {
                            if (hasEdge(seenEdges, item.rawEdges[edgeCount])) {
                                continue;
                            }
                            var name = actor.getAssociateNameById(item.rawEdges[edgeCount].actorId);
                            addNodeAndLink(actor.name, name, 500 * (edgeCount + pause));
                        }
                        ;
                    }
                }
            }
            function hideSingleCollab(count) {
                if (graphCanvas)
                    graphCanvas.showFilterByCount(count);
            }
            Insight.hideSingleCollab = hideSingleCollab;
            function resetDataAndUI() {
                jQuery("#log").empty();
                jQuery("#message").empty();
                jQuery("#steplist option").remove();
                jQuery("#maxValue").text("1");
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
            function initializePage(reach) {
                resetDataAndUI();
                var seenEdges = [];
                jQuery(document).ready(function () {
                    graphCanvas = Insight.Graph.init("forceGraph");
                    SP.SOD.executeFunc("sp.requestexecutor.js", "SP.RequestExecutor", function () {
                        var runfunc;
                        if (!selectedActor) {
                            runfunc = searchHelper.loadAllOfMe(reach);
                        }
                        else {
                            runfunc = searchHelper.populateActor(selectedActor, reach);
                        }
                        runfunc.delay(1000).done(function (me) {
                            log("Processing edges for " + me.name);
                            log(me.name + "(" + me.id + ")" + " has " + me.associates.length + " associates and " + me.collabItems.length + " items");
                            updateStats(me, true);
                            graphEdges(me, seenEdges);
                            for (var i = 0; i < me.associates.length; i++) {
                                var c = me.associates[i];
                                log("Processing edges for " + c.name);
                                if (c.name === me.name) {
                                    continue;
                                }
                                searchHelper.populateActor(c, reach).delay(500 * i).done(function (c) {
                                    if (c.collabItems.length === 0) {
                                        log("No collaborative edges found for " + c.name);
                                        return;
                                    }
                                    graphEdges(c, seenEdges);
                                    log(c.name + "(" + c.id + ")" + " has " + c.associates.length + " associates and " + c.collabItems.length + " items");
                                    updateStats(c, false);
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
var selectedActor = null;
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
                    selectedActor = actors[0];
                });
            }
            else {
                selectedActor = undefined;
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