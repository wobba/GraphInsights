///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 
///<reference path="typings/d3/d3.d.ts" /> 
///<reference path="typings/hashtable/hashtable.d.ts" /> 


'use strict';

//ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

module Pzl.OfficeGraph.Insight {
    function log(message: string) {
        if (message.length === 0) {
            jQuery("#statusMessageArea").fadeOut();
        } else if (!jQuery("#statusMessageArea").is(":visible")) {
            jQuery("#statusMessageArea").fadeIn();
        }
        jQuery("#statusMessage").fadeOut(function () {
            $(this).text(message).fadeIn();
        });

        jQuery("#log").prepend(message + "<br/>");
        console.log(message);
    }

    enum HighScoreType {
        CollaborationItemCount, // Actor with most items he/she's co-authoring on - # of items
        CollaborationActorCount, // Actor with most people he/she's co-authoring with - # of actors
        ItemModificationsAverage, // Actor who has the highest average of saves per item - # of saves
        EgoSaveCount, // Actor who produces most on his/her own - # of items
        LongestLivingItemWithCollab, // Document still being co-authored - # of days
        HighestItemSaveCount, // Actor who has the highest save count on a single item - # of saves
        StarterCount, // Items created by actor which are collaborated on - # of items
        LastSaverCount, // Items with multiple co-authors which actor had the last save - # of items
    }

    class HighScoreEntry {
        value: number;
        actor: Actor;

        constructor(actor: Actor, value: number) {
            this.actor = actor;
            this.value = value;
        }
    }

    class HighScoreComparison {
        constructor(template: string) {
            this.template = template;
            this.rankValues = [];
        }

        rankType: HighScoreType;
        rankValues: HighScoreEntry[];
        template: string; // {benchmarkRank} {compareValue} {topActor} {topValue} 

        getMetricString(benchMarkActor: Actor): string {
            if (this.rankValues.length === 0) return "";
            this.rankValues.sort((a, b) => { return b.value - a.value; });
            var topEntry = this.rankValues[0];
            var compareActor = benchMarkActor;
            var compareItems = this.rankValues.filter(item => (item.actor.id === compareActor.id));
            if (compareItems.length === 0) return "";
            var compareEntry = compareItems[0];
            var index = this.rankValues.indexOf(compareEntry) + 1;

            return this.template
                .replace("{topActor}", topEntry.actor.name)
                .replace("{topValue}", topEntry.value.toString())
                .replace("{benchmarkRank}", index.toString())
                .replace("{total}", this.rankValues.length.toString())
                .replace("{compareValue}", compareEntry.value.toString());
        }
    }

    class ItemCountHighScore extends HighScoreComparison {
        constructor() { super("<p>Most active collaborator is <b>{topActor}</b> co-authoring on <b>{topValue}</b> items, while you rank <b>#{benchmarkRank}</b> of </b>{total}</b>"); }
    }

    class ActorCountHighScore extends HighScoreComparison {
        constructor() { super("<p>Most social collaborator is <b>{topActor}</b> with a reach of <b>{topValue}</b> colleagues. You rank <b>#{benchmarkRank}</b> of </b>{total}</b> with a reach of <b>{compareValue}</b>"); }
    }

    class ActorLowCollaboratorHighScore extends HighScoreComparison {
        constructor() {
            super("<p>Most selfish collaborator is <b>{topActor}</b> with only <b>{topValue}</b> items as co-author");
        }

        minCollabItems: number = 200000;

        getMetricString(benchMarkActor: Actor): string {
            if (this.rankValues.length === 0) return "";
            this.rankValues.sort((a, b) => { return a.value - b.value; });

            var zeroCollaborators = this.rankValues.filter(item => item.value === 0).map(item => item.actor.name);
            if (zeroCollaborators.length > 0) {
                this.template = "<p>The bunch of <b>{zero}</b> refuse to collaborate in public";
                var nameString = zeroCollaborators.join(", ").replace(/,([^,]*)$/, '</b> and <b>$1');
                return this.template.replace("{zero}", nameString);
            }

            var topEntry = this.rankValues[0];
            var compareActor = benchMarkActor;
            var compareItems = this.rankValues.filter(item => (item.actor.id === compareActor.id));
            if (compareItems.length === 0) return "";
            var compareEntry = compareItems[0];
            var index = this.rankValues.indexOf(compareEntry) + 1;

            return this.template
                .replace("{topActor}", topEntry.actor.name)
                .replace("{topValue}", topEntry.value.toString())
                .replace("{benchmarkRank}", index.toString())
                .replace("{total}", this.rankValues.length.toString())
                .replace("{compareValue}", compareEntry.value.toString());
        }
    }

    class EgoHighScore extends HighScoreComparison {
        constructor() {
            super("<p>Most active ego content producer is <b>{topActor}</b> with <b>{topValue}</b> items produced all alone (vs. {collab} collab). You rank <b>#{benchmarkRank}</b> of </b>{total}</b> with <b>{compareValue}</b> items");
        }

        getMetricString(benchMarkActor: Actor): string {
            var metric = super.getMetricString(benchMarkActor);

            var topEntry = this.rankValues[0];

            var count = topEntry.actor.getCollaborationItemCount();
            return metric.replace("{collab}", count.toString());
        }
    }

    class FrequentSaverHighScore extends HighScoreComparison {
        constructor() {
            super("<p><b>{topActor}</b> is the most frequent saver with an average of <b>{topValue}</b> saves per item. You rank <b>#{benchmarkRank}</b> of </b>{total} with an average of <b>{compareValue}</b> saves");
        }
    }

    class TopSaverHighScore extends HighScoreComparison {
        constructor() {
            super("<p>If you're afraid to lose your work, talk to <b>{topActor}</b> who saved a single item a total of <b>{topValue}(!)</b> times. You rank <b>#{benchmarkRank}</b> of </b>{total} with a top save count of <b>{compareValue}</b>");
        }
    }

    class ItemStarterHighScore extends HighScoreComparison {
        constructor() {
            super("<p>#1 item starter is <b>{topActor}</b> igniting a whopping <b>{topValue}</b> items. You rank <b>#{benchmarkRank}</b> of </b>{total} by creating <b>{compareValue}</b> new item(s)");
        }
    }

    class LastModifierHighScore extends HighScoreComparison {
        constructor() {
            super("<p>Last dude on the ball <b>{topValue}</b> times was <b>{topActor}</b>. You rank <b>#{benchmarkRank}</b> of </b>{total} with a measly <b>{compareValue}</b> save(s)");
        }
    }
    import MyGraph = Graph.MyGraph;
    export var searchHelper = new SearchHelper();
    var graphCanvas: MyGraph,
        edgeLength = 400,
        collabItemHighScore = new ItemCountHighScore(),
        collabActorHighScore = new ActorCountHighScore(),
        collabMinActorHightScore = new ActorLowCollaboratorHighScore(),
        collabEgoHighScore = new EgoHighScore(),
        frequentSaverHighScore = new FrequentSaverHighScore(),
        topSaverHighScore = new TopSaverHighScore(),
        itemStarterHighScore = new ItemStarterHighScore,
        lastModifierHighScore = new LastModifierHighScore,
        longestItem: Item,
        benchmarkActor: Actor;

    function getActorById(actorId: number): Actor {
        var vals = searchHelper.allReachedActors.values();
        for (var i = 0; i < vals.length; i++) {
            if (vals[i].id === actorId) {
                return vals[i];
            }
        }
        return null;
    }

    function setStatsComparisonActor(actor: Actor) {
        benchmarkActor = actor;
    }

    function showStats(actorId: number) {
        var actor: Actor;
        var lastBenchMarkActor = benchmarkActor;

        if (actorId === 0) {
            actor = benchmarkActor;
        } else {
            actor = getActorById(actorId);
        }
        setStatsComparisonActor(actor);
        jQuery("#avatar").attr("src", actor.pictureUrl.replace("MThumb", "LThumb"));

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

        if (benchmarkActor.accountName.indexOf(_spPageContextInfo.userLoginName) === -1) {
            jQuery("#message").html(jQuery("#message").html().replace("you", "<b>" + actor.name + "</b>"));
        }

        //reset to benchmarkActor
        setStatsComparisonActor(lastBenchMarkActor);
    }

    function updateStats(actor: Actor) {
        try {
            //if (benchMarkAgainstActor) {
            //    benchmarkActor = actor;
            //    return;
            //}
            if (!actor.collabItems && actor.id !== benchmarkActor.id) return; // ensure benchmark user is included even though there is no collab docs seen

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

            //jQuery("#message").empty();
            //jQuery("#message").append(collabItemHighScore.getMetricString(benchmarkActor));
            //jQuery("#message").append(collabMinActorHightScore.getMetricString(benchmarkActor));
            //jQuery("#message").append(collabActorHighScore.getMetricString(benchmarkActor));
            //jQuery("#message").append(collabEgoHighScore.getMetricString(benchmarkActor));
            //jQuery("#message").append(frequentSaverHighScore.getMetricString(benchmarkActor));
            //jQuery("#message").append(topSaverHighScore.getMetricString(benchmarkActor));

            //if (longestItem) {
            //    jQuery("#message").append("<p><b>" + longestItem.lastModifiedByName + "</b> refuse to let go and has kept an item alive for <b>" + longestItem.itemLifeSpanInDays() + "</b> days");
            //}

            //jQuery("#message").append(itemStarterHighScore.getMetricString(benchmarkActor));
            //jQuery("#message").append(lastModifierHighScore.getMetricString(benchmarkActor));

        } catch (e) {
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

    function getAssociateNameById(actorId: number): string {
        var vals = searchHelper.allReachedActors.values();
        for (var i = 0; i < vals.length; i++) {
            if (vals[i].id === actorId) {
                return vals[i].name;
            }
        }
        return actorId.toString();
    }

    function addNodeAndLink(srcId: number, destId: number, timeout: number) {
        if (srcId === destId) return;
        var srcName = getAssociateNameById(srcId);
        var destName = getAssociateNameById(destId);

        Q.delay(timeout).done(() => {
            graphCanvas.addNode(srcName, srcId);
            graphCanvas.addNode(destName, destId);
            graphCanvas.addLink(srcName, destName, edgeLength);
            Graph.keepNodesOnTop();
            updateSlider();
        });
    }

    var seenEdges: GraphedEdge[] = [];
    function hasEdge(edge: Edge, actor: Actor) {
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
        var ge = new GraphedEdge(n1, n2, edge.workid);
        seenEdges.push(ge);
        return false;
    }

    function graphEdges(actor: Actor, lastActor: boolean) {
        //TODO: add promise - add if last then log message
        log("Graphing edges for " + actor.name);
        var pause = 0;
        for (var i = 0; i < actor.collabItems.length; i++) {
            var item = actor.collabItems[i];
            if (item.getNumberOfContributors() > 1) {
                pause++;
                for (var edgeCount = 0; edgeCount < item.rawEdges.length; edgeCount++) {
                    if (hasEdge(item.rawEdges[edgeCount], actor)) {
                        //console.log("edge seen");
                        continue;
                    }
                    //var name = getAssociateNameById(item.rawEdges[edgeCount].actorId);
                    //addNodeAndLink(actor.name, name, 500 * (edgeCount + pause));
                    addNodeAndLink(actor.id, item.rawEdges[edgeCount].actorId, 500 * (edgeCount + pause));
                };
            }
        }
    }

    var lastfilterCount: number = 0;
    export function hideSingleCollab(count: number) {
        if (graphCanvas && count !== lastfilterCount) {
            lastfilterCount = count;
            graphCanvas.showFilterByCount(count);
        }
    }

    function resetDataAndUI() {
        jQuery("#log").empty();
        jQuery("#message").empty();
        jQuery("#avatar").attr("src", "");
        jQuery("#steplist option").remove();
        jQuery("#maxValue").text("1");

        lastfilterCount = 0;
        searchHelper.allReachedActors.clear();
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

    function loadColleaguesFor(count: number, total: number, associateActor: Actor, reach: number) {
        var deferred = Q.defer<boolean>();
        Q.delay(500 * count).done(() => {
            log("Hiring co-actors for " + associateActor.name);
            searchHelper.loadColleagues(associateActor, reach).then(actors => {
                for (var j = 0; j < actors.length; j++) {
                    if (!searchHelper.allReachedActors.containsKey(actors[j].id)) {
                        searchHelper.allReachedActors.put(actors[j].id, actors[j]);
                    }
                }
                if (count === total - 1) {
                    log("Total cast is " + searchHelper.allReachedActors.size() + " actors");
                    Q.delay(2500).then(() => {
                        deferred.resolve(true); // loaded all actors
                    });
                } else {
                    log(searchHelper.allReachedActors.size() + " actors on audition so far");
                    deferred.resolve(false);
                }
            });
        });
        return deferred.promise;
    }

    function queueActorForGraph(actor: Actor, count: number) {
        Q.delay(count * 500).then(() => {
            searchHelper.loadCollabModifiedItemsForActor(actor).then(items => {
                if (items.length > 0) {
                    actor.collabItems = items;
                    console.log("Collab actors and items:" + actor.name + " : " + actor.getCollaborationActorCount() + ":" + actor.getCollaborationItemCount());
                    graphEdges(actor, count === searchHelper.allReachedActors.size());
                    updateStats(actor);
                }
                if (count === searchHelper.allReachedActors.size()) {
                    log("All actors on edge have been cast!");
                    Q.delay(3500).then(() => { log(""); });
                }
            });
        });
    }

    function loadEdgesForAll() {
        Q.delay(1000).then(() => { log("5"); }).delay(1000).then(() => { log("4"); }).delay(1000).then(() => { log("3"); }).delay(1000).then(() => { log("2"); }).delay(1000).then(() => { log("1"); }).delay(1000).then(() => { log("GO!"); }).then(() => {
            var count = 0;
            searchHelper.allReachedActors.each((actorId, actor) => {
                count++;
                queueActorForGraph(actor, count);
            });

        });

    }

    export function initializePage(reach: number) {
        resetDataAndUI();
        jQuery(document).ready(() => {
            graphCanvas = Graph.init("forceGraph", showStats);

            SP.SOD.executeFunc("sp.requestexecutor.js", "SP.RequestExecutor",() => {
                var runfunc: Q.IPromise<Actor[]>;
                if (!peoplePickerActor) {
                    runfunc = searchHelper.loadAllOfMe(reach);
                } else {
                    searchHelper.mainActor = peoplePickerActor;
                    runfunc = searchHelper.loadColleagues(peoplePickerActor, reach);
                }

                runfunc.then(associates => {
                    setStatsComparisonActor(searchHelper.mainActor);
                    //                    updateStats(searchHelper.mainActor, true);

                    searchHelper.allReachedActors.put(searchHelper.mainActor.id, searchHelper.mainActor);
                    for (var i = 0; i < associates.length; i++) {
                        var associate = associates[i];
                        if (!searchHelper.allReachedActors.containsKey(associate.id)) {
                            searchHelper.allReachedActors.put(associate.id, associate);
                        }

                        if (associate.id === searchHelper.mainActor.id) {
                            continue; // skip main actor - as we have loaded colleagues already
                        }

                        loadColleaguesFor(i, associates.length, associate, reach).done(isLastActor => {
                            if (isLastActor) {
                                loadEdgesForAll();
                            }
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

var peoplePickerActor: Pzl.OfficeGraph.Insight.Actor = null;

//Load the people picker
function loadPeoplePicker(peoplePickerElementId) {
    var EnsurePeoplePickerRefinementInit = () => {
        var schema = new Array();
        schema["PrincipalAccountType"] = "User";
        schema["AllowMultipleValues"] = false;
        schema["Width"] = 200;
        schema["OnUserResolvedClientScript"] = () => {
            var pickerObj = SPClientPeoplePicker.SPClientPeoplePickerDict.peoplePickerDiv_TopSpan;
            var users = pickerObj.GetAllUserInfo();
            var person = users[0];

            if (person != null) {
                var query = "accountname:" + person.AutoFillKey;
                var helper = new Pzl.OfficeGraph.Insight.SearchHelper();
                helper.loadActorsByQuery(query).done(actors => {
                    peoplePickerActor = actors[0];
                });
            } else {
                peoplePickerActor = undefined;
            }
        };

        SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates",() => {
            SP.SOD.executeFunc("clientforms.js", "SPClientPeoplePicker_InitStandaloneControlWrapper",() => {
                SPClientPeoplePicker_InitStandaloneControlWrapper(peoplePickerElementId, null, schema);
            });
        });
    };
    EnsurePeoplePickerRefinementInit();
}
