///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 
///<reference path="typings/d3/d3.d.ts" /> 

'use strict';

//ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

module Pzl.OfficeGraph.Insight {
    function log(message: string) {
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
        }

        rankType: HighScoreType;
        rankValues: HighScoreEntry[] = [];
        template: string; // {benchmarkRank} {compareValue} {topActor} {topValue} 

        getMetricString(benchMarkActor: Actor): string {
            if (this.rankValues.length === 0) return "";
            this.rankValues.sort((a, b) => { return b.value - a.value; });
            var topEntry = this.rankValues[0];
            var compareActor = benchMarkActor;
            var compareEntry = this.rankValues.filter(item => (item.actor.id === compareActor.id))[0];
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
            var compareEntry = this.rankValues.filter(item => (item.actor.id === compareActor.id))[0];
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
    var searchHelper = new SearchHelper(),
        graphCanvas: MyGraph,
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

    function updateStats(actor: Actor, benchMarkAgainstActor: boolean) {
        try {

            if (benchMarkAgainstActor) benchmarkActor = actor;

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

        } catch (e) {
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
        if (src === dest) return;
        setTimeout(
            () => {
                graphCanvas.addNode(src);
                graphCanvas.addNode(dest);
                graphCanvas.addLink(src, dest, edgeLength);
                Graph.keepNodesOnTop();
                updateSlider();
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
    }

    export function hideSingleCollab(count: number) {
        if (graphCanvas) graphCanvas.showFilterByCount(count);
    }

    export function initializePage(reach: number) {
        jQuery("#log").empty();
        jQuery("#steplist option").remove();
        jQuery("#maxValue").text("1");
        var seenEdges: Edge[] = [];
        jQuery(document).ready(() => {
            graphCanvas = Graph.init("forceGraph");

            SP.SOD.executeFunc("sp.requestexecutor.js", "SP.RequestExecutor",() => {
                var runfunc;
                if (!selectedActor) {
                    runfunc = searchHelper.loadAllOfMe(reach);
                } else {
                    runfunc = searchHelper.populateActor(selectedActor, reach);
                }

                runfunc.delay(1000).done(me => {
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
                        searchHelper.populateActor(c, reach).delay(500 * i).done(c => {
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
    EnsurePeoplePickerRefinementInit();
}
