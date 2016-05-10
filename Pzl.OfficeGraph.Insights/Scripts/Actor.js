/// <reference path="typings/moment/moment.d.ts" />
///<reference path="typings/jquery/jquery.d.ts" /> 
"use strict";
var Pzl;
(function (Pzl) {
    var OfficeGraph;
    (function (OfficeGraph) {
        var Insight;
        (function (Insight) {
            var Actor = (function () {
                function Actor() {
                }
                // Average number of recorded saves per item
                Actor.prototype.getItemModificationsAverage = function () {
                    var count = 0;
                    for (var i = 0; i < this.collabItems.length; i++) {
                        count = count + this.collabItems[i].getNumberOfEditsByActor(this, Insight.Inclusion.ActorOnly);
                    }
                    if (count === 0)
                        return 0;
                    return Math.round(count / this.collabItems.length);
                };
                // Number of documents edited by actor only
                Actor.prototype.getEgoSaveCount = function () {
                    var meOnly = 0;
                    if (this.collabItems) {
                        for (var i = 0; i < this.collabItems.length; i++) {
                            var item = this.collabItems[i];
                            if (item.getNumberOfContributors() === 1) {
                                meOnly++;
                            }
                        }
                    }
                    return meOnly;
                };
                Actor.prototype.getCollaborationRatio = function () {
                    var meOnly = 0;
                    var all = 0;
                    if (this.collabItems) {
                        for (var i = 0; i < this.collabItems.length; i++) {
                            var item = this.collabItems[i];
                            if (item.getNumberOfContributors() === 1) {
                                meOnly++;
                            }
                            else {
                                all++;
                            }
                        }
                    }
                    if (meOnly === 0)
                        return 0;
                    return meOnly / all;
                };
                // Item count with at least 2 authors
                Actor.prototype.getCollaborationItemCount = function () {
                    var count = 0;
                    if (this.collabItems) {
                        for (var i = 0; i < this.collabItems.length; i++) {
                            var item = this.collabItems[i];
                            if (item.getNumberOfContributors() > 1) {
                                count++;
                            }
                        }
                    }
                    return count;
                };
                // Get all actors a user collaborates with
                Actor.prototype.getCollaborationActorCount = function () {
                    var uniqueActors = [];
                    if (this.collabItems) {
                        for (var i = 0; i < this.collabItems.length; i++) {
                            var actorIds = this.collabItems[i].getContributorActorIds();
                            for (var j = 0; j < actorIds.length; j++) {
                                if (uniqueActors.indexOf(actorIds[j]) === -1) {
                                    uniqueActors.push(actorIds[j]);
                                }
                            }
                        }
                    }
                    return uniqueActors.length;
                };
                // Find oldest created date with more than two authors
                Actor.prototype.getLongestLivingItemWithCollab = function () {
                    var oldestItem;
                    if (this.collabItems) {
                        for (var i = 0; i < this.collabItems.length; i++) {
                            var item = this.collabItems[i];
                            if (item.getNumberOfContributors() > 1) {
                                if (oldestItem === undefined || item.itemLifeSpanInDays() > oldestItem.itemLifeSpanInDays()) {
                                    oldestItem = item;
                                }
                            }
                        }
                    }
                    return oldestItem;
                };
                Actor.prototype.getStarterCount = function () {
                    var creatorCount = 0;
                    if (this.collabItems) {
                        for (var i = 0; i < this.collabItems.length; i++) {
                            var item = this.collabItems[i];
                            if (item.getNumberOfContributors() > 1) {
                                if (item.actorIsCreator(this)) {
                                    creatorCount++;
                                }
                            }
                        }
                    }
                    return creatorCount;
                };
                Actor.prototype.getLastSaverCount = function () {
                    var saverCount = 0;
                    if (this.collabItems) {
                        for (var i = 0; i < this.collabItems.length; i++) {
                            var item = this.collabItems[i];
                            if (item.getNumberOfContributors() > 1) {
                                if (item.actorIsLastModifed(this)) {
                                    saverCount++;
                                }
                            }
                        }
                    }
                    return saverCount;
                };
                // Get item you have most saves for
                Actor.prototype.getHighestItemSaveCount = function () {
                    var count = 0;
                    if (this.collabItems) {
                        for (var i = 0; i < this.collabItems.length; i++) {
                            var item = this.collabItems[i];
                            var itemCount = item.getMaxSaveCountforActor(this);
                            if (itemCount > count) {
                                count = itemCount;
                            }
                        }
                    }
                    return count;
                };
                return Actor;
            }());
            Insight.Actor = Actor;
            var Edge = (function () {
                function Edge() {
                }
                return Edge;
            }());
            Insight.Edge = Edge;
            var GraphedEdge = (function () {
                function GraphedEdge(actorId1, actorId2, workId) {
                    this.actorId1 = actorId1;
                    this.actorId2 = actorId2;
                    this.workId = workId;
                }
                return GraphedEdge;
            }());
            Insight.GraphedEdge = GraphedEdge;
            (function (Gender) {
                Gender[Gender["Male"] = 0] = "Male";
                Gender[Gender["Female"] = 1] = "Female";
            })(Insight.Gender || (Insight.Gender = {}));
            var Gender = Insight.Gender;
            (function (Action) {
                Action[Action["Modified"] = 1003] = "Modified";
                Action[Action["Colleague"] = 1015] = "Colleague";
                Action[Action["WorkingWithPublic"] = 1033] = "WorkingWithPublic";
                Action[Action["Manager"] = 1013] = "Manager";
            })(Insight.Action || (Insight.Action = {}));
            var Action = Insight.Action;
        })(Insight = OfficeGraph.Insight || (OfficeGraph.Insight = {}));
    })(OfficeGraph = Pzl.OfficeGraph || (Pzl.OfficeGraph = {}));
})(Pzl || (Pzl = {}));
//# sourceMappingURL=Actor.js.map