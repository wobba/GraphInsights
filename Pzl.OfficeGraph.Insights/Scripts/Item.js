/// <reference path="typings/moment/moment.d.ts" />
"use strict";
var Pzl;
(function (Pzl) {
    var OfficeGraph;
    (function (OfficeGraph) {
        var Insight;
        (function (Insight) {
            var Item = (function () {
                function Item() {
                }
                Item.prototype.getNumberOfEditsByActor = function (actor, mode) {
                    var edits = 0;
                    for (var i = 0; i < this.rawEdges.length; i++) {
                        var edge = this.rawEdges[i];
                        if ((mode === 0 /* ActorOnly */ && edge.actorId === actor.id) || (mode === 1 /* AllButActor */ && edge.actorId !== actor.id)) {
                            edits = edits + edge.weight;
                        }
                    }
                    return edits;
                };
                Item.prototype.getNumberOfContributors = function () {
                    return this.rawEdges.length;
                };
                Item.prototype.getContributorActorIds = function () {
                    var actorIds = [];
                    for (var i = 0; i < this.rawEdges.length; i++) {
                        actorIds.push(this.rawEdges[i].actorId);
                    }
                    return actorIds;
                };
                Item.prototype.actorIsCreator = function (actor) {
                    return this.createdBy.indexOf(actor.accountName) >= 0;
                };
                Item.prototype.actorIsLastModifed = function (actor) {
                    return this.lastModifiedByAccount.indexOf(actor.accountName) >= 0;
                };
                Item.prototype.getMinDateEdge = function (actorId) {
                    var date = new Date(2099, 12, 31);
                    for (var i = 0; i < this.rawEdges.length; i++) {
                        if (this.rawEdges[i].time < date && this.rawEdges[i].actorId === actorId) {
                            date = this.rawEdges[i].time;
                        }
                    }
                    return date;
                };
                Item.prototype.getMaxDateEdge = function () {
                    var date = new Date(1970, 1, 1);
                    for (var i = 0; i < this.rawEdges.length; i++) {
                        if (this.rawEdges[i].time > date) {
                            date = this.rawEdges[i].time;
                        }
                    }
                    return date;
                };
                Item.prototype.itemLifeSpanInDays = function () {
                    var ms = moment(this.lastModifiedDate).diff(moment(this.createdDate));
                    var d = moment.duration(ms);
                    return d.days();
                };
                return Item;
            })();
            Insight.Item = Item;
            (function (Inclusion) {
                Inclusion[Inclusion["ActorOnly"] = 0] = "ActorOnly";
                Inclusion[Inclusion["AllButActor"] = 1] = "AllButActor";
            })(Insight.Inclusion || (Insight.Inclusion = {}));
            var Inclusion = Insight.Inclusion;
        })(Insight = OfficeGraph.Insight || (OfficeGraph.Insight = {}));
    })(OfficeGraph = Pzl.OfficeGraph || (Pzl.OfficeGraph = {}));
})(Pzl || (Pzl = {}));
//# sourceMappingURL=Item.js.map