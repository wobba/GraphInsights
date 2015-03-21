/// <reference path="typings/moment/moment.d.ts" />
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
                Actor.prototype.getNumberOfModificationsByYou = function () {
                    var count = 0;
                    for (var i = 0; i < this.items.length; i++) {
                        count = count + this.items[i].getNumberOfEditsByActor(this, 0 /* ActorOnly */);
                    }
                    return count;
                };
                Actor.prototype.getModificationsPerDay = function () {
                    var start = this.getMinEdgeDate();
                    var end = this.getMaxEdgeDate();
                    var ms = moment(end).diff(moment(start));
                    var d = moment.duration(ms);
                    var days = d.days();
                    var mods = this.getNumberOfModificationsByYou();
                    return Math.round(mods / days);
                };
                Actor.prototype.getMinEdgeDate = function () {
                    var date = new Date(2099, 12, 31);
                    for (var i = 0; i < this.items.length; i++) {
                        var itemDate = this.items[i].getMinDateEdge();
                        if (itemDate < date) {
                            date = itemDate;
                        }
                    }
                    return date;
                };
                Actor.prototype.getMaxEdgeDate = function () {
                    var date = new Date(1970, 1, 1);
                    for (var i = 0; i < this.items.length; i++) {
                        var itemDate = this.items[i].getMaxDateEdge();
                        if (itemDate > date) {
                            date = itemDate;
                        }
                    }
                    return date;
                };
                return Actor;
            })();
            Insight.Actor = Actor;
            var Edge = (function () {
                function Edge() {
                }
                return Edge;
            })();
            Insight.Edge = Edge;
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