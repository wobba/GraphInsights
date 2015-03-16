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
                //constructor() {
                //    this.id = workId;
                //}
                Actor.prototype.getNumberOfModifications = function () {
                    var count = 0;
                    var edges = this.edges;
                    for (var edge in edges) {
                        if (edges.hasOwnProperty(edge)) {
                            if (edge.action === 1003 /* Modified */) {
                                count += edge.weight;
                            }
                        }
                    }
                    return count;
                };
                Actor.prototype.getMinEdgeDate = function () {
                    var date = new Date(2099, 12, 31);
                    var edges = this.edges;
                    for (var edge in edges) {
                        if (edges.hasOwnProperty(edge)) {
                            if (edge.time < date) {
                                date = edge.time;
                            }
                        }
                    }
                    return date;
                };
                Actor.prototype.getMaxEdgeDate = function () {
                    var date = new Date(1970, 1, 1);
                    var edges = this.edges;
                    for (var edge in edges) {
                        if (edges.hasOwnProperty(edge)) {
                            if (edge.time > date) {
                                date = edge.time;
                            }
                        }
                    }
                    return date;
                };
                Actor.prototype.getModificationsPerDay = function () {
                    var start = this.getMinEdgeDate();
                    var end = this.getMaxEdgeDate();
                    var ms = moment(end).diff(moment(start));
                    var d = moment.duration(ms);
                    var days = d.days();
                    var mods = this.getNumberOfModifications();
                    return Math.round(mods / days);
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