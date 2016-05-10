/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/jquery/jquery.d.ts" />
/// <reference path="typings/q/Q.d.ts" />
"use strict";
var Pzl;
(function (Pzl) {
    var OfficeGraph;
    (function (OfficeGraph) {
        var Insight;
        (function (Insight) {
            var SearchHelper = (function () {
                function SearchHelper() {
                    this.allReachedActors = new Hashtable();
                }
                SearchHelper.prototype.postJson = function (payload, success, failure) {
                    var searchUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/search/postquery";
                    $.ajax({
                        type: "POST",
                        headers: {
                            "accept": "application/json;odata=minimalmetadata",
                            "content-type": "application/json;odata=minimalmetadata",
                            "X-RequestDigest": $("#__REQUESTDIGEST").val()
                        },
                        data: JSON.stringify(payload),
                        url: searchUrl,
                        success: success,
                        error: failure
                    });
                };
                SearchHelper.prototype.loadMe = function () {
                    var deferred = Q.defer();
                    var me = this.loadActorsByQuery(_spPageContextInfo.userLoginName);
                    me.done(function (actors) {
                        deferred.resolve(actors[0]);
                    });
                    return deferred.promise;
                };
                SearchHelper.prototype.loadActorsByQuery = function (query) {
                    var _this = this;
                    var deferred = Q.defer();
                    var searchPayload = {
                        'request': {
                            'Querytext': query,
                            'RowLimit': 500,
                            'SourceId': "b09a7990-05ea-4af9-81ef-edfab16c4e31",
                            'ClientType': 'PzlGraphInsight'
                        }
                    };
                    this.postJson(searchPayload, function (data) {
                        var actors = [];
                        var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                        for (var i = 0; i < resultsCount; i++) {
                            var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                            var actor = _this.parseActorResults(row);
                            actors.push(actor);
                        }
                        deferred.resolve(actors);
                    }, function (error) {
                        console.log(JSON.stringify(error));
                        deferred.reject(JSON.stringify(error));
                    });
                    return deferred.promise;
                };
                SearchHelper.prototype.loadModifiedItemsForActor = function (actor) {
                    var _this = this;
                    var deferred = Q.defer();
                    var searchPayload = this.getPayload("*", "ACTOR(" + actor.id + ", action:" + Insight.Action.Modified + ")");
                    this.postJson(searchPayload, function (data) {
                        var items = [];
                        if (data.PrimaryQueryResult != null) {
                            var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                            for (var i = 0; i < resultsCount; i++) {
                                var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                                var item = _this.parseItemResults(row);
                                items.push(item);
                            }
                        }
                        deferred.resolve(items);
                    }, function (error) {
                        console.log(JSON.stringify(error));
                        deferred.reject(JSON.stringify(error));
                    });
                    return deferred.promise;
                };
                SearchHelper.prototype.loadCollabModifiedItemsForActor = function (actor) {
                    var _this = this;
                    var deferred = Q.defer();
                    var template = "actor(#ID#,action:" + Insight.Action.Modified + ")";
                    var parts = [];
                    parts.push(template.replace("#ID#", actor.id.toString()));
                    var actorIds = this.allReachedActors.keys();
                    for (var i = 0; i < actorIds.length; i++) {
                        parts.push(template.replace("#ID#", actorIds[i].toString()));
                    }
                    if (parts.length === 1) {
                        parts.push(parts[0]); // fix to not fail or query
                    }
                    var fql = "and(actor(" + actor.id + ",action:" + Insight.Action.Modified + "),or(" + parts.join() + "))";
                    var searchPayload = this.getPayload("*", fql);
                    this.postJson(searchPayload, function (data) {
                        var items = [];
                        if (data.PrimaryQueryResult != null) {
                            var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                            for (var i = 0; i < resultsCount; i++) {
                                var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                                var item = _this.parseItemResults(row);
                                items.push(item);
                            }
                        }
                        deferred.resolve(items);
                    }, function (error) {
                        console.log(JSON.stringify(error));
                        deferred.reject(JSON.stringify(error));
                    });
                    return deferred.promise;
                };
                SearchHelper.prototype.loadColleagues = function (actor, reach) {
                    var _this = this;
                    var deferred = Q.defer();
                    var searchPayload = this.getPayloadActor(reach, "*", "ACTOR(" + actor.id + ", or(action:1013,action:1014,action:1015,action:1016,action:1019,action:1033,action:1035,action:1041))");
                    this.postJson(searchPayload, function (data) {
                        var actors = [];
                        if (data.PrimaryQueryResult != null) {
                            var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                            for (var i = 0; i < resultsCount; i++) {
                                var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                                var newActor = _this.parseActorResults(row);
                                actors.push(newActor);
                            }
                        }
                        deferred.resolve(actors);
                    }, function (error) {
                        console.log(JSON.stringify(error));
                        deferred.reject(JSON.stringify(error));
                    });
                    return deferred.promise;
                };
                SearchHelper.prototype.loadAllOfMe = function (reach) {
                    var _this = this;
                    var deferred = Q.defer();
                    this.loadMe()
                        .then(function (me) {
                        _this.mainActor = me;
                        return _this.loadColleagues(me, reach);
                    }).then(function (colleagues) {
                        deferred.resolve(colleagues);
                    });
                    return deferred.promise;
                };
                //populateActor(actor: Actor, reach: number): Q.Promise<Actor[]> {
                //    var deferred = Q.defer<Actor[]>();
                //    this.loadColleagues(actor, reach)
                //        .then(colleagues => {
                //        deferred.resolve(colleagues);
                //    });
                //    return deferred.promise;
                //}
                SearchHelper.prototype.getPayload = function (query, graphQuery) {
                    return {
                        "request": {
                            "Querytext": query,
                            "RowLimit": 500,
                            "TrimDuplicates": false,
                            "RankingModelId": "0c77ded8-c3ef-466d-929d-905670ea1d72",
                            'SelectProperties': ['Title', 'Write', 'Path', 'Created', 'AuthorOWSUSER', 'EditorOWSUSER', 'ModifiedBy', 'DocId', 'Edges'],
                            "ClientType": "PzlGraphInsight",
                            "Properties": [
                                {
                                    "Name": "GraphQuery",
                                    "Value": { "StrVal": graphQuery, "QueryPropertyValueTypeIndex": 1 }
                                },
                                {
                                    "Name": "GraphRankingModel",
                                    "Value": {
                                        "StrVal": "{\"features\":[{\"function\":\"EdgeTime\"}]}",
                                        "QueryPropertyValueTypeIndex": 1
                                    }
                                }]
                        }
                    };
                };
                SearchHelper.prototype.getPayloadActor = function (rowLimit, query, graphQuery) {
                    //action:1033,weight:1,edgeFunc:weight,mergeFunc:max
                    return {
                        "request": {
                            "Querytext": query,
                            "RowLimit": rowLimit,
                            "RankingModelId": "0c77ded8-c3ef-466d-929d-905670ea1d72",
                            'SelectProperties': ['AccountName', 'PreferredName', 'PictureURL'],
                            "ClientType": "PzlGraphInsight",
                            "Properties": [
                                {
                                    "Name": "GraphQuery",
                                    "Value": { "StrVal": graphQuery, "QueryPropertyValueTypeIndex": 1 }
                                },
                                {
                                    "Name": "GraphRankingModel",
                                    "Value": {
                                        //"StrVal": "{\"features\":[{\"function\":\"EdgeWeight\"}]}",
                                        "StrVal": "{\"features\":[{\"action\":\"1033\",\"function\":\"EdgeWeight\"},{\"action\":\"1019\",\"function\":\"EdgeWeight\"}]}",
                                        "QueryPropertyValueTypeIndex": 1
                                    }
                                }]
                        }
                    };
                };
                SearchHelper.prototype.parseActorResults = function (row) {
                    var actor = new Insight.Actor();
                    for (var i = 0; i < row.Cells.length; i++) {
                        var cell = row.Cells[i];
                        if (cell.Key === 'PreferredName') {
                            actor.name = cell.Value;
                        }
                        else if (cell.Key === 'PictureURL') {
                            actor.pictureUrl = cell.Value;
                        }
                        else if (cell.Key === 'DocId') {
                            actor.id = parseInt(cell.Value);
                        }
                        else if (cell.Key === 'AccountName') {
                            actor.accountName = cell.Value;
                        }
                    }
                    return actor;
                };
                SearchHelper.prototype.parseItemResults = function (row) {
                    var item = new Insight.Item();
                    for (var i = 0; i < row.Cells.length; i++) {
                        var cell = row.Cells[i];
                        if (cell.Key === 'Title') {
                            item.title = cell.Value;
                        }
                        else if (cell.Key === 'AuthorOWSUSER') {
                            item.createdBy = cell.Value;
                        }
                        else if (cell.Key === 'EditorOWSUSER') {
                            item.lastModifiedByAccount = cell.Value;
                        }
                        else if (cell.Key === 'ModifiedBy') {
                            item.lastModifiedByName = cell.Value;
                        }
                        else if (cell.Key === 'DocId') {
                            item.id = parseInt(cell.Value);
                        }
                        else if (cell.Key === 'Write') {
                            item.lastModifiedDate = moment(cell.Value).toDate();
                        }
                        else if (cell.Key === 'Created') {
                            item.createdDate = moment(cell.Value).toDate();
                        }
                        else if (cell.Key === 'Edges') {
                            //get the highest edge weight
                            var edges = JSON.parse(cell.Value);
                            item.rawEdges = this.parseEdgeResults(edges);
                        }
                    }
                    for (var j = 0; j < item.rawEdges.length; j++) {
                        item.rawEdges[j].workid = item.id;
                    }
                    return item;
                };
                SearchHelper.prototype.parseEdgeResults = function (inputEdges) {
                    var edges = [];
                    for (var i = 0; i < inputEdges.length; i++) {
                        var edge = new Insight.Edge();
                        edge.actorId = inputEdges[i].ActorId;
                        edge.objectId = inputEdges[i].ObjectId;
                        var actionString = inputEdges[i].Properties.Action;
                        edge.action = Insight.Action[actionString];
                        edge.weight = parseInt(inputEdges[i].Properties.Weight);
                        edge.time = moment(inputEdges[i].Properties.Time).toDate();
                        edges.push(edge);
                    }
                    return edges;
                };
                return SearchHelper;
            }());
            Insight.SearchHelper = SearchHelper;
        })(Insight = OfficeGraph.Insight || (OfficeGraph.Insight = {}));
    })(OfficeGraph = Pzl.OfficeGraph || (Pzl.OfficeGraph = {}));
})(Pzl || (Pzl = {}));
//# sourceMappingURL=SearchHelper.js.map