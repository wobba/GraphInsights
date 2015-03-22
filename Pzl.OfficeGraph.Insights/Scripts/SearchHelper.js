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
                    this.backupActorAssociates = [];
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
                    if (actor.associates.length === 0) {
                        // if no associates replace with backup actors
                        actor.associates = this.backupActorAssociates;
                    }
                    var template = "actor(#ID#,action:" + Insight.Action.Modified + ")";
                    var parts = [];
                    parts.push(template.replace("#ID#", actor.id.toString()));
                    for (var j = 0; j < actor.associates.length; j++) {
                        parts.push(template.replace("#ID#", actor.associates[j].id.toString()));
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
                SearchHelper.prototype.loadColleagues = function (actor) {
                    var _this = this;
                    var deferred = Q.defer();
                    var searchPayload = this.getPayloadActor("*", "ACTOR(" + actor.id + ", or(action:1013,action:1014,action:1015,action:1016,action:1019,action:1033,action:1035,action:1041))");
                    this.postJson(searchPayload, function (data) {
                        var actors = [];
                        if (data.PrimaryQueryResult != null) {
                            var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                            for (var i = 0; i < resultsCount; i++) {
                                var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                                var actor = _this.parseActorResults(row);
                                actors.push(actor);
                            }
                        }
                        deferred.resolve(actors);
                    }, function (error) {
                        console.log(JSON.stringify(error));
                        deferred.reject(JSON.stringify(error));
                    });
                    return deferred.promise;
                };
                SearchHelper.prototype.loadAllOfMe = function () {
                    var _this = this;
                    var deferred = Q.defer();
                    var actor;
                    this.loadMe().then(function (me) {
                        actor = me;
                        return _this.loadColleagues(me);
                    }).then(function (actors) {
                        actor.associates = actors;
                        if (_this.backupActorAssociates.length === 0 || actors.length > _this.backupActorAssociates.length) {
                            _this.backupActorAssociates = actors;
                        }
                        _this.loadCollabModifiedItemsForActor(actor).then(function (items) {
                            actor.collabItems = items;
                            deferred.resolve(actor);
                        });
                        //Q.all<any>([
                        //    this.loadCollabModifiedItemsForActor(actor).then(items => {
                        //        actor.collabItems = items;
                        //    })
                        //]).done(() => {
                        //    deferred.resolve(actor);
                        //});
                    });
                    return deferred.promise;
                };
                SearchHelper.prototype.populateActor = function (actor) {
                    var _this = this;
                    var deferred = Q.defer();
                    this.loadColleagues(actor).then(function (colleagues) {
                        actor.associates = colleagues;
                    }).then(function () {
                        Q.all([
                            _this.loadColleagues(actor).then(function (colleagues) {
                                actor.associates = colleagues;
                            }),
                            _this.loadCollabModifiedItemsForActor(actor).then(function (items) {
                                actor.collabItems = items;
                            })
                        ]).done(function () {
                            deferred.resolve(actor);
                        });
                    });
                    return deferred.promise;
                };
                SearchHelper.prototype.getPayload = function (query, graphQuery) {
                    return {
                        "request": {
                            "Querytext": query,
                            "RowLimit": 500,
                            "TrimDuplicates": false,
                            "RankingModelId": "0c77ded8-c3ef-466d-929d-905670ea1d72",
                            //title,write,path,created,AuthorOWSUSER,EditorOWSUSER
                            'SelectProperties': ['Title', 'Write', 'Path', 'Created', 'AuthorOWSUSER', 'EditorOWSUSER', 'DocId', 'Edges'],
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
                                }
                            ]
                        }
                    };
                };
                SearchHelper.prototype.getPayloadActor = function (query, graphQuery) {
                    return {
                        "request": {
                            "Querytext": query,
                            "RowLimit": 500,
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
                                        "StrVal": "{\"features\":[{\"function\":\"EdgeTime\"}]}",
                                        "QueryPropertyValueTypeIndex": 1
                                    }
                                }
                            ]
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
                    //$(row.Cells.results).each(function (ii, ee:any) {
                    //    if (ee.Key == 'PreferredName')
                    //        actor.name = ee.Value;
                    //        o.title = ee.Value;
                    //    else if (ee.Key == 'PictureURL')
                    //        o.pic = ee.Value;
                    //    else if (ee.Key == 'JobTitle')
                    //        o.text1 = ee.Value;
                    //    else if (ee.Key == 'Department')
                    //        o.text2 = ee.Value;
                    //    else if (ee.Key == 'Path')
                    //        o.path = ee.Value;
                    //    else if (ee.Key == 'DocId')
                    //        o.docId = ee.Value;
                    //    else if (ee.Key == 'Rank')
                    //        o.rank = parseFloat(ee.Value);
                    //    else if (ee.Key == 'Edges') {
                    //        //get the highest edge weight
                    //        var edges = JSON.parse(ee.Value);
                    //        o.objectId = edges[0].ObjectId;
                    //        o.actorId = edges[0].ActorId;
                    //        $(edges).each(function (i, e) {
                    //            var w = parseInt(e.Properties.Weight);
                    //            if (o.edgeWeight == null || w > o.edgeWeight)
                    //                o.edgeWeight = w;
                    //        });
                    //    }
                    //});
                    //return o;
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
                            item.lastModifiedBy = cell.Value;
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
                    //for (var i = 0; i < row.Cells.length; i++) {
                    //    var cell = row.Cells[i];
                    //    if (cell.Key === 'Edges') {
                    //        //get the highest edge weight
                    //        var edges = JSON.parse(cell.Value);
                    //        // TODO: combine edges - store all actors/weights/times
                    //        edge.actorId = edges[0].ActorId;
                    //        edge.objectId = edges[0].ObjectId;
                    //        var actionString = <string>edges[0].Properties.Action;
                    //        edge.action = Action[actionString];
                    //        edge.weight = parseInt(edges[0].Properties.Weight);
                    //        edge.time = moment(edges[0].Properties.Time).toDate();
                    //    }
                    //}
                    //return edge;
                };
                return SearchHelper;
            })();
            Insight.SearchHelper = SearchHelper;
        })(Insight = OfficeGraph.Insight || (OfficeGraph.Insight = {}));
    })(OfficeGraph = Pzl.OfficeGraph || (Pzl.OfficeGraph = {}));
})(Pzl || (Pzl = {}));
//# sourceMappingURL=SearchHelper.js.map