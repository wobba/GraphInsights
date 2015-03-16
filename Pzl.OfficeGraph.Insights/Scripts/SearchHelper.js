/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/jquery/jquery.d.ts" />
"use strict";
var Pzl;
(function (Pzl) {
    var OfficeGraph;
    (function (OfficeGraph) {
        var Insight;
        (function (Insight) {
            var SearchHelper = (function () {
                function SearchHelper() {
                }
                SearchHelper.prototype.postJson = function (payload, success, failure) {
                    $.ajax({
                        type: "POST",
                        headers: {
                            "accept": "application/json;odata=minimal",
                            "content-type": "application/json;odata=minimal",
                            "X-RequestDigest": $("#__REQUESTDIGEST").val()
                        },
                        data: JSON.stringify(payload),
                        url: _spPageContextInfo.webAbsoluteUrl + "/_api/search/postquery",
                        success: success,
                        failure: failure
                    });
                };
                SearchHelper.prototype.executeActorQuery = function (query) {
                    var _this = this;
                    var result = jQuery.Deferred();
                    //var dfd = $.Deferred<void>();
                    var searchPayload = {
                        'request': {
                            '__metadata': { 'type': "Microsoft.Office.Server.Search.REST.SearchRequest" },
                            'Querytext': query,
                            'RowLimit': 500,
                            'SourceId': "b09a7990-05ea-4af9-81ef-edfab16c4e31"
                        }
                    };
                    this.postJson(searchPayload, function (data) {
                        var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                        for (var i = 0; i < resultsCount; i++) {
                            var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                            var actor = _this.parseActorResults(row);
                        }
                    }, function (error) {
                        console.log(JSON.stringify(error));
                    });
                };
                SearchHelper.prototype.executeGraphQuery = function (query, graphQuery) {
                    var searchPayload = {
                        'request': {
                            '__metadata': { 'type': "Microsoft.Office.Server.Search.REST.SearchRequest" },
                            'Querytext': query,
                            'RowLimit': 500,
                            'RankingModelId': "0c77ded8-c3ef-466d-929d-905670ea1d72",
                            'Properties': {
                                'results': [
                                    {
                                        'Name': "GraphQuery",
                                        'Value': { 'StrVal': graphQuery, 'QueryPropertyValueTypeIndex': 1 }
                                    },
                                    {
                                        'Name': "GraphRankingModel",
                                        'Value': {
                                            'StrVal': "{\"features\":[{\"function\":\"EdgeTime\"}]}",
                                            'QueryPropertyValueTypeIndex': 1
                                        }
                                    }
                                ]
                            }
                        }
                    };
                    this.postJson(searchPayload, function (data) {
                        var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                        for (var i = 0; i < resultsCount; i++) {
                            var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                        }
                    }, function (error) {
                        console.log(JSON.stringify(error));
                    });
                };
                SearchHelper.prototype.parseActorResults = function (row) {
                    var actor = new Insight.Actor();
                    for (var cell in row.Cells) {
                        if (cell.Key === 'PreferredName') {
                            actor.name = cell.Value;
                        }
                        else if (cell.Key === 'PictureURL') {
                            actor.pictureUrl = cell.Value;
                        }
                        else if (cell.Key === 'DocId') {
                            actor.id = cell.Value;
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
                return SearchHelper;
            })();
            Insight.SearchHelper = SearchHelper;
        })(Insight = OfficeGraph.Insight || (OfficeGraph.Insight = {}));
    })(OfficeGraph = Pzl.OfficeGraph || (Pzl.OfficeGraph = {}));
})(Pzl || (Pzl = {}));
//# sourceMappingURL=SearchHelper.js.map