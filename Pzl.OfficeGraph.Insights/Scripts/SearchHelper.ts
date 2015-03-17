/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/jquery/jquery.d.ts" />
/// <reference path="typings/q/Q.d.ts" />
"use strict";

module Pzl.OfficeGraph.Insight {

    export class SearchHelper {
        private postJson(payload, success, failure) {
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
        }

        loadActorsByQuery(query: string): Q.Promise<Actor[]> {
            var deferred = Q.defer<Actor[]>();

            var searchPayload = {
                'request': {
                    'Querytext': query,
                    'RowLimit': 500,
                    'SourceId': "b09a7990-05ea-4af9-81ef-edfab16c4e31",
                    'ClientType': 'PzlGraphInsight'
                }
            };

            this.postJson(searchPayload, data => {
                var actors: Actor[] = [];
                var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                for (var i = 0; i < resultsCount; i++) {
                    var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                    var actor = this.parseActorResults(row);
                    actors.push(actor);
                }
                deferred.resolve(actors);
            },
                error => {
                    console.log(JSON.stringify(error));
                    deferred.reject(JSON.stringify(error));
                });
            return deferred.promise;
        }

        loadColleagues(): Q.Promise<Actor[]> {
            var deferred = Q.defer<Actor[]>();

            var searchPayload = this.getPayload("*", "ACTOR(ME, action:1015)");

            this.postJson(searchPayload, data => {
                var actors: Actor[] = [];
                var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                for (var i = 0; i < resultsCount; i++) {
                    var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                    var actor = this.parseActorResults(row);
                    actors.push(actor);
                }
                deferred.resolve(actors);
            },
                error => {
                    console.log(JSON.stringify(error));
                    deferred.reject(JSON.stringify(error));
                });
            return deferred.promise;
        }

        private getPayload(query: string, graphQuery: string) {
            return {
                "request": {
                    "Querytext": query,
                    "RowLimit": 500,
                    "RankingModelId": "0c77ded8-c3ef-466d-929d-905670ea1d72",
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
        }

        private parseActorResults(row): Actor {
            var actor = new Actor();
            for (var i = 0; i < row.Cells.length; i++) {
                var cell = row.Cells[i];
                if (cell.Key === 'PreferredName') {
                    actor.name = cell.Value;
                } else if (cell.Key === 'PictureURL') {
                    actor.pictureUrl = cell.Value;
                } else if (cell.Key === 'DocId') {
                    actor.id = cell.Value;
                } else if (cell.Key === 'AccountName') {
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
        }
    }
}