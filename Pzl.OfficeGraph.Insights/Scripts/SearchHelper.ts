/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/jquery/jquery.d.ts" />
/// <reference path="typings/q/Q.d.ts" />
"use strict";

module Pzl.OfficeGraph.Insight {

    export class SearchHelper {
        backupActorAssociates : Actor[] = [];

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

        loadMe(): Q.IPromise<Actor> {
            var deferred = Q.defer<Actor>();
            var me = this.loadActorsByQuery(_spPageContextInfo.userLoginName);
            me.done(actors => {
                deferred.resolve(actors[0]);
            });
            return deferred.promise;
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

        loadModifiedItemsForActor(actor: Actor): Q.IPromise<Item[]> {
            var deferred = Q.defer<Item[]>();

            var searchPayload = this.getPayload("*", "ACTOR(" + actor.id + ", action:" + Action.Modified + ")");

            this.postJson(searchPayload, data => {
                var items: Item[] = [];
                if (data.PrimaryQueryResult != null) {
                    var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                    for (var i = 0; i < resultsCount; i++) {
                        var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                        var item = this.parseItemResults(row);
                        items.push(item);
                    }
                }
                deferred.resolve(items);
            },
                error => {
                    console.log(JSON.stringify(error));
                    deferred.reject(JSON.stringify(error));
                });
            return deferred.promise;
        }

        loadCollabModifiedItemsForActor(actor: Actor): Q.IPromise<Item[]> {
            var deferred = Q.defer<Item[]>();
            if (actor.associates.length === 0) {
                // if no associates replace with backup actors
                actor.associates = this.backupActorAssociates; 
            }
            var template = "actor(#ID#,action:" + Action.Modified + ")";
            var parts = [];
            parts.push(template.replace("#ID#", actor.id.toString()));
            for (var j = 0; j < actor.associates.length; j++) {
                parts.push(template.replace("#ID#", actor.associates[j].id.toString()));
            }

            var fql = "and(actor(" + actor.id + ",action:" + Action.Modified + "),or(" + parts.join() + "))";

            var searchPayload = this.getPayload("*", fql);

            this.postJson(searchPayload, data => {
                var items: Item[] = [];
                if (data.PrimaryQueryResult != null) {
                    var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                    for (var i = 0; i < resultsCount; i++) {
                        var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                        var item = this.parseItemResults(row);
                        items.push(item);
                    }
                }
                deferred.resolve(items);
            },
                error => {
                    console.log(JSON.stringify(error));
                    deferred.reject(JSON.stringify(error));
                });

            return deferred.promise;
        }

        loadColleagues(actor: Actor): Q.IPromise<Actor[]> {
            var deferred = Q.defer<Actor[]>();

            var searchPayload = this.getPayloadActor("*", "ACTOR(" + actor.id + ", or(action:1013,action:1014,action:1015,action:1016,action:1019,action:1033,action:1035,action:1041))");

            this.postJson(searchPayload, data => {
                var actors: Actor[] = [];
                if (data.PrimaryQueryResult != null) {
                    var resultsCount = data.PrimaryQueryResult.RelevantResults.RowCount;
                    for (var i = 0; i < resultsCount; i++) {
                        var row = data.PrimaryQueryResult.RelevantResults.Table.Rows[i];
                        var actor = this.parseActorResults(row);
                        actors.push(actor);
                    }
                }
                deferred.resolve(actors);
            },
                error => {
                    console.log(JSON.stringify(error));
                    deferred.reject(JSON.stringify(error));
                });
            return deferred.promise;
        }

        loadAllOfMe(): Q.Promise<Actor> {
            var deferred = Q.defer<Actor>();
            var actor: Actor;

            this.loadMe()
                .then(me => {
                actor = me;
                return this.loadColleagues(me);
            }).then(actors => {
                actor.associates = actors;
                if (this.backupActorAssociates.length === 0 || actors.length > this.backupActorAssociates.length) {
                    this.backupActorAssociates = actors;
                }
                this.loadCollabModifiedItemsForActor(actor).then(items => {
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
        }

        populateActor(actor: Actor): Q.Promise<Actor> {
            var deferred = Q.defer<Actor>();

            this.loadColleagues(actor).
                then(colleagues => {
                actor.associates = colleagues;
            }).then(() => {
                Q.all<any>([
                    this.loadColleagues(actor).then(colleagues => {
                        actor.associates = colleagues;
                    }),
                    //this.loadModifiedItemsForActor(actor).then(items => {
                    //    actor.items = items;
                    //}),
                    this.loadCollabModifiedItemsForActor(actor).then(items => {
                        actor.collabItems = items;
                    })
                ]).done(() => {
                    deferred.resolve(actor);
                });
            });

            return deferred.promise;
        }

        private getPayload(query: string, graphQuery: string) {
            return {
                "request": {
                    "Querytext": query,
                    "RowLimit": 500,
                    "TrimDuplicates": false,
                    "RankingModelId": "0c77ded8-c3ef-466d-929d-905670ea1d72",
                    //title,write,path,created,AuthorOWSUSER,EditorOWSUSER
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
        }

        private getPayloadActor(query: string, graphQuery: string) {
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
                    actor.id = parseInt(cell.Value);
                } else if (cell.Key === 'AccountName') {
                    actor.accountName = cell.Value;
                }
            }
            return actor;
        }

        private parseItemResults(row): Item {
            var item = new Item();
            for (var i = 0; i < row.Cells.length; i++) {
                var cell = row.Cells[i];
                if (cell.Key === 'Title') {
                    item.title = cell.Value;
                } else if (cell.Key === 'AuthorOWSUSER') {
                    item.createdBy = cell.Value;
                } else if (cell.Key === 'EditorOWSUSER') {
                    item.lastModifiedByAccount = cell.Value;
                } else if (cell.Key === 'ModifiedBy') {
                    item.lastModifiedByName = cell.Value;
                } else if (cell.Key === 'DocId') {
                    item.id = parseInt(cell.Value);
                } else if (cell.Key === 'Write') {
                    item.lastModifiedDate = moment(cell.Value).toDate();
                } else if (cell.Key === 'Created') {
                    item.createdDate = moment(cell.Value).toDate();
                } else if (cell.Key === 'Edges') {
                    //get the highest edge weight
                    var edges = JSON.parse(cell.Value);
                    item.rawEdges = this.parseEdgeResults(edges);
                }
            }
            return item;
        }

        private parseEdgeResults(inputEdges): Edge[] {
            var edges = [];
            for (var i = 0; i < inputEdges.length; i++) {
                var edge = new Edge();
                edge.actorId = inputEdges[i].ActorId;
                edge.objectId = inputEdges[i].ObjectId;
                var actionString = <string>inputEdges[i].Properties.Action;
                edge.action = Action[actionString];
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
        }
    }
}