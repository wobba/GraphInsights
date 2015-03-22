/// <reference path="typings/moment/moment.d.ts" />
///<reference path="typings/jquery/jquery.d.ts" /> 
"use strict";

module Pzl.OfficeGraph.Insight {
    export class Actor {
        id: number;
        name: string;
        accountName: string;
        pictureUrl: string;
        gender: Gender;
        age: number;
        //items: Item[];
        collabItems: Item[];
        associates: Actor[];

        // Average number of recorded saves per item
        getItemModificationsAverage(): number {
            var count = 0;
            for (var i = 0; i < this.collabItems.length; i++) {
                count = count + this.collabItems[i].getNumberOfEditsByActor(this, Inclusion.ActorOnly);
            }
            return Math.round(count / this.collabItems.length);
        }

        // Average number of recorded saves per item
        getEgoSaveCount(): number {
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
        }

        //getModificationsPerDay(): number {
        //    var start = this.getMinEdgeDate();
        //    var end = this.getMaxEdgeDate();
        //    var ms = moment(end).diff(moment(start));
        //    var d = moment.duration(ms);
        //    var days = d.days();
        //    if (days === 0) { days = 1 };
        //    var mods = this.getNumberOfModificationsByYou();
        //    //return Math.round(mods / days);
        //    console.log(days + ":" + mods + " - " + start + ":" + end);
        //    return mods / days;
        //}

        getCollaborationRatio(): number {
            var meOnly = 0;
            var all = 0;
            if (this.collabItems) {
                for (var i = 0; i < this.collabItems.length; i++) {
                    var item = this.collabItems[i];
                    if (item.getNumberOfContributors() === 1) {
                        meOnly++;
                    } else {
                        all++;
                    }
                }
            }
            return meOnly / all;
        }

        // Item count with at least 2 authors
        getCollaborationItemCount(): number {
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
        }

        // Get all actors a user collaborates with
        getCollaborationActorCount(): number {
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
        }

        private getMinEdgeDate() {
            var date = new Date(2099, 12, 31);
            for (var i = 0; i < this.collabItems.length; i++) {
                var itemDate = this.collabItems[i].getMinDateEdge(this.id);
                if (itemDate < date) {
                    date = itemDate;
                }
            }
            return date;
        }

        private getMaxEdgeDate() {
            var date = new Date(1970, 1, 1);
            for (var i = 0; i < this.collabItems.length; i++) {
                var itemDate = this.collabItems[i].getMaxDateEdge();
                if (itemDate > date) {
                    date = itemDate;
                }
            }
            return date;
        }
    }

    export class Edge {
        actorId: number;
        objectId: number;
        action: Action;
        time: Date;
        weight: number;
    }

    export enum Gender {
        Male,
        Female
    }

    export enum Action {
        Modified = 1003,
        Colleague = 1015,
        WorkingWithPublic = 1033,
        Manager = 1013
    }
}