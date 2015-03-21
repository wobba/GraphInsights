/// <reference path="typings/moment/moment.d.ts" />
"use strict";

module Pzl.OfficeGraph.Insight {
    export class Actor {
        id: number;
        name: string;
        accountName: string;
        pictureUrl: string;
        gender: Gender;
        age: number;
        items: Item[];
        collabItems: Item[];
        associates: Actor[];

        getNumberOfModificationsByYou() {
            var count = 0;
            for (var i = 0; i < this.items.length; i++) {
                count = count + this.items[i].getNumberOfEditsByActor(this, Inclusion.ActorOnly);
            }
            return count;
        }

        getModificationsPerDay() {
            var start = this.getMinEdgeDate();
            var end = this.getMaxEdgeDate();
            var ms = moment(end).diff(moment(start));
            var d = moment.duration(ms);
            var days = d.days();
            var mods = this.getNumberOfModificationsByYou();
            return Math.round(mods / days);
        }

        private getMinEdgeDate() {
            var date = new Date(2099, 12, 31);
            for (var i = 0; i < this.items.length; i++) {
                var itemDate = this.items[i].getMinDateEdge();
                if (itemDate < date) {
                    date = itemDate;
                }
            }
            return date;
        }

        private getMaxEdgeDate() {
            var date = new Date(1970, 1, 1);
            for (var i = 0; i < this.items.length; i++) {
                var itemDate = this.items[i].getMaxDateEdge();
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