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
         edges: Edge[];

        //constructor() {
        //    this.id = workId;
        //}

        getNumberOfModifications() {
            var count = 0;
            var edges = this.edges;
            for (var edge in edges) {
                if (edges.hasOwnProperty(edge)) {
                    if (edge.action === Action.Modified) {
                        count += edge.weight;
                    }
                }
            }
            return count;
        }

        private getMinEdgeDate() {
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
        }

        private getMaxEdgeDate() {
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
        }

        getModificationsPerDay() {
            var start = this.getMinEdgeDate();
            var end = this.getMaxEdgeDate();
            var ms = moment(end).diff(moment(start));
            var d = moment.duration(ms);
            var days = d.days();
            var mods = this.getNumberOfModifications();
            return Math.round(mods / days);
        }
    }

    export class Edge {
        actorId: number;
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
        Colleague =1015,
        WorkingWithPublic = 1033,
        Manager = 1013
    }
}