﻿/// <reference path="typings/moment/moment.d.ts" />
"use strict";

module Pzl.OfficeGraph.Insight {
    export class Item {
        id : number;
        createdBy: string;
        lastModifiedBy: string;
        createdDate: Date;
        lastModifiedDate: Date;
        rawEdges: Edge[];

        getNumberOfEditsByActor(actor: Actor, mode : Inclusion): number {
            var edits = 0;
            for (var i = 0; i < this.rawEdges.length; i++) {
                var edge = this.rawEdges[i];
                if ((mode === Inclusion.ActorOnly && edge.actorId === actor.id)
                    || (mode === Inclusion.AllButActor && edge.actorId !== actor.id)) {
                    edits = edits + edge.weight;
                }
            }
            return edits;
        }

        getNumberOfContributors(): number {
            return this.rawEdges.length;
        }

        actorIsCreator(actor: Actor): boolean {
            return this.createdBy.indexOf(actor.accountName) >= 0;
        }

        actorIsLastModifed(actor: Actor): boolean {
            return this.lastModifiedBy.indexOf(actor.accountName) >= 0;
        }

        getMinDateEdge(): Date {
            var date = new Date(2099, 12, 31);
            for (var i = 0; i < this.rawEdges.length; i++) {
                if (this.rawEdges[i].time < date) {
                    date = this.rawEdges[i].time;
                }
            }
            return date;
        }

        getMaxDateEdge(): Date {
            var date = new Date(1970, 1, 1);
            for (var i = 0; i < this.rawEdges.length; i++) {
                if (this.rawEdges[i].time > date) {
                    date = this.rawEdges[i].time;
                }
            }
            return date;
        }

        itemLifeSpanInDays() : number {
            var ms = moment(this.lastModifiedDate).diff(moment(this.createdDate));
            var d = moment.duration(ms);
            return d.days();
        }
    }

    export enum Inclusion {
        ActorOnly,
        AllButActor
    }

} 