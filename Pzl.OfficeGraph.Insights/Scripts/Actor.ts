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

        // Number of documents edited by actor only
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

        // Find oldest created date with more than two authors
        getLongestLivingItemWithCollab(): Item {
            var oldestItem: Item;
            if (this.collabItems) {
                for (var i = 0; i < this.collabItems.length; i++) {
                    var item = this.collabItems[i];
                    if (item.getNumberOfContributors() > 1) {
                        if (oldestItem === undefined || item.itemLifeSpanInDays() > oldestItem.itemLifeSpanInDays()) {
                            oldestItem = item;
                        }
                    }
                }
            }
            return oldestItem;
        }

        getStarterCount(): number {
            var creatorCount = 0;
            if (this.collabItems) {
                for (var i = 0; i < this.collabItems.length; i++) {
                    var item = this.collabItems[i];
                    if (item.getNumberOfContributors() > 1) {
                        if (item.actorIsCreator(this)) {
                            creatorCount++;
                        }
                    }
                }
            }
            return creatorCount;
        }

        getLastSaverCount(): number {
            var saverCount = 0;
            if (this.collabItems) {
                for (var i = 0; i < this.collabItems.length; i++) {
                    var item = this.collabItems[i];
                    if (item.getNumberOfContributors() > 1) {
                        if (item.actorIsLastModifed(this)) {
                            saverCount++;
                        }
                    }
                }
            }
            return saverCount;
        }

        // Get item you have most saves for
        getHighestItemSaveCount(): number {
            var count = 0;
            if (this.collabItems) {
                for (var i = 0; i < this.collabItems.length; i++) {
                    var item = this.collabItems[i];
                    var itemCount = item.getMaxSaveCountforActor(this);
                    if (itemCount > count) {
                        count = itemCount;
                    }
                }
            }
            return count;
        }

        getAssociateNameById(actorId: number): string {
            if (actorId === this.id) {
                return this.name;
            }
            for (var i = 0; i < this.associates.length; i++) {
                if (this.associates[i].id === actorId) {
                    return this.associates[i].name;
                }
            }
            return actorId.toString();
        }
    }

    export class Edge {
        workid: number;
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