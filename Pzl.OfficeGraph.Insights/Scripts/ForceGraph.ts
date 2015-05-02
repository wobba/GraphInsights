/// <reference path="typings/d3/d3.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 

"use strict";

declare module D3 {
    export interface Base {
        tip: any;
    }
}

module Pzl.OfficeGraph.Insight.Graph {
    declare var d3: D3.Base;
    var graph;

    export class MyGraph {
        addLink;
        removeAllNodes;
        removeallLinks;
        removeLink;
        removeNode;
        showFilterByCount;
        isSingleNode;
        linkCountNode;
        validCssName;
        addNode;
        getLinks;
        highlightNode;
        resetNode;
        links;
        nodes;
        updateCallback;
        maxCount; // max number of collabs
        maxCountB: number = 1;
        hideCount: number = 0;

        constructor(domId: string) {
            var findNodeIndex = id => {
                for (var i = 0; i < this.nodes.length; i++) {
                    if (this.nodes[i].id == id) {
                        return i;
                    }
                }
                return -1;
            };

            var reCssPattern = /[^a-zA-Z0-9]/g;
            this.validCssName = (name: string) => name.replace(reCssPattern, "");

            this.isSingleNode = (source: string, hideCount: number) => {
                var count = 0;
                for (var i = 0; i < this.links.length; i++) {
                    if ((this.links[i].source.id === source || this.links[i].target.id === source) && this.links[i].count > hideCount) {
                        count++;
                        //return true;
                    }
                }
                return count === 0;
            };

            this.linkCountNode = (source: string) => {
                var count = 0;
                for (var i = 0; i < this.links.length; i++) {
                    if ((this.links[i].source.id === source || this.links[i].target.id === source)) {
                        count++;
                    }
                }
                return count;
            };

            var lastState = "";

            this.highlightNode = function (node, highlightClass: string, opacity: number) {
                for (var i = this.links.length - 1; i >= 0; i--) {
                    var link = this.links[i];
                    var id = "line#" + this.validCssName(link.source.id + "-" + link.target.id);
                    if ((link.source.id === node.id || link.target.id === node.id) && link.count > this.hideCount) {
                        d3.select(id).transition().style("opacity", 1)
                            .attr("class", highlightClass);
                    } else if (link.count > this.hideCount) {
                        d3.select(id).transition().style("opacity", opacity)
                            .attr("class", "link");
                    }
                }
                if (lastState !== (node.id + opacity)) {
                    lastState = (node.id + opacity);
                    if (opacity !== 1) {
                        this.updateCallback(node.actorId);
                    } else {
                        this.updateCallback(0);
                    }                    
                }

            };

            this.showFilterByCount = function (hideCount) {
                this.hideCount = hideCount;
                var animDuration = 250;
                for (var i = this.links.length - 1; i >= 0; i--) {
                    var link = this.links[i];
                    //console.log(link.source.id + ":" + link.target.id + ":" + link.value + ":" + link.count);
                    var id = "line#" + this.validCssName(link.source.id + "-" + link.target.id);
                    if (link.count <= hideCount) {
                        //this.removeLink(link.source.id, link.target.id); //TODO: perhaps save in a list and re-add
                        d3.select(id).transition().duration(animDuration).style("opacity", 0);
                    } else {
                        d3.select(id).transition().duration(animDuration).style("opacity", 1);
                    }
                }

                for (var j = 0; j < this.nodes.length; j++) {
                    var node = this.nodes[j];
                    var selectorNode = "#Node" + this.validCssName(node.id);
                    var selectorText = "#NodeText" + this.validCssName(node.id);
                    if (this.isSingleNode(node.id, hideCount)) {
                        d3.select(selectorNode).transition().duration(animDuration).style("opacity", 0); // hide links
                        d3.select(selectorText).transition().duration(animDuration).style("opacity", 0); // hide label
                    } else {
                        d3.select(selectorNode).transition().duration(animDuration).style("opacity", 1); // show label
                        d3.select(selectorText).transition().duration(animDuration).style("opacity", 1); // show label
                    }
                }
                update();
            }

            // Add and remove elements on the graph object
            this.addNode = (name,actorId) => {
                var idx = findNodeIndex(name);
                if (idx === -1) {
                    this.nodes.push({ "id": name, "actorId": actorId });
                    update();
                }
            };

            this.removeNode = id => {
                var i = 0;
                var n = findNode(id);
                while (i < this.links.length) {
                    if ((this.links[i]['source'] == n) || (this.links[i]['target'] == n)) {
                        this.links.splice(i, 1);
                    }
                    else i++;
                }
                this.nodes.splice(findNodeIndex(id), 1);
                update();
            };

            this.removeLink = (source, target) => {
                for (var i = 0; i < this.links.length; i++) {
                    if (this.links[i].source.id == source && this.links[i].target.id == target) {
                        this.links.splice(i, 1);
                        break;
                    }
                }
                update();
            };

            this.removeallLinks = () => {
                this.links.splice(0, this.links.length);
                update();
            };

            this.removeAllNodes = () => {
                this.nodes.splice(0, this.links.length);
                update();
            };

            this.maxCount = function () {
                return this.maxCountB;
            }

            this.addLink = (source: string, target: string, value: number) => {
                if (source > target) {
                    // sort names
                    var temp = target;
                    target = source;
                    source = temp;
                }

                var linkNodeA;
                var linkNodeB;

                var found = false;
                for (var i = 0; i < this.links.length; i++) {
                    // links are the same if source/target are the same
                    if ((this.links[i].source.id === source && this.links[i].target.id === target)
                        || (this.links[i].source.id === target && this.links[i].target.id === source)
                        ) {
                        found = true;
                        this.links[i].count += 1; // keep track of number of collabs between actors
                        // existing link - shorten to show closeness
                        if (this.links[i].value > 100) {
                            this.links[i].value = this.links[i].value / 2;
                        }

                        if (this.links[i].count > this.maxCountB) {
                            this.maxCountB = this.links[i].count;
                        }

                        value = this.links[i].value;

                        linkNodeA = this.links[i].source;
                        linkNodeB = this.links[i].target;
                        break;
                    }
                }
                if (!found) {
                    linkNodeA = findNode(source);
                    linkNodeB = findNode(target);
                    this.links.push({ "source": linkNodeA, "target": linkNodeB, "value": value, "count": 1 });
                }
                
                // set A node size
                var id = "#Node" + this.validCssName(linkNodeA.id);
                var numberOfLinks = this.linkCountNode(linkNodeA.id) - 1;
                numberOfLinks = Math.min(16, numberOfLinks);
                var r = 16 * (1 + numberOfLinks / 16);                
                d3.select(id).attr("r", r);


                // set B node size
                id = "#Node" + this.validCssName(linkNodeB.id);
                numberOfLinks = this.linkCountNode(linkNodeB.id) - 1;
                numberOfLinks = Math.min(16, numberOfLinks);
                r = 16 * (1 + numberOfLinks / 16);          
                d3.select(id).attr("r", r);

                update();
            };

            var findNode = id => {
                for (var i = 0; i < this.nodes.length; i++) {
                    if (this.nodes[i]["id"] === id) return this.nodes[i];
                }
                return null;
            }

            // rescale g
            function rescale() {
                var trans = d3.event.translate;
                var scale = d3.event.scale;

                vis.attr("transform",
                    "translate(" + trans + ")"
                    + " scale(" + scale + ")");
            }

            var w = jQuery("#" + domId).width();
            var h = jQuery("#" + domId).height();
            var r = 16;

            var color = d3.scale.category20();

            var vis = d3.select("#" + domId)
                .append("svg:svg")
                .attr("width", w)
                .attr("height", h)
                .attr("id", "svg")
                .attr("pointer-events", "all")
                .attr("viewBox", "0 0 " + w + " " + h)
                .attr("perserveAspectRatio", "xMinYMid")
                .append('svg:g')
            //.call(d3.behavior.zoom().on("zoom", rescale))
                ;

            var force = d3.layout.force();

            this.nodes = force.nodes();
            this.links = force.links();

            var fadeinTime = 500;

            var update = () => {
                var link = vis.selectAll("line")
                    .data(this.links, d => (d.source.id + "-" + d.target.id));

                link.enter().append("line")
                    .attr("id", d => (this.validCssName(d.source.id + "-" + d.target.id)))
                    .attr("stroke-width", d => (d.value / 10))
                    .attr("class", "link linkHidden")
                    .transition().duration(fadeinTime).style("opacity", 1);

                //d3.selectAll(id).transition().duration(animDuration).style("opacity", 1);
                link.append("title")
                    .text(d => d.value);
                link.exit().remove();

                var node = vis.selectAll("g.node")
                    .data(this.nodes, d => d.id);

                var nodeEnter = node.enter().append("g")
                    .attr("class", "node")
                    .call(force.drag);

                //nodeEnter.filter(d=> { return this.isSingleNode(d.id, 1) }).append("svg:circle")
                //    .attr("r", r)
                //    .attr("id", d => ("Node" + this.validCssName(d.id)))
                //    .attr("class", "nodeStrokeClass")
                //    .attr("fill", d => color(d.id))
                //    .transition().duration(fadeinTime).style("opacity", 1);

                //nodeEnter.filter(d=> { return !this.isSingleNode(d.id, 1) }).append("svg:circle")
                //    .attr("r", r)
                //    .attr("id", d => ("Node" + this.validCssName(d.id)))
                //    .attr("class", "nodeStrokeClass")
                //    .attr("fill", d => color(d.id))
                //    .style("opacity", 0);


                nodeEnter.append("svg:circle")
                    .attr("r", r)
                    .attr("id", d => ("Node" + this.validCssName(d.id)))
                    .attr("class", "nodeStrokeClass")
                    .attr("fill", d => color(d.id))
                    .transition().duration(fadeinTime).style("opacity", 1);

                nodeEnter.append("svg:text")
                    .attr("class", "textClass")
                    .attr("id", d => ("NodeText" + this.validCssName(d.id)))
                    .attr("x", 18)
                    .attr("y", ".31em")
                    .transition().duration(fadeinTime).style("opacity", 1)
                    .text(d => d.id);

                node.exit().remove();

                node.on("mousedown", d => {
                    this.highlightNode(d, "linkHightLight", .2);
                    //jQuery("#lala").css({ top: (d.y + 20), left: (d.x + 40) }).show();
                    //put actor image + data
                })
                    .on("mouseup", d => {
                    //jQuery("#lala").hide();
                    this.highlightNode(d, "link", 1);
                }).on("mouseout", d => {
                    //jQuery("#lala").hide();
                    this.highlightNode(d, "link", 1);
                });

                force.on("tick",() => {
                    link.attr("x1", d => d.source.x)
                        .attr("y1", d => d.source.y)
                        .attr("x2", d => d.target.x)
                        .attr("y2", d => d.target.y);

                    node.attr("transform", d => {                        
                        // keep nodes inside canvas - code by mikael
                        //var move = 50;
                        //if (d.x < 0) {
                        //    d.x = move;
                        //}
                        //if (d.x > w) {
                        //    d.x = w - move; // to keep labels visible
                        //}
                        //if (d.y < 0) {
                        //    d.y = move;
                        //    d.x = d.x + move;
                        //}
                        //if (d.y > h) {
                        //    d.y = h - move;
                        //    d.x = d.x - move;
                        //}

                        return "translate(" + d.x + "," + d.y + ")";
                    });
                });

                //http://stackoverflow.com/questions/9901565/charge-based-on-size-d3-force-layout
                var k = Math.sqrt(20 / (w * h));

                // Restart the force layout.
                force
                    .charge(-10 / k)
                //.gravity(10 * k) 100
                    .gravity(30 * k)
                    .friction(0.2) //default 0.9
                    .linkDistance(d => d.value)
                    .size([w, h])
                    .start();
            };
            // Make it all go
            update();
        }
    }

    function initGraph(domId: string, updateCallbackFunc: (actorId: number) => void): MyGraph {
        graph = new MyGraph(domId);
        graph.updateCallback = updateCallbackFunc;
        return graph;
    }

    // because of the way the network is created, nodes are created first, and links second,
    // so the lines were on top of the nodes, this just reorders the DOM to put the svg:g on top
    export function keepNodesOnTop() {
        $(".nodeStrokeClass").each(function (index) {
            var gnode = this.parentNode;
            gnode.parentNode.appendChild(gnode);
        });
    }
    export function init(domId: string, updateCallbackFunc: (actorId: number) => void): MyGraph {
        d3.select("svg")
            .remove();        
        return initGraph(domId, updateCallbackFunc);
    }
}