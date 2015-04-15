/// <reference path="typings/d3/d3.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" /> 
"use strict";
var Pzl;
(function (Pzl) {
    var OfficeGraph;
    (function (OfficeGraph) {
        var Insight;
        (function (Insight) {
            var Graph;
            (function (Graph) {
                var graph;
                var MyGraph = (function () {
                    function MyGraph(domId) {
                        var _this = this;
                        this.maxCountB = 1;
                        var findNodeIndex = function (id) {
                            for (var i = 0; i < _this.nodes.length; i++) {
                                if (_this.nodes[i].id == id) {
                                    return i;
                                }
                            }
                            return -1;
                        };
                        var reCssPattern = /[^a-zA-Z0-9]/g;
                        this.validCssName = function (name) { return name.replace(reCssPattern, ""); };
                        this.isSingleNode = function (source, hideCount) {
                            var count = 0;
                            for (var i = 0; i < _this.links.length; i++) {
                                if ((_this.links[i].source.id === source || _this.links[i].target.id === source) && _this.links[i].count > hideCount) {
                                    count++;
                                }
                            }
                            return count === 0;
                        };
                        this.highlightNode = function (node, highlightClass, opacity) {
                            for (var i = this.links.length - 1; i >= 0; i--) {
                                var link = this.links[i];
                                var id = "line#" + this.validCssName(link.source.id + "-" + link.target.id);
                                if (link.source.id === node.id || link.target.id === node.id) {
                                    d3.select(id).transition().style("opacity", 1).attr("class", highlightClass);
                                }
                                else {
                                    d3.select(id).transition().style("opacity", opacity).attr("class", "link");
                                }
                            }
                        };
                        this.showFilterByCount = function (hideCount) {
                            var animDuration = 250;
                            for (var i = this.links.length - 1; i >= 0; i--) {
                                var link = this.links[i];
                                console.log(link.source.id + ":" + link.target.id + ":" + link.value + ":" + link.count);
                                var id = "line#" + this.validCssName(link.source.id + "-" + link.target.id);
                                if (link.count <= hideCount) {
                                    //this.removeLink(link.source.id, link.target.id); //TODO: perhaps save in a list and re-add
                                    d3.select(id).transition().duration(animDuration).style("opacity", 0);
                                }
                                else {
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
                                }
                                else {
                                    d3.select(selectorNode).transition().duration(animDuration).style("opacity", 1); // show label
                                    d3.select(selectorText).transition().duration(animDuration).style("opacity", 1); // show label
                                }
                            }
                            update();
                        };
                        // Add and remove elements on the graph object
                        this.addNode = function (id) {
                            var idx = findNodeIndex(id);
                            if (idx === -1) {
                                _this.nodes.push({ "id": id });
                                update();
                            }
                        };
                        this.removeNode = function (id) {
                            var i = 0;
                            var n = findNode(id);
                            while (i < _this.links.length) {
                                if ((_this.links[i]['source'] == n) || (_this.links[i]['target'] == n)) {
                                    _this.links.splice(i, 1);
                                }
                                else
                                    i++;
                            }
                            _this.nodes.splice(findNodeIndex(id), 1);
                            update();
                        };
                        this.removeLink = function (source, target) {
                            for (var i = 0; i < _this.links.length; i++) {
                                if (_this.links[i].source.id == source && _this.links[i].target.id == target) {
                                    _this.links.splice(i, 1);
                                    break;
                                }
                            }
                            update();
                        };
                        this.removeallLinks = function () {
                            _this.links.splice(0, _this.links.length);
                            update();
                        };
                        this.removeAllNodes = function () {
                            _this.nodes.splice(0, _this.links.length);
                            update();
                        };
                        this.maxCount = function () {
                            return this.maxCountB;
                        };
                        this.addLink = function (source, target, value) {
                            if (source > target) {
                                // sort names
                                var temp = target;
                                target = source;
                                source = temp;
                            }
                            var found = false;
                            for (var i = 0; i < _this.links.length; i++) {
                                // links are the same if source/target are the same
                                if ((_this.links[i].source.id === source && _this.links[i].target.id === target) || (_this.links[i].source.id === target && _this.links[i].target.id === source)) {
                                    found = true;
                                    _this.links[i].count += 1; // keep track of number of collabs between actors
                                    // existing link - shorten to show closeness
                                    if (_this.links[i].value > 50) {
                                        _this.links[i].value = _this.links[i].value / 2;
                                    }
                                    if (_this.links[i].count > _this.maxCountB) {
                                        _this.maxCountB = _this.links[i].count;
                                    }
                                    value = _this.links[i].value;
                                    break;
                                }
                            }
                            if ((target.indexOf("Elsa") !== -1 || source.indexOf("Elsa") !== -1) && (target.indexOf("Tormod") !== -1 || source.indexOf("Tormod") !== -1)) {
                                console.log(found + ":" + source + ":" + target + ":" + value);
                            }
                            if (!found) {
                                _this.links.push({ "source": findNode(source), "target": findNode(target), "value": value, "count": 1 });
                            }
                            update();
                        };
                        var findNode = function (id) {
                            for (var i = 0; i < _this.nodes.length; i++) {
                                if (_this.nodes[i]["id"] === id)
                                    return _this.nodes[i];
                            }
                            return null;
                        };
                        // rescale g
                        function rescale() {
                            var trans = d3.event.translate;
                            var scale = d3.event.scale;
                            vis.attr("transform", "translate(" + trans + ")" + " scale(" + scale + ")");
                        }
                        var w = jQuery("#" + domId).width();
                        var h = jQuery("#" + domId).height();
                        var r = 16;
                        var color = d3.scale.category20();
                        var vis = d3.select("#" + domId).append("svg:svg").attr("width", w).attr("height", h).attr("id", "svg").attr("pointer-events", "all").attr("viewBox", "0 0 " + w + " " + h).attr("perserveAspectRatio", "xMinYMid").append('svg:g');
                        var force = d3.layout.force();
                        this.nodes = force.nodes();
                        this.links = force.links();
                        var fadeinTime = 500;
                        var update = function () {
                            var link = vis.selectAll("line").data(_this.links, function (d) { return (d.source.id + "-" + d.target.id); });
                            link.enter().append("line").attr("id", function (d) { return (_this.validCssName(d.source.id + "-" + d.target.id)); }).attr("stroke-width", function (d) { return (d.value / 10); }).attr("class", "link linkHidden").transition().duration(fadeinTime).style("opacity", 1);
                            //d3.selectAll(id).transition().duration(animDuration).style("opacity", 1);
                            link.append("title").text(function (d) { return d.value; });
                            link.exit().remove();
                            var node = vis.selectAll("g.node").data(_this.nodes, function (d) { return d.id; });
                            var nodeEnter = node.enter().append("g").attr("class", "node").call(force.drag);
                            nodeEnter.append("svg:circle").attr("r", r).attr("id", function (d) { return ("Node" + _this.validCssName(d.id)); }).attr("class", "nodeStrokeClass").attr("fill", function (d) { return color(d.id); }).transition().duration(fadeinTime).style("opacity", 1);
                            nodeEnter.append("svg:text").attr("class", "textClass").attr("id", function (d) { return ("NodeText" + _this.validCssName(d.id)); }).attr("x", 18).attr("y", ".31em").transition().duration(fadeinTime).style("opacity", 1).text(function (d) { return d.id; });
                            node.exit().remove();
                            node.on("mousedown", function (d) {
                                _this.highlightNode(d, "linkHightLight", .2);
                                //jQuery("#lala").css({ top: (d.y + 20), left: (d.x + 40) }).show();
                                //put actor image + data
                            }).on("mouseup", function (d) {
                                //jQuery("#lala").hide();
                                _this.highlightNode(d, "link", 1);
                            }).on("mouseout", function (d) {
                                //jQuery("#lala").hide();
                                _this.highlightNode(d, "link", 1);
                            });
                            force.on("tick", function () {
                                link.attr("x1", function (d) { return d.source.x; }).attr("y1", function (d) { return d.source.y; }).attr("x2", function (d) { return d.target.x; }).attr("y2", function (d) { return d.target.y; });
                                node.attr("transform", function (d) {
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
                            force.charge(-10 / k).gravity(30 * k).friction(0.2).linkDistance(function (d) { return d.value; }).size([w, h]).start();
                        };
                        // Make it all go
                        update();
                    }
                    return MyGraph;
                })();
                Graph.MyGraph = MyGraph;
                function initGraph(domId) {
                    graph = new MyGraph(domId);
                    // callback for the changes in the network
                    //var step = -1;
                    //function nextval() {
                    //    step++;
                    //    return 2000 + (1500 * step); // initial time, wait time
                    //}
                    return graph;
                }
                // because of the way the network is created, nodes are created first, and links second,
                // so the lines were on top of the nodes, this just reorders the DOM to put the svg:g on top
                function keepNodesOnTop() {
                    $(".nodeStrokeClass").each(function (index) {
                        var gnode = this.parentNode;
                        gnode.parentNode.appendChild(gnode);
                    });
                }
                Graph.keepNodesOnTop = keepNodesOnTop;
                function init(domId) {
                    d3.select("svg").remove();
                    return initGraph(domId);
                }
                Graph.init = init;
            })(Graph = Insight.Graph || (Insight.Graph = {}));
        })(Insight = OfficeGraph.Insight || (OfficeGraph.Insight = {}));
    })(OfficeGraph = Pzl.OfficeGraph || (Pzl.OfficeGraph = {}));
})(Pzl || (Pzl = {}));
//# sourceMappingURL=ForceGraph.js.map