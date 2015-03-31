<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/jquery-2.1.3.min.js"></script>
    <script type="text/javascript" src="../Scripts/moment.js"></script>
    <script type="text/javascript" src="../Scripts/q.js"></script>
    <script type="text/javascript" src="../Scripts/d3/d3.min.js"></script>

    <script type="text/javascript" src="../Scripts/Actor.js"></script>
    <script type="text/javascript" src="../Scripts/SearchHelper.js"></script>
    <script type="text/javascript" src="../Scripts/Item.js"></script>
    <script type="text/javascript" src="../Scripts/ForceGraph.js"></script>

    <SharePoint:ScriptLink Name="clienttemplates.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink Name="clientforms.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink Name="clientpeoplepicker.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink Name="autofill.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink Name="sp.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink Name="sp.runtime.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink Name="sp.core.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <meta name="WebPartPageExpansion" content="full" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />
    <link rel="Stylesheet" type="text/css" href="../Content/bootstrap.css" />



    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>

    <style>
        .link {
            stroke: #aaa;
            stroke-width: 2px;
            opacity: 0;
        }

        .node {
            stroke: #fff;
            stroke-width: 2px;
        }

        .nodeStrokeClass {
            opacity: 0;
        }

        .textClass {
            stroke: #323232;
            font-family: "Lucida Grande", "Droid Sans", Arial, Helvetica, sans-serif;
            font-weight: normal;
            stroke-width: .5;
            font-size: 14px;
            opacity: 0;
        }

        .container {
            width: 100%;
            min-width: 830px;
            /*border: 1px solid;*/
            clear: both;
        }

        .graphArea {
            width: auto;
            overflow: hidden;
        }

        .statsArea {
            width: 500px;
            /*background: blue;*/
            float: right;
            border: solid 1px #aaa;
            height: 600px;
        }

        #actionControls {
            min-width: 830px;
            height: 80px;
            clear: both;
        }

        #forceGraph {
            width: 100%;
            height: 600px;
            border: solid 1px #aaa;
        }

        #message {
            /*margin-top: 80px;*/
            margin-left: 10px;
            margin-right: 10px;
        }

        .sp-peoplepicker-topLevel {
            width: 200px;
        }

        .controlItem {
            float: left;
            margin-right: 20px;
        }

        input.kickIt {
            width: 200px;
            height: 70px;
            font-size: 30px;
        }

        .controlNudge {
            margin-top: 7px;
        }

        #log {
            margin-top: 50px;
            width: 100%;
            height: 200px;
            border: solid 1px #aaa;
            overflow: scroll;
        }
    </style>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    Pzl Edge Insights aka Co-Auth Monitor
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <h2>All stats presented in this App revolves around how people work together on documents, seen from <b>your</b> point of view. If there are clusters of documents you don't have access to they will not show up. Go poke an eye!</h2>
    <hr />
    <div>
        <div id="actionControls">
            <div class="controlItem">
                Pick start actor:<div id="peoplePickerDiv" class="controlNudge"></div>
                <input type="hidden" id="actorId" />
            </div>
            <div class="controlItem">
                Colleague reach #
                    <div>
                        <input id="colleagueReach" type="text" value="30" class="controlNudge" />
                    </div>
            </div>
            <div class="controlItem">
                <div>Filter</div>
                <span>0</span><input id="filterSlider" type="range" min="0" max="1" step="1" value="0" onchange="Pzl.OfficeGraph.Insight.hideSingleCollab(this.value)" onmousemove="Pzl.OfficeGraph.Insight.hideSingleCollab(this.value)" list="steplist" /><span id="maxValue">1</span>
                <datalist id="steplist" />
            </div>
            <div class="controlItem">
                <input type="button" onclick=" Pzl.OfficeGraph.Insight.initializePage(jQuery('#colleagueReach').val()); return false; " value="Kick it!" class="kickIt" />
            </div>
        </div>
    </div>
    <div class="container">
        <div class="statsArea">
            <div id="message"></div>
        </div>
        <div class="graphArea">
            <div id="forceGraph"></div>
        </div>
    </div>
    <div id="log">
    </div>

</asp:Content>
