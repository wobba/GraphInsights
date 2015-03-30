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
    
    <SharePoint:ScriptLink name="clienttemplates.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink name="clientforms.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink name="clientpeoplepicker.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink name="autofill.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink name="sp.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink name="sp.runtime.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink name="sp.core.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <meta name="WebPartPageExpansion" content="full" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />



    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>

    <style>
        .link {
            stroke: #aaa;
            stroke-width: 2px;
        }
        .node {
            stroke: #fff;
            stroke-width: 2px;
        }

        .textClass {
            stroke: #323232;
            font-family: "Lucida Grande", "Droid Sans", Arial, Helvetica, sans-serif;
            font-weight: normal;
            stroke-width: .5;
            font-size: 14px;
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
        Pick start actor:<div id="peoplePickerDiv"></div>
        <input type="hidden" id="actorId"/>
    </div>
    <div style="margin-top: 10px; margin-bottom: 10px">
        Colleague reach #
        <input id="colleagueReach" type="text" value="30" />
        <input type="button" onclick=" Pzl.OfficeGraph.Insight.initializePage(jQuery('#colleagueReach').val()); return false; " value="Kick it!" />
        Filter <input id="slider1" type="range" min="0" max="10" step="1" value="0" onchange="Pzl.OfficeGraph.Insight.hideSingleCollab(this.value)" onmousemove="Pzl.OfficeGraph.Insight.hideSingleCollab(this.value)" /> 
        <%--<input type="button" onclick="Pzl.OfficeGraph.Insight.hideSingleCollab(1); return false; " value="Remove single item collaborators" />--%>
    </div>

    <div style="width: 1000px">
        <div id="forceGraph" style="width: 100%; height: 600px; border: solid 1px #aaa; clear: both"></div>
        <div id="message" style="float: left; width: 50%">
        </div>
        <div id="log" style="float: left; width: 45%; height: 400px; border: solid 1px black; overflow: scroll">
        </div>
    </div>

</asp:Content>
