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
    <script type="text/javascript" src="../Scripts/jshashtable-3.0.js"></script>

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
    <SharePoint:ScriptLink Name="sp.requestexecutor.js" runat="server" LoadAfterUI="true" Localizable="false" />
    <meta name="WebPartPageExpansion" content="full" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    <div style="display: table-cell;vertical-align: middle">
        CollaboGraph by
        <img src="../Images/Puzzlepart_logo.png" style="height: 40px; position:relative; top:8px" />
    </div>
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <h2>All stats presented revolves around how people work together on items, seen from <b>your</b> point of view.
        <br />
        If there are clusters of documents you don't have access to they will not show up. Go poke an eye!</h2>
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
                <input type="button" onclick="filterSlider.value = 0; filterSlider.max = 1; Pzl.OfficeGraph.Insight.initializePage(jQuery('#colleagueReach').val()); return false;" value="Kick it!" class="kickIt" />
            </div>
        </div>
    </div>
    <div class="container">
        <div class="statsArea">
            <div id="message"></div>
            <img id="avatar"/>
        </div>
        <div class="graphArea">
            <div id="forceGraph">
                <div id="statusMessageArea" style="display:none">
                    <div id="statusMessage"></div>
                </div>
                <div id="lala" style="border: solid 1px #555; display: none; position: absolute; width: 200px; height: 200px; background-color: white"></div>
            </div>
        </div>
    </div>
    <div id="log">
    </div>
</asp:Content>
