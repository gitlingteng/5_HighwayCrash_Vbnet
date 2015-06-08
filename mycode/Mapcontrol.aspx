<%@ Page Language="vb" AutoEventWireup="false" Inherits="Crashsafe.MapPage" CodeFile="Mapcontrol.aspx.vb" EnableEventValidation="false" %>

<%@ Register Assembly="ESRI.ArcGIS.ADF.Web.UI.WebControls, Version=9.2.6.1500, Culture=neutral, PublicKeyToken=8fc3cc631e44ad86"
    Namespace="ESRI.ArcGIS.ADF.Web.UI.WebControls" TagPrefix="esri" %>        
<%@ Register Assembly="ESRI.ArcGIS.ADF.Web.UI.WebControls, Version=9.2.0.1324, Culture=neutral, PublicKeyToken=8fc3cc631e44ad86"
    Namespace="ESRI.ArcGIS.ADF.Web.UI.WebControls" TagPrefix="esri" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Untitled Page</title>
    <meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR"/>
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE"/>
		<meta content="JavaScript" name="vs_defaultClientScript"/>
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema"/>
		<meta http-equiv="imagetoolbar" content="false"/>
		<script language="JavaScript" src="javascript/ZoomBox.js" type="text/javascript"></script>
		<script language="JavaScript" src="javascript/Point.js" type="text/javascript"></script>
		<script language="JavaScript" src="javascript/Rect.js" type="text/javascript"></script>
		<script language="JavaScript" src="javascript/Map.js" type="text/javascript"></script>
		<script language="JavaScript" src="javascript/functions.js" type="text/javascript"></script>
		<script language="JavaScript" src="javascript/main.js" type="text/javascript"></script>
	
</head>
<body>

<script language ="javascript" type ="text/javascript">
 var subwin;
  
       
       function SelectRow(index)
{

     if (subwin !=null && !(subwin.closed)) subwin.close();
       subwin = window.open('Information.aspx?CMD=INIT&amp;X='+index,"subwin",'width =250,heigh =650,left =800,top =10,scrollbars =yes');
       subwin.focus();
     


//subWin = window.open('Information.aspx?CMD=INIT&amp;X='+index+,"form","width=250,height=550,left=800,top=10,scrollbars = yes");
 //('Prospect.aspx?CMD=INIT&amp;X='+trueX+'&amp;Y='+trueY,"","width=475,height=540,left=530,top=30") 
}

function MyFunction()
{

  //if (subwin !=null && !(subwin.closed)) subwin.close();
  window.close('Information.aspx');
}

</script>



    <form id="form1" runat="server">
    <div>
        &nbsp;</div>
        &nbsp; &nbsp;&nbsp;
        
        
			<asp:label id="Label1" style="Z-INDEX: 100; LEFT: -134217728px; POSITION: absolute; TOP: 40px" runat="server" BorderColor="#FFC0FF" ForeColor="GreenYellow" Width="10px" Height="40px">111111111111111111111111111</asp:label><asp:checkbox id="CheckBox1" style="Z-INDEX: 101; LEFT: 502px; POSITION: absolute; TOP: 40px" runat="server" Text="Accs" BackColor="Transparent" Width="16px" Font-Size="XX-Small"></asp:checkbox><asp:button id="Button2" style="Z-INDEX: 102; LEFT: 662px; POSITION: absolute; TOP: 43px" runat="server" Font-Bold="True" Text="Ok" BackColor="Goldenrod" BorderColor="LemonChiffon" ForeColor="#FFFF80" Width="20px" Height="20px" Font-Size="XX-Small"></asp:button><asp:button id="Btn5" style="Z-INDEX: 104; LEFT: 968px; POSITION: absolute; TOP: 40px" runat="server" Font-Bold="True" Text="Ok" BackColor="Goldenrod" BorderColor="LemonChiffon" ForeColor="#FFFF80" Width="20px" Height="20px" Font-Size="XX-Small"></asp:button><asp:label id="Label3" style="Z-INDEX: 106; LEFT: 232px; POSITION: absolute; TOP: 504px" runat="server" Visible="False">Label</asp:label><asp:dropdownlist id="cmbYear" style="Z-INDEX: 107; LEFT: 590px; POSITION: absolute; TOP: 43px" runat="server" Width="64px" Font-Size="X-Small" Visible="False">
				<asp:ListItem Value="2000">2000</asp:ListItem>
				<asp:ListItem Value="2001">2001</asp:ListItem>
				<asp:ListItem Value="2002">2002</asp:ListItem>
				<asp:ListItem Value="2003">2003</asp:ListItem>
				<asp:ListItem Value="2004">2004</asp:ListItem>
					<asp:ListItem Value="2005">2005</asp:ListItem>
											<asp:ListItem Value="2006">2006</asp:ListItem>
											
			</asp:dropdownlist><asp:button id="Label4" style="Z-INDEX: 108; LEFT: 24px; POSITION: absolute; TOP: 24px" runat="server" BackColor="Transparent" ForeColor="#C04000" Font-Size="X-Small" BorderStyle="None" EnableViewState="False"></asp:button><asp:label id="Labelnew" style="Z-INDEX: 109; LEFT: 24px; POSITION: absolute; TOP: 0px" runat="server" Font-Bold="True" BackColor="#C0FFC0" BorderColor="#FFC0FF" ForeColor="Red" Width="973px" Height="40px" Font-Size="X-Small"></asp:label>
        <esri:Map ID="Map1" runat="server" Height="392px" MapResourceManager="MapResourceManager1"
            Width="552px" style="left: 8px; position: absolute; top: 64px; z-index: 110;" InitialExtent="Full">
        </esri:Map>
   <%--     <esri:Toc ID="Toc1" runat="server" BuddyControl="Map1" Style="left: 576px; position: absolute;
            top: 64px; z-index: 111;" Height="168px" Width="120px" />
  --%>      <esri:Toolbar ID="Toolbar1" runat="server" BuddyControlType="Map" Group="Toolbar1_Group"
            Height="64px" ToolbarItemDefaultStyle-BackColor="White" ToolbarItemDefaultStyle-Font-Names="Arial"
            ToolbarItemDefaultStyle-Font-Size="Smaller" ToolbarItemDisabledStyle-BackColor="White"
            ToolbarItemDisabledStyle-Font-Names="Arial" ToolbarItemDisabledStyle-Font-Size="Smaller"
            ToolbarItemDisabledStyle-ForeColor="Gray" ToolbarItemHoverStyle-BackColor="White"
            ToolbarItemHoverStyle-Font-Bold="True" ToolbarItemHoverStyle-Font-Italic="True"
            ToolbarItemHoverStyle-Font-Names="Arial" ToolbarItemHoverStyle-Font-Size="Smaller"
            ToolbarItemSelectedStyle-BackColor="White" ToolbarItemSelectedStyle-Font-Bold="True"
            ToolbarItemSelectedStyle-Font-Names="Arial" ToolbarItemSelectedStyle-Font-Size="Smaller"
            WebResourceLocation="/aspnet_client/ESRI/WebADF/" Width="464px" style="left: 8px; position: absolute; top: 464px; z-index: 111;">
            <ToolbarItems>
                <esri:Tool ClientAction="DragRectangle" DefaultImage="~/Images/zoomin.gif"
                    HoverImage="~/Images/zoominU.gif" JavaScriptFile=""
                    Name="MapZoomIn" SelectedImage="~/Images/zoominD.gif"
                    ServerActionAssembly="ESRI.ArcGIS.ADF.Web.UI.WebControls" ServerActionClass="ESRI.ArcGIS.ADF.Web.UI.WebControls.Tools.MapZoomIn"
                    Text="Zoom In" ToolTip="Zoom In" />
                <esri:Tool ClientAction="DragRectangle" DefaultImage="~/Images/zoomout.GIF"
                    HoverImage="~/Images/zoomoutU.gif" JavaScriptFile=""
                    Name="MapZoomOut" SelectedImage="~/Images/zoomoutD.gif"
                    ServerActionAssembly="ESRI.ArcGIS.ADF.Web.UI.WebControls" ServerActionClass="ESRI.ArcGIS.ADF.Web.UI.WebControls.Tools.MapZoomOut"
                    Text="Zoom Out" ToolTip="Zoom Out" />
                <esri:Tool ClientAction="DragImage" DefaultImage="~/Images/pan.gif"
                    HoverImage="~/Images/panU.gif" JavaScriptFile=""
                    Name="MapPan" SelectedImage="~/Images/panD.gif"
                    ServerActionAssembly="ESRI.ArcGIS.ADF.Web.UI.WebControls" ServerActionClass="ESRI.ArcGIS.ADF.Web.UI.WebControls.Tools.MapPan"
                    Text="Pan" ToolTip="Pan" />
                <esri:Command ClientAction="" DefaultImage="~/Images/fullext.gif"
                    HoverImage="~/Images/fullextU.gif" JavaScriptFile=""
                    Name="MapFullExtent" SelectedImage="~/Images/fullextD.gif"
                    ServerActionAssembly="ESRI.ArcGIS.ADF.Web.UI.WebControls" ServerActionClass="ESRI.ArcGIS.ADF.Web.UI.WebControls.Tools.MapFullExtent"
                    Text="Full Extent" ToolTip="Full Extent" />
                 <esri:Tool ClientAction="Point" DefaultImage="~/Images/pan.gif"
                    HoverImage="~/Images/panU.gif" SelectedImage="~/Images/panD.gif" JavaScriptFile="" Name="BufferSelect" ServerActionAssembly="App_Code"
                    ServerActionClass="BufferTool" Text="Select Features" ToolTip="Select Features" />
           
            </ToolbarItems>
            <BuddyControls>
                <esri:BuddyControl Name="Map1" />
            </BuddyControls>
        </esri:Toolbar>
        <esri:MapResourceManager ID="MapResourceManager1" runat="server" style="left: 80px; top: 104px" Height="32px" Width="128px">
            <ResourceItems>
                <esri:MapResourceItem DisplaySettings="visible=True:transparency=0:mime=True:imgFormat=PNG8:height=100:width=100:dpi=96:color=:transbg=False:displayInToc=True"
                    Name="Buffer" Definition="&lt;Definition DataSourceDefinition=&quot;In Memory&quot; DataSourceType=&quot;GraphicsLayer&quot; Identity=&quot;&quot; ResourceDefinition=&quot;&quot; DataSourceShared=&quot;True&quot; /&gt;" />
                <esri:MapResourceItem Definition="&lt;Definition DataSourceDefinition=&quot;In Memory&quot; DataSourceType=&quot;GraphicsLayer&quot; Identity=&quot;&quot; ResourceDefinition=&quot;&quot; DataSourceShared=&quot;True&quot; /&gt;"
                    DisplaySettings="visible=True:transparency=0:mime=True:imgFormat=PNG8:height=100:width=100:dpi=96:color=:transbg=False:displayInToc=True"
                    Name="Selection" />
                <esri:MapResourceItem Definition="&lt;Definition DataSourceDefinition=&quot;In Memory&quot; DataSourceType=&quot;GraphicsLayer&quot; Identity=&quot;&quot; ResourceDefinition=&quot;&quot; DataSourceShared=&quot;True&quot; /&gt;"
                    DisplaySettings="visible=True:transparency=0:mime=True:imgFormat=PNG8:height=100:width=100:dpi=96:color=:transbg=False:displayInToc=True"
                    Name="PartialSelect" />
                <esri:MapResourceItem Definition="&lt;Definition DataSourceDefinition=&quot;In Memory&quot; DataSourceType=&quot;GraphicsLayer&quot; Identity=&quot;&quot; ResourceDefinition=&quot;&quot; DataSourceShared=&quot;True&quot; /&gt;"
                    DisplaySettings="visible=True:transparency=0:mime=True:imgFormat=PNG8:height=100:width=100:dpi=96:color=:transbg=False:displayInToc=True"
                    Name="QuerySelection" />
                <esri:MapResourceItem Definition="&lt;Definition DataSourceDefinition=&quot;localhost&quot; DataSourceType=&quot;ArcGIS Server Local&quot; Identity=&quot;To set, right-click project and 'Add ArcGIS Identity'&quot; ResourceDefinition=&quot;(default)@ac2000&quot; DataSourceShared=&quot;True&quot; /&gt;"
                    DisplaySettings="visible=True:transparency=0:mime=True:imgFormat=PNG8:height=100:width=100:dpi=96:color=:transbg=False:displayInToc=True"
                    Name="Accident" />
            </ResourceItems>
        </esri:MapResourceManager>
        <asp:CheckBox ID="Showstreet" runat="server" Style="z-index: 120; left: 568px; position: absolute;
            top: 80px" Text="Show Highway" AutoPostBack="True" Height="19px" Width="140px" /><asp:CheckBox ID="ShowInt" runat="server" Style="z-index: 120; left: 569px; position: absolute;
            top: 111px" Text="Show Intersection" AutoPostBack="True" Height="10px" Width="154px" />
        <esri:Toc ID="Toc1" runat="server" BuddyControl="Map1" Height="192px" Style="z-index: 120;
            left: 568px; position: absolute; top: 64px" Width="136px" Visible="False" />
        <asp:button id="Button4" runat="server" Font-Bold="True" Text="Home" BackColor="CornflowerBlue" BorderColor="Aqua" ForeColor="#FFFF80" Width="64px" style="left: 720px; position: absolute; top: 432px; z-index: 113;"></asp:button><font face="ËÎÌו">&nbsp;</font>
        &nbsp;&nbsp;
        &nbsp;
    
        <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" Style="z-index: 114;
            left: 576px; position: absolute; top: 488px" Visible="False">
            <asp:ListItem>accident selection</asp:ListItem>
            <asp:ListItem>section selection</asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID="Button1" runat="server" Style="z-index: 115; left: 16px; position: absolute;
            top: 536px; width: 100px;" Text="Load data" />
        
        
        <div id ="griddiv" style="z-index: 119; left: 16px; width: 280px; position: absolute; top: 568px;
            height: 184px;">
        <asp:GridView ID="GridView1" runat="server" Style="z-index: 116; left: 16px; position: absolute;
            top: 16px">
            <Columns>
                <asp:CommandField ButtonType="Button" ShowSelectButton="True" />
            </Columns>
            <RowStyle Font-Size="Smaller" />
            <HeaderStyle BackColor="#C0C0FF" Font-Size="Smaller" />
        </asp:GridView>
       
   </div>
        <esri:OverviewMap ID="OverviewMap1" runat="server" ExpandPercentage="1000" Height="104px"
            Map="Map1" MapResourceManager="MapResourceManager1" OverviewMapResource="Accident"
            Style="z-index: 116; left: 560px; position: absolute; top: 352px" Width="128px" />
        <asp:Label ID="Label2" runat="server" ForeColor="#FF8080" Style="z-index: 117; left: 136px;
            position: absolute; top: 536px" Text="Please Select Featues first!"></asp:Label>
        <asp:Button ID="Button3" runat="server" Height="24px" Style="z-index: 118; left: 472px;
            position: absolute; top: 496px" Text="Clear Selection" Width="96px" />
        &nbsp;
		
    </form>
    
    
</body>
</html>
