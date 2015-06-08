Imports System.Data.OleDb
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Collections
Imports ESRI.ArcGIS.ADF.Web.DataSources
Imports ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer
Imports ESRI.ArcGIS.ADF.Web.UI.WebControls
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.ADF.ArcGISServer


Namespace Crashsafe


    Partial Class MapPage
        Inherits System.Web.UI.Page
        'Implements System.Web.UI.ICallbackEventHandler
        'Protected WithEvents Label6 As System.Web.UI.WebControls.Label

        Dim cstype As Type = Me.GetType()
        Public sADFCallBackFunctionInvocation As String

        Private Returnstring As String = ""
		Dim streetlay As LayerDescription
		Dim interlay As LayerDescription


		'global variable

		Dim strscript As String
		Dim optDistChosen As Boolean
		Dim optstateChosen As Boolean
		Dim indi As StrIndividual
		Dim caption As String
		Dim caption1 As String
		Dim queryString As String
		Dim distNo As String
		Dim mutipleYear As Boolean
		Dim YearNum As Integer
		Dim NoShow As Boolean
		Dim connString As String
		Dim showlargechar As Boolean

		Dim QueryType As Integer
		Dim SubQueryType As Integer
		Dim subQueryString As String
		Dim queryStringNum As String
		Dim optRateChosen As Boolean
		Dim queryStringRate As String
		Dim intersection As String

		Dim HowManyResult(10) As Integer
		Dim newcaption As String
		Dim gridviewtable As DataTable
		Dim optsUrbanChosen As Boolean
		Dim wholeselect As DataTable
		Dim BlackNum(800) As BlackMapshow
		Dim BlackRate(800) As BlackMapshow

		Dim exp As String

		'local variables
		Dim showmap As Integer
		Dim ShowWholemap As Boolean
		Dim IsCsect As Boolean
		Dim ShowAcc As Boolean



#Region " Web Form Designer Generated Code "

        'This call is required by the Web Form Designer.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: This method call is required by the Web Form Designer
            'Do not modify it using the code editor.
            InitializeComponent()
        End Sub

#End Region



        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            'QueryType 1.bytime 2.by CraCha by type of collision etc.3.by black spot 4.by total highway csection 
            '5.by total highway crash 6.7.bytotal highway based on selected unit 8.byinteldef.aspx by flexible query


            If Not IsPostBack Then
                Dim login As String
                login = Session("succ")
                If login = "" Then ' fake user
                    Response.Redirect("LoginIn.aspx")

                End If

                getSession()

                showmap = 0 'used to control map showing size

				

					Btn5.Visible = False


					CheckBox1.Visible = False
					Button2.Visible = True
					cmbYear.Visible = True

					indi.year = 2000
					Labelnew.Text = "For map control(Year:" & indi.year & ")"
					queryString = ""

					'mapcontrol............



				Dim layerName As String = Nothing
				loadResource(layerName)
				Dim sResource As String = ""

				Session("whole") = 1
				loadMap(exp, sResource, layerName)
				Session("layerName") = layerName
				Session("resName") = sResource

			End If
        End Sub

        Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

				Response.Redirect("options.aspx")
            
        End Sub

        Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
            If CheckBox1.Checked = True Then
                ShowAcc = True
            Else
                ShowAcc = False
            End If
        End Sub

        Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
            'If Not IsPostBack Then
            Session.Add("VALID_USER", True)
			indi.year = CInt(cmbYear.SelectedItem.Text)
			Labelnew.Text = "Map control: " & indi.year
            Dim layerName As String = Nothing
            loadResource(layerName)
            Session("layerName") = layerName
            Dim sResource As String = ""

            ' exp = "HWY_NUM = '0010' AND (MILE_POST >=0 AND MILE_POST <249.73 )  AND (HOUR >= '04' AND HOUR <= '11')"
            Showstreet.Checked = False
			 ShowInt.Checked = False

			loadMap("", sResource, layerName)


        End Sub

      


        Private Sub cmbYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbYear.SelectedIndexChanged

			'	indi.year = cmbYear.SelectedItem.Text
			'	Labelnew.Text = "For map control(Year:" & indi.year & ")"


			'Session.Add("VALID_USER", True)

		End Sub

      
        ''' <summary>
        ''' Gets the ID of the first layer that contains the string passed in for name. Layer must be queryable.
        ''' </summary>
        ''' <param name="name">Layer name. Can be part of the name. Not case sensitive.</param>
        ''' <param name="qfunc">IQueryFunctionality object.</param>
        ''' <returns>Integer ID of the matching layer, or -1 if no layer found.</returns>
        Private Function GetLayerId(ByVal name As String, ByVal qfunc As IQueryFunctionality) As String

            Dim theId As String = String.Empty
            Dim layerIds() As String = Nothing
            Dim layerNames() As String = Nothing

            qfunc.GetQueryableLayers(Nothing, layerIds, layerNames)

            Dim i As Integer
            For i = 0 To layerNames.Length - 1
                If layerNames(i).IndexOf(name, System.StringComparison.CurrentCultureIgnoreCase) > -1 Then
                    theId = layerIds(i)
                    Exit For
                End If
            Next

            Return theId

        End Function

    

        Private Sub loadResource(ByRef layerName)
            ' Dim layerName As String = ""
            Dim rdef As String = Nothing
            Dim itemNum As Integer = 4


            'very important,mapresource mannager and map control must be initialized before used

            MapResourceManager1.Initialize()
            'Dim intYear As Integer = CInt(indi.year)
           

            If indi.year = "2000" Then
					layerName = "Accidents2000"
				rdef = "(default)@ac2000"
			ElseIf indi.year = "2001" Then

				
					layerName = "Crashes_State_2001"

				rdef = "(default)@ac2001"
			ElseIf indi.year = "2002" Then

				
					layerName = "Crashes_State_2002"


				rdef = "(default)@ac2002"
			ElseIf indi.year = "2003" Then

					layerName = "Crashes_State_2003"
				rdef = "(default)@ac2003"
			ElseIf indi.year = "2004" Then

					layerName = "Crashes_State_2004"
				rdef = "(default)@ac2004"

			ElseIf indi.year = "2005" Then

					layerName = "ACC2005"
				rdef = "(default)@ac2005"

			ElseIf indi.year = "2006" Then

					layerName = "ACC2006"
				rdef = "(default)@ac2006"

			End If

			DropDownList1.Items(0).Value = layerName
			DropDownList1.Items(1).Value = "section2000"

			Dim item As MapResourceItem
			Dim def As New GISResourceItemDefinition
			Dim res As New ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapResourceLocal


			item = MapResourceManager1.ResourceItems(itemNum)
			def = item.Definition

			def.ResourceDefinition = rdef
			item.CreateResource()


        End Sub

        Private Sub loadMap(ByRef exp, ByVal sResourcename, ByVal layerName)

            Try

                Dim cstype As Type = Me.GetType()
                Dim mapSP As MapServerProxy
                Dim FeaidSet As New ESRI.ArcGIS.ADF.ArcGISServer.FIDSet
                Dim queryfilter As New ESRI.ArcGIS.ADF.Web.QueryFilter

                Dim mrl As ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapResourceLocal
                Dim qf As ESRI.ArcGIS.ADF.Web.DataSources.IQueryFunctionality

                Dim mapName As String

                Dim mapDpt As New ESRI.ArcGIS.ADF.ArcGISServer.MapDescription
                Dim color As New ESRI.ArcGIS.ADF.ArcGISServer.RgbColor
                Dim itemNum As Integer = 4

                Dim intYear As Integer = CInt(indi.year)

                If intYear = 2001 Or intYear = 2002 Or intYear = 2003 Or intYear = 2004 Then
                    exp = ChangeExp(exp, intYear)
                End If

                If Session("whole") = 1 Then
                    Session("exp") = exp
                Else
                    Session("subexp") = exp
                End If


                Map1.InitializeFunctionalities()

                Dim mf As ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapFunctionality
                mf = CType(Map1.GetFunctionality(itemNum), _
                   ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapFunctionality)

                mrl = CType(mf.MapResource, _
                   ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapResourceLocal)
                mapSP = mrl.MapServerProxy()
                qf = CType(mrl.CreateFunctionality(GetType(ESRI.ArcGIS.ADF.Web.DataSources.IQueryFunctionality), Nothing), ESRI.ArcGIS.ADF.Web.DataSources.IQueryFunctionality)
                qf.Initialize()

                queryfilter.WhereClause = exp

                Dim id As String
				Dim strid As String
				Dim intid As String
                id = GetLayerId(layerName, qf)
				strid = GetLayerId("section2000", qf)
                intid = GetLayerId("ALLIntersections", qf)
				'
                mf.DisplaySettings.ImageDescriptor.TransparentBackground = True

                mapDpt = mf.MapDescription

                Dim i As Integer
                For i = 0 To mapDpt.LayerDescriptions.Length - 1
                    Dim layer As LayerDescription = mapDpt.LayerDescriptions(i)
                    If layer.SelectionSymbol.GetType Is GetType(ESRI.ArcGIS.ADF.ArcGISServer.SimpleMarkerSymbol) Or layer.SelectionSymbol.GetType Is GetType(ESRI.ArcGIS.ADF.ArcGISServer.SimpleLineSymbol) Then
                        layer.Visible = False
                        'datatable.Columns(i).DataType Is GetType(ESRI.ArcGIS.ADF.Web.Geometry.Geometry)
                    End If
                Next

                Dim showlay As LayerDescription = mapDpt.LayerDescriptions(CInt(id))
				streetlay = mapDpt.LayerDescriptions(CInt(strid))
				interlay = mapDpt.LayerDescriptions(CInt(Intid))
				'clear the section layer's difinition expression   

If Not (QueryType = 3 Or QueryType = 4) Then
	 streetlay.DefinitionExpression = ""
	 interlay.DefinitionExpression = ""
End If

                Session("streetlay") = streetlay
				Session("interlay") = interlay
                Dim gfc_s As IEnumerable = Map1.GetFunctionalities()
                Dim gResource_s As ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource = Nothing
                Dim gfunc_s As IGISFunctionality


                For Each gfunc_s In gfc_s
                    If Not gfunc_s.Resource.Name = "Accident" Then
                        gResource_s = CType(gfunc_s.Resource, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)


                        gResource_s.Graphics.Clear()


                    End If
                Next gfunc_s


                If sResourcename = "" Then

                    showlay.DefinitionExpression = exp
                    showlay.Visible = True
                Else

                    mapName = mrl.MapServer.MapName(0)
                    Dim lids() As Object = Nothing
                    Dim flds As String() = qf.GetFields(Nothing, id)

                    Dim scoll As New ESRI.ArcGIS.ADF.StringCollection(flds)
                    queryfilter.SubFields = scoll
                    ' queryfilter.ReturnADFGeometries = True

                    Dim datatbl As System.Data.DataTable = qf.Query(Nothing, id, queryfilter)

                    Session("datatable") = datatbl
                    If sResourcename = "QuerySelection" Then
                        Session("wholeselect") = datatbl
                    End If
                    Dim drs_s As DataRowCollection = datatbl.Rows

                    Dim shpind As Integer = -1
                    Dim j As Integer
                    For j = 0 To datatbl.Columns.Count - 1
                        If datatbl.Columns(j).DataType Is GetType(ESRI.ArcGIS.ADF.Web.Geometry.Geometry) Then
                            shpind = j
                            Exit For
                        End If
                    Next j
                    ' Dim mapctrl As ESRI.ArcGIS.ADF.Web.UI.WebControls.Map = CType(Map1, ESRI.ArcGIS.ADF.Web.UI.WebControls.Map)

                   

                    For Each gfunc_s In gfc_s
                        If gfunc_s.Resource.Name = sResourcename Then
                            gResource_s = CType(gfunc_s.Resource, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)
                            Exit For
                        End If
                    Next gfunc_s

                    If gResource_s Is Nothing Then
                        Throw New Exception("Selection Graphics layer not in MapResourceManager")
                    End If
                    Dim gselectionlayer As ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer = Nothing

                    Dim dt_s As System.Data.DataTable
                    For Each dt_s In gResource_s.Graphics.Tables
                        If TypeOf dt_s Is ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer Then
                            gselectionlayer = CType(dt_s, ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer)
                            Exit For
                        End If
                    Next dt_s

                    If gselectionlayer Is Nothing Then
                        gselectionlayer = New ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer()
                        gResource_s.Graphics.Tables.Add(gselectionlayer)
                    End If

                    Dim dr_s As DataRow
                    For Each dr_s In drs_s
                        Dim geom_s As ESRI.ArcGIS.ADF.Web.Geometry.Geometry = CType(dr_s(shpind), ESRI.ArcGIS.ADF.Web.Geometry.Geometry)
                        Dim ge_s As ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement = Nothing

                        If sResourcename = "QuerySelection" Then


                            If QueryType = 3 Then
                                Dim sym As New ESRI.ArcGIS.ADF.Web.Display.Symbol.SimpleLineSymbol

                                sym.Width = 15.0
                                sym.Color = Drawing.Color.Blue


                                ge_s = New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, sym)
                            Else
                                ge_s = New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, System.Drawing.Color.Blue)


                            End If


                        Else

                            If QueryType = 3 Then
                                Dim sym As New ESRI.ArcGIS.ADF.Web.Display.Symbol.SimpleLineSymbol

                                sym.Width = 15.0
                                sym.Color = Drawing.Color.Yellow


                                ge_s = New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, sym)
                            Else
                                ge_s = New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, System.Drawing.Color.Yellow)
                            End If
                        End If
                        ge_s.Symbol.Transparency = 0.0
                        gselectionlayer.Add(ge_s)
                    Next dr_s

                    '  gResource_s.DisplaySettings.DisplayInTableOfContents = True
                    gResource_s.DisplaySettings.Visible = True


                    'OverviewMap1.Extent = mapDpt.MapArea.Extent
                    ' Map1.Extent = gselectionlayer.FullExtent()

                End If

                showlay.Visible = True


               
                Map1.Extent = Map1.GetFullExtent()

                If Map1.ImageBlendingMode = ImageBlendingMode.WebTier Then
                    Map1.Refresh()
                Else
                    If Map1.ImageBlendingMode = ImageBlendingMode.Browser Then
                        Map1.RefreshResource(gResource_s.Name)
                        ' Map1.RefreshResource(mri.Name)
                    End If
                End If

            Catch e As Exception
                Map1.Refresh()
                System.Diagnostics.Debug.WriteLine(("Exception: " + e.Message))
                Return
            End Try

            Session("TargetLayer") = DropDownList1.SelectedItem.Value
            Session("DropIndex") = DropDownList1.SelectedIndex
            ' Toc1.Refresh()
        End Sub

        Protected Sub DropDownList1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.TextChanged

            Session("TargetLayer") = DropDownList1.SelectedItem.Value
            Session("DropIndex") = DropDownList1.SelectedIndex
            If DropDownList1.SelectedItem.Value = "section2000" Then
                IsCsect = True
            Else
                IsCsect = False

            End If

            'clear Map1's QuerySelection, and Accident layer
            Dim gfc_s As IEnumerable = Map1.GetFunctionalities()
            Dim gResource_s As ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource = Nothing
            Dim gfunc_s As IGISFunctionality
            For Each gfunc_s In gfc_s
                If Not (gfunc_s.Resource.Name = "Accident" Or gfunc_s.Resource.Name = "QuerySelection") Then
                    gResource_s = CType(gfunc_s.Resource, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)
                    gResource_s.Graphics.Clear()
                End If
            Next gfunc_s
            'set griddiv unvisible
            GridView1.Visible = False


        End Sub

        Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
            Dim rowt As Object = e.Row.RowType

            Dim str As String = "Select$" & e.Row.RowIndex & " "
            Dim index As Integer = e.Row.RowIndex
            Dim s As String = " & index & "
            If (rowt = DataControlRowType.DataRow) Then

                e.Row.Attributes("onmousedown") = ClientScript.GetPostBackClientHyperlink(Me.GridView1, str)
                e.Row.Attributes.Add("onclick", "javascript:SelectRow(" & index & ")")


            End If
        End Sub

        Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged


            IsCsect = Session("IsCsect")

            GridView1.SelectedRow.BackColor = Drawing.Color.LightBlue

            Dim i As Integer
            Dim j As Integer
            Dim shpind As Integer
            Dim selint As Integer
            Dim gfc_s As IEnumerable = Map1.GetFunctionalities()
            Dim gResource_s As ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource = Nothing
            Dim gfunc_s As IGISFunctionality
            For Each gfunc_s In gfc_s
                If gfunc_s.Resource.Name = "Selection" Then
                    gResource_s = CType(gfunc_s.Resource, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)
                    Exit For
                End If
            Next gfunc_s
            Dim gtable As DataTable = gResource_s.Graphics.Tables(0)

            'set the all selected features color as green
            Dim n As Integer
            For n = 0 To gtable.Rows.Count - 1
                Dim ele As ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement = gtable.Rows.Item(n).ItemArray(2)
                If IsCsect Then
                    Dim symb As ESRI.ArcGIS.ADF.Web.Display.Symbol.SimpleLineSymbol = ele.Symbol
                    symb.Color = Drawing.Color.Green
                Else
                    Dim symb As ESRI.ArcGIS.ADF.Web.Display.Symbol.SimpleMarkerSymbol = ele.Symbol
                    symb.Color = Drawing.Color.Green
                End If
          
            Next


            For i = 0 To gtable.Columns.Count - 1
                If gtable.Columns(i).DataType Is GetType(ESRI.ArcGIS.ADF.Web.Geometry.Geometry) Then
                    shpind = i
                    Exit For
                End If
            Next i

            Dim t As Integer
            t = GridView1.SelectedIndex
            'Session("sindex") = t

            gridviewtable = Session("gridviewtable")

            For j = 0 To gridviewtable.Columns.Count - 1
                If gridviewtable.Columns(j).DataType Is GetType(ESRI.ArcGIS.ADF.Web.Geometry.Geometry) Then
                    selint = j
                    Exit For
                End If
            Next j


            Dim element As ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement = gtable.Rows.Item(t).ItemArray(2)

            If IsCsect Then
                Dim sym As ESRI.ArcGIS.ADF.Web.Display.Symbol.SimpleLineSymbol = element.Symbol
                sym.Color = Drawing.Color.Red
            Else
                Dim sym As ESRI.ArcGIS.ADF.Web.Display.Symbol.SimpleMarkerSymbol = element.Symbol
                sym.Color = Drawing.Color.Red
            End If

            Map1.Refresh()
            ' Response.Write("<script language ='javascript'>if (window.open('Information.aspx?')) != null subwin =window.open('Information.aspx','form','width=250,height=550,left=800,top=10,scrollbars = yes');</script>")

        End Sub

        

        Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

            gridviewtable = Session("gridviewtable")

            Dim copydata As DataTable = gridviewtable.Copy
            Dim newdata As New DataTable
            Dim i As Integer
            For i = 0 To gridviewtable.Columns.Count - 1
                Dim name As String = gridviewtable.Columns(i).ColumnName
                Dim col As DataColumn = gridviewtable.Columns(i)

                If Not (name = "FID" Or name = "Shape" Or name = "IS_SELECTED" Or name = "CRASH_DATE" Or name = "CSECT" Or name = "ROUTE") Then
                    copydata.Columns.Remove(name)
                End If
            Next

            GridView1.DataSource = copydata
            GridView1.DataBind()
            GridView1.Visible = True
           
        End Sub

        Private Sub ClearGraphic()
            Dim funs As IEnumerable = Map1.GetFunctionalities
            Dim gisfun As ESRI.ArcGIS.ADF.Web.DataSources.IGISFunctionality
            For Each gisfun In funs
                If Not gisfun.Resource.Name = "Accident" And Not gisfun.Resource.Name = "QuerySelection" Then
                    Dim res As ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource = CType(gisfun.Resource, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)
                    res.Graphics.Clear()
                End If
            Next
        End Sub

        Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
            ClearGraphic()
            GridView1.Visible = False
        End Sub

       
        Protected Sub Showstreet_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Showstreet.CheckedChanged

            Dim linelay As LayerDescription = Session("streetlay")

            If Showstreet.Checked Then
                linelay.Visible = True
                Map1.Refresh()
            Else
                linelay.Visible = False
                Map1.Refresh()
            End If
        End Sub

Protected Sub ShowInt_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ShowInt.CheckedChanged
			Dim interlay As LayerDescription = Session("interlay")

			If ShowInt.Checked Then
				interlay.Visible = True
				Map1.Refresh()
			Else
				interlay.Visible = False
				Map1.Refresh()
			End If
		End Sub


        Sub getSession()

            QueryType = Session("QueryType")
            SubQueryType = Session("SubQueryType")
            caption = Session("caption")
            caption1 = Session("caption1")
            queryString = Session("queryString")
            subQueryString = Session("subQueryString")
            queryStringNum = Session("queryStringNum")
            optRateChosen = Session("optRateChosen")
            queryStringRate = Session("queryStringRate")
            intersection = Session("intersection")
            HowManyResult = Session("HowManyResult")
            newcaption = Session("newcaption")

            optsUrbanChosen = Session("optsUrbanChosen")


            BlackNum = Session("BlackNum")
            BlackRate = Session("BlackRate")

            indi = Session("indi")

            optstateChosen = Session("optstateChosen")
            optDistChosen = Session("optDistChosen")
            distNo = Session("distNo")


        End Sub


			
		 
	End Class
End Namespace
