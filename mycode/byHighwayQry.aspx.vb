Imports System.Data.OleDb
Imports System

Imports System.Data
Imports System.math
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Collections
Imports ESRI.ArcGIS.ADF.Web.Geometry.envelope
Imports ESRI.ArcGIS.ADF.Web.DataSources
Imports ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer
Imports ESRI.ArcGIS.ADF.Web.UI.WebControls
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.ADF.ArcGISServer

Namespace Crashsafe


Partial Class byHighwayQry
    Inherits System.Web.UI.Page
    Public myDs As New DataSet()
        Public pagenum As Integer

        Public AppPath As String = "C:\Data"
        Dim cstype As Type = Me.GetType()
        Dim exp As String
        Public Datafile As String = "C:\Data\mdbdata\prospect.mdb"
        Public whereString As String
		Public condsql As String
	  ''global variable

		Dim strscript As String
		Dim queryString As String
		Dim connString As String
		'Dim showlargechar As Boolean
		Dim QueryType As Integer
		Dim SubQueryType As Integer

	 'local variable

	 Dim predicttable As DataTable
	  Public index As Integer

        'Public Ds As DataSet
#Region " Web Form Designer Generated Code "

        'This call is required by the Web Form Designer.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents Panel2 As System.Web.UI.WebControls.Panel


        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: This method call is required by the Web Form Designer
            'Do not modify it using the code editor.
            InitializeComponent()
        End Sub

#End Region

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

            If Not IsPostBack Then
                Dim login As String
                login = Session("succ")
                If login = "" Then ' fake user
                    Response.Redirect("LoginIn.aspx")
                End If

                QueryType = 8
                Map1.Visible = False
                Map2.Visible = False
                'add the highwaynumber to the droplist
                If TxtHW.Items.Count() <> 0 Then
                    Exit Sub
                End If
                TxtHW.Items.Clear()
                TxtHW.Items.Add("")

                'add all the highwaynum from table
                Dim accDB As String
                Dim accConn As ADODB.Connection
                Dim rst As ADODB.Recordset
                Dim tempsql As String

                accDB = "C:\DATA\mdbdata\prospect.mdb"
                accConn = New ADODB.Connection()
                With accConn
                    'Telling ADO to use JOLT Here
                    .Provider = "Microsoft.Jet.OLEDB.4.0"
                    .Open(accDB)
                End With

                tempsql = "SELECT DISTINCT[ROUTE] FROM prospect"
                rst = New ADODB.Recordset()
                'rst.Open(tempsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                rst.Open(tempsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                With rst
                    If .RecordCount < 1 Then
                        Exit Sub
                    Else
                        rst.MoveFirst()
                        Do While (Not .EOF)
                            TxtHW.Items.Add(.Fields(0).Value)
                            .MoveNext()
                        Loop
                    End If
                End With

                accConn.Close()
                rst = Nothing
            End If


        End Sub

        Private Sub TxtHW_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtHW.SelectedIndexChanged
            Image1.Visible = False
            Panel15.Visible = False
            Panel3.Visible = False
            Panel12.Visible = False
            PanelA.Visible = False
            PanelB.Visible = False
            PanelC.Visible = False
            PanelD.Visible = False
            Map1.Visible = False
            Map2.Visible = False
            GridView1.Visible = False

            BtnExplain.Visible = False
            Call inputCesectVal(TxtHW.SelectedItem.Text)
        End Sub

        Private Sub inputCesectVal(ByVal HWtext As String)
            Dim accConn As ADODB.Connection
            Dim rst As ADODB.Recordset
            Dim accDB As String
            Dim condsql As String

            cmbCsect.Items.Clear()
            cmbCsect.Items.Add("")

            accConn = New ADODB.Connection()
            'rst = New ADODB.Recordset
            accDB = "C:\DATA\mdbdata\section2000.mdb"
            With accConn
                'Telling ADO to use JOLT Here
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Open(accDB)
            End With

            condsql = "SELECT DISTINCT[CSECT] FROM section2000 WHERE HWY_NUM='" & HWtext & "'"
            rst = New ADODB.Recordset()
            rst.Open(condsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            With rst
                If .RecordCount < 1 Then
                    Exit Sub
                Else
                    rst.MoveFirst()
                    Do While (Not .EOF)
                        cmbCsect.Items.Add(.Fields(0).Value)
                        .MoveNext()
                    Loop
                End If
            End With

            accConn.Close()
            rst = Nothing
        End Sub

        Private Sub BtnPro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPro.Click
            Dim highwaynum As String
            Dim cesectnum As String
            Dim Ds As New DataSet


            On Error GoTo Proc_Err

            If TxtHW.SelectedItem.Text = "" Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('The highway number can not be empty!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            Else
                highwaynum = TxtHW.SelectedItem.Text
                Panel3.Visible = False
                Panel12.Visible = False
                PanelA.Visible = False
                PanelB.Visible = False
                PanelC.Visible = False
                PanelD.Visible = False

                '**
                Image1.Visible = False
                Panel15.Visible = False
                BtnExplain.Text = ">"
                BtnExplain.Visible = False
            End If
            cesectnum = cmbCsect.SelectedItem.Text

            '*****************
            '****search the result
            '****************
            Dim accDB As String
            Dim accConn As ADODB.Connection
            Dim rst As ADODB.Recordset

            condsql = "WHERE ROUTE='" & highwaynum & "'"
            queryString = condsql

            If cesectnum <> "" Then

                condsql = condsql & " AND CSECT='" & cesectnum & "' "
                Dim querycsect As String = cesectnum
                querycsect = querycsect.Insert(3, "-")
                queryString = queryString & " AND CSECT='" & querycsect & "' "
            End If


            condsql = "SELECT * FROM prospect " & condsql & " ORDER BY [CSECT]"

            '***************
            Dim timeAdapter As OleDbDataAdapter
            Dim connection1 As OleDbConnection
            Dim connection2 As OleDbConnection
            Dim myReader As OleDbDataReader
            Dim cmd As OleDbCommand
            'Dim Datafile As String
            ' Dim Ds As DataSet

            connString = "Provider=Microsoft.Jet.OLEDB.4.0; User Id=; Password=; Data Source= C:\Data\mdbdata\prospect.mdb"
            connection1 = New OleDbConnection(connString)

            connection1.Open()
            cmd = New OleDbCommand(condsql, connection1)
            myReader = cmd.ExecuteReader()
            pagenum = 0
            While myReader.Read
                pagenum = pagenum + 1
            End While
            If pagenum <> 0 Then
                'show the map

                Panel3.Visible = True
                ' ShowMap(highwaynum)

                connection2 = New OleDbConnection(connString)
                Ds = New DataSet()
                timeAdapter = New OleDbDataAdapter(condsql, connection2)
                'Datafile = "C:\Data\prospect.mdb"
                timeAdapter.Fill(Ds, Datafile)
                GridView1.Visible = True
                '**
                BtnExplain.Visible = True

                '**
                predicttable = Ds.Tables(Datafile)
                GridView1.DataSource = predicttable
                GridView1.DataBind()

                'show the map
                exp = queryString
                exp = Replace(exp, "[", "")
                exp = Replace(exp, "]", "")
                exp = Replace(exp, "WHERE", "")
                loadMap(exp)
                'Map1.Visible = True

            Else
                Panel3.Visible = False
                GridView1.Visible = False
                '**
                Panel15.Visible = False
                BtnExplain.Visible = False

                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Nothing has been found!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            End If
Proc_Exit:
            myReader = Nothing
            cmd = Nothing
            'Ds = Nothing
            timeAdapter = Nothing
            connection1 = Nothing
            connection2 = Nothing
            Exit Sub
Proc_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error' & Err.Description)"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            Resume Proc_Exit
        End Sub

        Private Sub GetPage()
            Dim timeAdapter As OleDbDataAdapter
            Dim connection As OleDbConnection
            Dim Datafile As String
            Dim condsql As String

            condsql = "WHERE ROUTE='" & TxtHW.SelectedItem.Text & "'"
            If cmbCsect.SelectedItem.Text <> "" Then
                condsql = condsql & "AND CSECT='" & cmbCsect.SelectedItem.Text & "'"
            End If
            condsql = "SELECT * FROM prospect " & condsql & " ORDER BY [CSECT]"

            connString = "Provider=Microsoft.Jet.OLEDB.4.0; User Id=; Password=; Data Source= C:\Data\mdbdata\prospect.mdb"
            connection = New OleDbConnection(connString)
            timeAdapter = New OleDbDataAdapter(condsql, connection)
            myDs = New DataSet()
            Datafile = "C:\Data\mdbdata\prospect.mdb"
            timeAdapter.Fill(myDs, Datafile)
            connection.Close()
            timeAdapter = Nothing
        End Sub

        Private Sub EmptyCsect()
            CSECTYEAR.Text = "" '
            CSECT.Text = ""
            LOGMI_FROM.Text = ""
            LOGMI_TO.Text = ""
            PARISH.Text = ""
            ROUTE.Text = ""
            MILEPOINT.Text = ""
            LENGTH.Text = ""
            FUN_CLASS.Text = ""
            ADT.Text = ""
            SHOUL_TYPE.Text = ""
            PAVE_TYPE1.Text = ""
            PAVE_WIDTH.Text = ""
            NUM_LANES.Text = ""
            MED_TYPE.Text = ""
            MED_WIDTH.Text = ""
            SHOU_WIDTH.Text = ""
            HWY_CLASS.Text = ""
            HWY_TYPE.Text = ""
            HWY_NUM.Text = ""
            BYPASS.Text = ""
            MIPOST_FR.Text = ""
            MIPOST_TO.Text = ""
            TOT_ACC.Text = ""
            INT_ACC.Text = ""
            NON_INT.Text = ""
            TOT_PERMVM.Text = ""
            INT_PERMV.Text = ""
            NON_PERMVM.Text = ""
            AVERSEV.Text = ""
        End Sub

        

        Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging

            'Dim Ds As DataSet
            'Dim Datafile As String
            'Dim timeAdapter As OleDbDataAdapter
            'Dim connection1 As OleDbConnection
            'Dim connection2 As OleDbConnection
            'Dim myReader As OleDbDataReader
            'Dim cmd As OleDbCommand
            ''Dim Datafile As String
            '' Dim Ds As DataSet

            'connString = "Provider=Microsoft.Jet.OLEDB.4.0; User Id=; Password=; Data Source= C:\Data\prospect.mdb"
            'connection1 = New OleDbConnection(connString)

            'connection1.Open()
            'cmd = New OleDbCommand(condsql, connection1)
            'myReader = cmd.ExecuteReader()


            'connection2 = New OleDbConnection(connString)
            'Ds = New DataSet()
            'timeAdapter = New OleDbDataAdapter(condsql, connection2)
            'Datafile = "C:\Data\prospect.mdb"
            'timeAdapter.Fill(Ds, Datafile)
            GridView1.Visible = True

            GridView1.PageIndex = e.NewPageIndex
            GridView1.DataSource = predicttable
            GridView1.DataBind()


        End Sub

        'Private Sub BtnFirst_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BtnFirst.Click
        '    Call GetPage()
        '    Call EmptyCsect()

        '    GridView1.CurrentPageIndex = 0
        '    GridView1.DataSource = myDs.Tables(0).DefaultView
        '    GridView1.DataBind()
        '    myDs = Nothing

        '    BtnFirst.Enabled = False
        '    BtnFirst.ImageUrl = ("images/lastpageCan.gif")

        '    If BtnLast.Enabled Then
        '        BtnLast.Enabled = False
        '        BtnLast.ImageUrl = ("images/lastCancel.gif")
        '    End If
        '    If BtnNext.Enabled = False Then
        '        BtnNext.Enabled = True
        '        BtnNext.ImageUrl = ("images/next.gif")
        '    End If
        '    If BtnFinal.Enabled = False Then
        '        BtnFinal.Enabled = True
        '        BtnFinal.ImageUrl = ("images/nextpage.gif")
        '    End If
        '    labPage.Text = "page:" & (GridView1.CurrentPageIndex + 1)
        '    click = True
        '    Image1.Visible = False 'chart
        '    Image4.Visible = False 'map 2
        '    Image5.Visible = False 'map 3
        'End Sub

        'Private Sub BtnLast_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BtnLast.Click
        '    Call GetPage()
        '    Call EmptyCsect()

        '    If GridView1.CurrentPageIndex >= 1 Then
        '        GridView1.CurrentPageIndex = GridView1.CurrentPageIndex - 1
        '        GridView1.DataSource = myDs.Tables(0).DefaultView
        '        GridView1.DataBind()

        '        If GridView1.CurrentPageIndex = 0 Then
        '            BtnLast.Enabled = False
        '            BtnLast.ImageUrl = ("images/lastCancel.gif")
        '            BtnFirst.Enabled = False
        '            BtnFirst.ImageUrl = ("images/lastpageCan.gif")
        '        End If

        '        myDs = Nothing
        '    Else
        '        myDs = Nothing
        '        BtnLast.Enabled = False
        '        BtnLast.ImageUrl = ("images/lastCancel.gif")
        '    End If
        '    '
        '    If BtnNext.Enabled = False Then
        '        BtnNext.Enabled = True
        '        BtnNext.ImageUrl = ("images/next.gif")
        '    End If
        '    If BtnFinal.Enabled = False Then
        '        BtnFinal.Enabled = True
        '        BtnFinal.ImageUrl = ("images/nextpage.gif")
        '    End If
        '    labPage.Text = "page:" & (GridView1.CurrentPageIndex + 1)
        '    click = True
        '    Image1.Visible = False 'chart
        '    Image4.Visible = False 'map 2
        '    Image5.Visible = False 'map 3
        'End Sub

        'Private Sub BtnNext_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BtnNext.Click
        '    Call GetPage()
        '    Call EmptyCsect()

        '    If GridView1.CurrentPageIndex <= GridView1.PageCount - 2 Then
        '        GridView1.CurrentPageIndex = GridView1.CurrentPageIndex + 1
        '        GridView1.DataSource = myDs.Tables(0).DefaultView
        '        GridView1.DataBind()

        '        If GridView1.CurrentPageIndex = GridView1.PageCount - 1 Then
        '            BtnNext.Enabled = False
        '            BtnNext.ImageUrl = ("images/lastCancel.gif")
        '            BtnFinal.Enabled = False
        '            BtnFinal.ImageUrl = ("images/lastpageCan.gif")
        '        End If

        '        myDs = Nothing

        '    Else
        '        myDs = Nothing
        '        BtnNext.Enabled = False
        '        BtnNext.ImageUrl = ("images/nextCancel.gif")
        '    End If

        '    If BtnLast.Enabled = False Then
        '        BtnLast.Enabled = True
        '        BtnLast.ImageUrl = ("images/last.gif")
        '    End If
        '    If BtnFirst.Enabled = False Then
        '        BtnFirst.Enabled = True
        '        BtnFirst.ImageUrl = ("images/lastpage.gif")
        '    End If
        '    labPage.Text = "page:" & (GridView1.CurrentPageIndex + 1)
        '    click = True
        '    Image1.Visible = False 'chart
        '    Image4.Visible = False 'map 2
        '    Image5.Visible = False 'map 3
        'End Sub

        'Private Sub BtnFinal_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BtnFinal.Click
        '    Call GetPage()
        '    Call EmptyCsect()

        '    GridView1.CurrentPageIndex = GridView1.PageCount - 1
        '    GridView1.DataSource = myDs.Tables(0).DefaultView
        '    GridView1.DataBind()

        '    myDs = Nothing

        '    BtnFinal.Enabled = False
        '    BtnFinal.ImageUrl = ("images/nextpageCan.gif")

        '    If BtnNext.Enabled Then
        '        BtnNext.Enabled = False
        '        BtnNext.ImageUrl = ("images/nextCancel.gif")
        '    End If
        '    If BtnLast.Enabled = False And GridView1.PageCount > 1 Then
        '        BtnLast.Enabled = True
        '        BtnLast.ImageUrl = ("images/last.gif")
        '    End If
        '    If BtnFirst.Enabled = False And GridView1.PageCount > 1 Then
        '        BtnFirst.Enabled = True
        '        BtnFirst.ImageUrl = ("images/lastpage.gif")
        '    End If
        '    labPage.Text = "page:" & (GridView1.CurrentPageIndex + 1)
        '    click = True
        '    Image1.Visible = False 'chart
        '    Image4.Visible = False 'map 2
        '    Image5.Visible = False 'map 3
        'End Sub

        'Private Sub ShowMap(ByVal highwaynum As String)
        '    '*********************************image1 show
        '    Dim exp As String
        '    Dim sAXLText As String
        '    Dim iWidth As Integer = 136
        '    Dim iHeight As Integer = 136

        '    Dim imageURL As String
        '    Dim sServer As String = System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPSERVER")
        '    Dim iPort As Integer = CInt(System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPPORT"))
        '    Dim sService As String = System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPSERVICE")
        '    Dim conArcIMS As New ESRI.ArcIMS.Server.ServerConnection(sServer, iPort)
        '    Dim axlResponse As New System.Xml.XmlDocument()

        '    conArcIMS.ServiceName = sService

        '    exp = "HWY_NUM='" & highwaynum & "'"

        '    sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
        '    sAXLText = sAXLText & "<REQUEST><GET_IMAGE><PROPERTIES>"
        '    sAXLText = sAXLText & "<IMAGESIZE width=""" & iWidth & """ height=""" & iHeight & """/>"
        '    sAXLText = sAXLText & "<ENVELOPE minx="" -94.065 "" miny="" 28.82 "" maxx="" -88.88 "" maxy=""33.05"" />"
        '    sAXLText = sAXLText & "<LEGEND display=""false"" />"
        '    sAXLText = sAXLText & "<LAYERLIST>"
        '    sAXLText = sAXLText & "<LAYERDEF id=""3"" visible=""false"" />"
        '    sAXLText = sAXLText & "<LAYERDEF id=""2"" visible=""true"" >"
        '    sAXLText = sAXLText & "<SPATIALQUERY where=""" & exp & """ />"
        '    sAXLText = sAXLText & "<SIMPLERENDERER >"
        '    sAXLText = sAXLText & "<SIMPLEMARKERSYMBOL  color=""0,255,0"" width=""3.5"" />"
        '    sAXLText = sAXLText & "</SIMPLERENDERER> </LAYERDEF> </LAYERLIST>"
        '    sAXLText = sAXLText & "</PROPERTIES></GET_IMAGE></REQUEST></ARCXML>"

        '    axlResponse.LoadXml(conArcIMS.Send(sAXLText))
        '    If axlResponse.GetElementsByTagName("OUTPUT").Count = 1 Then

        '        Dim strDestination, strSource As String
        '        Dim nodeOutput As System.Xml.XmlNodeList = axlResponse.GetElementsByTagName("OUTPUT")
        '        imageURL = nodeOutput(0).Attributes("url").Value

        '        imageURL = Replace(imageURL, "http://dghj5sc1/output/", "")
        '        strSource = "C:\ARCIMS\output\" & imageURL
        '        strDestination = "C:\Inetpub\wwwroot\crashSafe\mapImages\" & imageURL
        '        FileCopy(strSource, strDestination)
        '        imageURL = "mapImages/" & imageURL
        '        Image3.ImageUrl = imageURL

        '    End If
        '    axlResponse = Nothing
        'End Sub


        Private Sub GridView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

            GridView1.SelectedRow.BackColor = System.Drawing.Color.PaleGreen
            index = GridView1.SelectedIndex

            'show the selectedfeature in Map2
            'show the talbe
            'show the cesct
            Panel12.Visible = True
            PanelA.Visible = True
            PanelB.Visible = True
            PanelC.Visible = True
            PanelD.Visible = True

            'get the information
            Dim accConn As ADODB.Connection
            Dim rst As ADODB.Recordset
            Dim accDB As String
            Dim condsql As String
            Dim i, count As Integer

            accConn = New ADODB.Connection()
            accDB = "C:\Data\mdbdata\section2000.mdb"
            With accConn
                'Telling ADO to use JOLT Here
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Open(accDB)
            End With

            count = CInt(GridView1.SelectedRow.Cells(3).Text)
            condsql = "WHERE CSECT='" & GridView1.SelectedRow.Cells(2).Text & "'"
            condsql = "SELECT *  FROM section2000 " & condsql & " ORDER BY [LOGMI_FROM]"

            rst = New ADODB.Recordset()
            rst.Open(condsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            Dim mid As String
            Try
                With rst
                    If .RecordCount >= 1 Then
                        rst.MoveFirst()
                        'show the result
                        For i = 0 To count - 2
                            rst.MoveNext()
                        Next
                        CSECTYEAR.Text = .Fields(0).Value  '
                        CSECT.Text = .Fields(1).Value
                        LOGMI_FROM.Text = .Fields(2).Value
                        LOGMI_TO.Text = .Fields(3).Value

                        ' PARISH.Text = GetParish(.Fields(4).Value)
                        mid = .Fields(4).Value
                        mid = GetParish(mid)
                        PARISH.Text = mid

                        ROUTE.Text = .Fields(5).Value
                        MILEPOINT.Text = .Fields(6).Value
                        LENGTH.Text = .Fields(7).Value

                        mid = .Fields(8).Value
                        FUN_CLASS.Text = GetFunClass(mid)

                        ADT.Text = .Fields(9).Value
                        SHOUL_TYPE.Text = .Fields(10).Value '*********************no define funtion yet

                        mid = .Fields(11).Value
                        PAVE_TYPE1.Text = GetPaveType(mid)

                        PAVE_WIDTH.Text = .Fields(12).Value
                        NUM_LANES.Text = .Fields(13).Value
                        MED_TYPE.Text = .Fields(14).Value '*********************no define funtion yet

                        If .Fields(15).Value.GetType Is System.DBNull.Value.GetType Then
                            mid = ""
                        Else
                            mid = .Fields(15).Value
                        End If
                        MED_WIDTH.Text = mid

                        SHOU_WIDTH.Text = .Fields(16).Value

                        mid = .Fields(17).Value
                        HWY_CLASS.Text = GetHwyClass(mid, 0)

                        mid = .Fields(18).Value
                        HWY_TYPE.Text = GetHwType(mid, 0)

                        HWY_NUM.Text = .Fields(19).Value

                        If .Fields(20).Value.GetType Is System.DBNull.Value.GetType Then
                            mid = ""
                        Else
                            mid = .Fields(20).Value
                        End If
                        BYPASS.Text = GetBypass(mid)

                        MIPOST_FR.Text = .Fields(21).Value
                        MIPOST_TO.Text = .Fields(22).Value
                        TOT_ACC.Text = .Fields(23).Value
                        INT_ACC.Text = .Fields(24).Value
                        NON_INT.Text = .Fields(25).Value
                        TOT_PERMVM.Text = .Fields(26).Value
                        INT_PERMV.Text = .Fields(27).Value
                        NON_PERMVM.Text = .Fields(28).Value
                        AVERSEV.Text = .Fields(29).Value
                        '*****************
                    End If
                End With

                accConn.Close()
                rst = Nothing
                Dim csectStr As String = CSECT.Text
                Dim s As String = "abc"

                csectStr = csectStr.Insert(3, "-")



                'ROUTE = '0010' AND (CRASH_HOUR >= '01' AND CRASH_HOUR <= '07') AND (MILE_POST BETWEEN 0 AND 12)
                whereString = "ROUTE =  '" & ROUTE.Text & "' AND CSECT = '" & CSECT.Text & "'  AND (LOGMI_FROM = " & CDbl(LOGMI_FROM.Text) & " AND LOGMI_TO = " & CDbl(LOGMI_TO.Text) & ")"
                whereString = "ROUTE =  '" & ROUTE.Text & "' AND CSECT = '" & csectStr & "'  AND (LOGMI_FROM = " & CDbl(LOGMI_FROM.Text) & " AND LOGMI_TO = " & CDbl(LOGMI_TO.Text) & ")"

                '(CSECT='283-08' AND LOGMI_FROM=1.77 AND LOGMI_TO=0.7) OR(CSECT='450-10' AND LOGMI_FROM=2.88 AND LOGMI_TO=3.64)
                'ROUTE =  '0010' AND CSECT = '03203'  AND (LOGMI_FROM = 4.01 AND LOGMI_TO = 5.72)
                getselfeature(whereString)

                Dim sectionNum As String
                sectionNum = GridView1.SelectedRow.Cells(2).Text
                sectionNum = sectionNum.Insert(3, "-")
                ' ShowCsectMap(sectionNum) 'show the map of cescet
            Catch
                Exit Sub
            End Try




            'draw the chart and show in the page
            Dim numbervlues(3) As String
            'GridView1.SelectedRow.Cells(2).Text
            numbervlues(1) = GridView1.SelectedRow.Cells(10).Text & vbTab & GridView1.SelectedRow.Cells(7).Text & vbTab & GridView1.SelectedRow.Cells(4).Text
            numbervlues(2) = GridView1.SelectedRow.Cells(11).Text & vbTab & GridView1.SelectedRow.Cells(8).Text & vbTab & GridView1.SelectedRow.Cells(5).Text
            numbervlues(3) = GridView1.SelectedRow.Cells(12).Text & vbTab & GridView1.SelectedRow.Cells(9).Text & vbTab & GridView1.SelectedRow.Cells(6).Text
            DrawChart(numbervlues)

        End Sub

        Private Sub DrawChart(ByRef NumValues() As String)
            Dim objCSpace1 As Microsoft.Office.Interop.Owc11.ChartSpace = New Microsoft.Office.Interop.Owc11.ChartSpaceClass()
            Dim objChart1
            Dim chart_title As String
            Dim strValues As String
            Dim i As Integer
            Dim DataLiteral(3) As String
            DataLiteral(1) = "Ob"
            DataLiteral(2) = "Nrs"
            DataLiteral(3) = "Ep"
            strValues = "2001" & vbTab & "2002" & vbTab & "2003"

            objChart1 = objCSpace1.Charts.Add(0)
            objChart1.Type = Microsoft.Office.Interop.Owc11.ChartChartTypeEnum.chChartTypeColumnClustered

            For i = 0 To 2
                'there are values of three years(2001 to 2003)
                objChart1.SeriesCollection.Add(i)
            Next i

            With objChart1
                .HasLegend = True
                For i = 0 To 2
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimSeriesNames, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, DataLiteral(i + 1))
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimValues, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, NumValues(i + 1)) 'y
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimCategories, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, strValues) 'x
                Next i
                .SeriesCollection(0).Name = "=""Comparison"""

                .HasTitle = True

                chart_title = "Comparison"
                .Title.Caption = chart_title

                .Axes(0).HasTitle = True
                .Axes(1).HasTitle = True

                .Axes(0).Title.Caption = "Year"

                .Axes(1).Title.Caption = "# of Crashes"
            End With

            'Now a chart is ready to export to a GIF.
			Dim ChartName As String = Rnd() & ".gif"  'times & Rnd() & ".gif"

            Dim strAbsolutePath As String = Server.MapPath(".") & "\" & ChartName

            Dim strRelativePath As String = "./" & ChartName

            objCSpace1.ExportPicture(strAbsolutePath, "GIF", 496, 220)
            Image1.ImageUrl = strRelativePath
            Image1.Visible = True
Proc_Exit:
            objChart1 = Nothing
            Exit Sub
Proc_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error Number and Description' & Err.Num & Err.Description)"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            objChart1 = Nothing
        End Sub

        'show the map of cescet
        'Private Sub ShowCsectMap(ByVal CscetNum As String)
        '    'for there is not fine way to search corresponding position now, try to search in fowlling way:
        '    'for there is not fine way to search corresponding position now, try to search in fowlling way:search whole map
        '    '-94~-93.1~-92.1~-91.2~-90.2~89.3
        '    '29.04~29.84~30.64~31.44~32.24~33.01
        '    Dim sAXLText As String
        '    Dim exp, midsec As String
        '    Dim i As Integer
        '    Dim minXX, minYY, maxXX, maxYY As Double
        '    Dim imageURL As String = ""

        '    Dim sServer As String = System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPSERVER")
        '    Dim iPort As Integer = CInt(System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPPORT"))
        '    Dim sService As String = System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPSERVICE")
        '    Dim conArcIMS As New ESRI.ArcIMS.Server.ServerConnection(sServer, iPort)
        '    Dim axlRequest As New ESRI.ArcIMS.Server.Xml.AxlRequests()
        '    Dim axlResponse As New System.Xml.XmlDocument()
        '    Dim response As String
        '    Dim CsectNodelist As System.Xml.XmlNodeList

        '    conArcIMS.ServiceName = sService

        '    exp = "CSECT='" & CscetNum & "'"

        '    '
        '    For i = 1 To 9
        '        If i = 1 Then
        '            minXX = -93.1
        '            minYY = 31.44
        '            maxXX = -92.1
        '            maxYY = 32.24
        '        ElseIf i = 2 Then
        '            minXX = -92.1
        '            minYY = 31.44
        '            maxXX = -91.2
        '            maxYY = 32.24
        '        ElseIf i = 3 Then
        '            minXX = -91.2
        '            minYY = 31.44
        '            maxXX = -90.2
        '            maxYY = 32.24
        '        ElseIf i = 4 Then
        '            minXX = -93.1
        '            minYY = 30.64
        '            maxXX = -92.1
        '            maxYY = 31.44
        '        ElseIf i = 5 Then
        '            minXX = -92.1
        '            minYY = 30.64
        '            maxXX = -91.2
        '            maxYY = 31.44
        '        ElseIf i = 6 Then
        '            minXX = -91.2
        '            minYY = 30.64
        '            maxXX = -90.2
        '            maxYY = 31.44
        '        ElseIf i = 7 Then
        '            minXX = -93.1
        '            minYY = 29.84
        '            maxXX = -92.1
        '            maxYY = 30.64
        '        ElseIf i = 8 Then
        '            minXX = -92.1
        '            minYY = 29.84
        '            maxXX = -91.2
        '            maxYY = 30.64
        '        ElseIf i = 9 Then
        '            minXX = -91.2
        '            minYY = 29.84
        '            maxXX = -90.2
        '            maxYY = 30.64
        '        End If
        '        sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
        '        sAXLText = sAXLText & "<REQUEST>"
        '        sAXLText = sAXLText & "<GET_FEATURES outputmode=""newxml"" geometry=""false"">"
        '        sAXLText = sAXLText & "<LAYER id=""2"" /> " 'search for information of csects
        '        sAXLText = sAXLText & "<SPATIALQUERY where=""" & exp & """ subfields=""#ALL#"">"
        '        sAXLText = sAXLText & "<SPATIALFILTER relation=""area_intersection"">"
        '        sAXLText = sAXLText & "<ENVELOPE minx=""" & minXX & """ miny=""" & minYY & """ maxx=""" & maxXX & """ maxy=""" & maxYY & """ />"
        '        sAXLText = sAXLText & "</SPATIALFILTER>"
        '        sAXLText = sAXLText & "</SPATIALQUERY>"
        '        sAXLText = sAXLText & "</GET_FEATURES></REQUEST></ARCXML>"
        '        response = conArcIMS.Send(sAXLText, "Query")
        '        axlResponse.LoadXml(response)
        '        'Save the value of cesect
        '        CsectNodelist = axlResponse.GetElementsByTagName("FIELD")

        '        Dim j, num As Integer
        '        num = CsectNodelist.Count / 34 - 1
        '        For j = 0 To num
        '            midsec = CsectNodelist(2 + 34 * j).Attributes("value").Value
        '            If midsec = CscetNum Then
        '                CsectNodelist = Nothing
        '                Dim k As Integer
        '                For k = 1 To 9
        '                    If k = 1 Then
        '                        maxXX = minXX + 0.4
        '                        maxYY = minYY + 0.3
        '                    ElseIf k = 2 Or k = 3 Or k = 5 Or k = 6 Or k = 8 Or k = 9 Then
        '                        minXX = maxXX
        '                        maxXX = maxXX + 0.4
        '                    ElseIf k = 4 Or k = 7 Then
        '                        minXX = maxXX - 1.2
        '                        maxXX = maxXX - 0.8
        '                        minYY = maxYY
        '                        maxYY = maxYY + 0.3
        '                    End If
        '                    sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
        '                    sAXLText = sAXLText & "<REQUEST>"
        '                    sAXLText = sAXLText & "<GET_FEATURES outputmode=""newxml"" geometry=""false"">"
        '                    sAXLText = sAXLText & "<LAYER id=""2"" /> " 'search for information of csects
        '                    sAXLText = sAXLText & "<SPATIALQUERY where=""" & exp & """ subfields=""#ALL#"">"
        '                    sAXLText = sAXLText & "<SPATIALFILTER relation=""area_intersection"">"
        '                    sAXLText = sAXLText & "<ENVELOPE minx=""" & minXX & """ miny=""" & minYY & """ maxx=""" & maxXX & """ maxy=""" & maxYY & """ />"
        '                    sAXLText = sAXLText & "</SPATIALFILTER>"
        '                    sAXLText = sAXLText & "</SPATIALQUERY>"
        '                    sAXLText = sAXLText & "</GET_FEATURES></REQUEST></ARCXML>"
        '                    response = conArcIMS.Send(sAXLText, "Query")
        '                    axlResponse.LoadXml(response)
        '                    'Save the value of cesect
        '                    CsectNodelist = axlResponse.GetElementsByTagName("FIELD")
        '                    Dim l, num2 As Integer
        '                    num2 = CsectNodelist.Count / 34 - 1
        '                    For l = 0 To num2
        '                        midsec = CsectNodelist(2 + 34 * j).Attributes("value").Value
        '                        If midsec = CscetNum Then
        '                            Dim exp2 As String
        '                            Dim iWidth As Integer = 150
        '                            Dim iHeight As Integer = 150
        '                            Dim sService2 As String = System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPSERVICE2")
        '                            conArcIMS.ServiceName = sService2

        '                            'layer 0:DistrictBoundary
        '                            'layer 1:URBANIZED
        '                            'layer 2:section2000
        '                            'layer 3:section2000
        '                            exp2 = "HWY_NUM='" & TxtHW.SelectedRow.Text & "'"
        '                            sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
        '                            sAXLText = sAXLText & "<REQUEST><GET_IMAGE><PROPERTIES>"
        '                            sAXLText = sAXLText & "<IMAGESIZE width=""" & iWidth & """ height=""" & iHeight & """/>"
        '                            sAXLText = sAXLText & "<ENVELOPE minx=""" & minXX & """ miny=""" & minYY & """ maxx=""" & maxXX & """ maxy=""" & maxYY & """ />"
        '                            sAXLText = sAXLText & "<LEGEND display=""false"" />"
        '                            sAXLText = sAXLText & "<LAYERLIST>"
        '                            sAXLText = sAXLText & "<LAYERDEF id=""0"" visible=""false"" />"
        '                            sAXLText = sAXLText & "<LAYERDEF id=""3"" visible=""true"" >"
        '                            sAXLText = sAXLText & "<SPATIALQUERY where=""" & exp2 & """ />"
        '                            sAXLText = sAXLText & "<SIMPLERENDERER >"
        '                            sAXLText = sAXLText & "<SIMPLELINESYMBOL  type=""solid"" color=""255,0,0"" width=""4.5"" />"
        '                            sAXLText = sAXLText & "</SIMPLERENDERER> </LAYERDEF> </LAYERLIST>"
        '                            sAXLText = sAXLText & "</PROPERTIES></GET_IMAGE></REQUEST></ARCXML>"

        '                            axlResponse.LoadXml(conArcIMS.Send(sAXLText))
        '                            If axlResponse.GetElementsByTagName("OUTPUT").Count = 1 Then
        '                                Dim strDestination, strSource As String
        '                                Dim nodeOutput As System.Xml.XmlNodeList = axlResponse.GetElementsByTagName("OUTPUT")

        '                                imageURL = nodeOutput(0).Attributes("url").Value
        '                                imageURL = Replace(imageURL, "http://dghj5sc1/output/", "")
        '                                strSource = "C:\ARCIMS\output\" & imageURL
        '                                strDestination = "C:\Inetpub\wwwroot\crashSafe\mapImages\" & imageURL
        '                                FileCopy(strSource, strDestination)
        '                                Image4.ImageUrl = "mapImages/" & imageURL
        '                                Image4.Visible = True
        '                            End If
        '                            '
        '                            sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
        '                            sAXLText = sAXLText & "<REQUEST><GET_IMAGE><PROPERTIES>"
        '                            sAXLText = sAXLText & "<IMAGESIZE width=""" & iWidth & """ height=""" & iHeight & """/>"
        '                            sAXLText = sAXLText & "<ENVELOPE minx=""" & minXX & """ miny=""" & minYY & """ maxx=""" & maxXX & """ maxy=""" & maxYY & """ />"
        '                            sAXLText = sAXLText & "<LEGEND display=""false"" />"
        '                            sAXLText = sAXLText & "<LAYERLIST>"
        '                            sAXLText = sAXLText & "<LAYERDEF id=""0"" visible=""false"" />"
        '                            sAXLText = sAXLText & "<LAYERDEF id=""3"" visible=""true"" >"
        '                            sAXLText = sAXLText & "<SPATIALQUERY where=""" & exp & """ />"
        '                            sAXLText = sAXLText & "<SIMPLERENDERER >"
        '                            sAXLText = sAXLText & "<SIMPLELINESYMBOL  type=""solid"" color=""255,0,0"" width=""4.5"" />"
        '                            sAXLText = sAXLText & "</SIMPLERENDERER> </LAYERDEF> </LAYERLIST>"
        '                            sAXLText = sAXLText & "</PROPERTIES></GET_IMAGE></REQUEST></ARCXML>"

        '                            axlResponse.LoadXml(conArcIMS.Send(sAXLText))
        '                            If axlResponse.GetElementsByTagName("OUTPUT").Count = 1 Then
        '                                Dim strDestination, strSource As String
        '                                Dim nodeOutput As System.Xml.XmlNodeList = axlResponse.GetElementsByTagName("OUTPUT")

        '                                imageURL = nodeOutput(0).Attributes("url").Value
        '                                imageURL = Replace(imageURL, "http://dghj5sc1/output/", "")
        '                                strSource = "C:\ARCIMS\output\" & imageURL
        '                                strDestination = "C:\Inetpub\wwwroot\crashSafe\mapImages\" & imageURL
        '                                FileCopy(strSource, strDestination)
        '                                Image5.ImageUrl = "mapImages/" & imageURL
        '                                Image5.Visible = True
        '                            End If
        '                        End If
        '                    Next l
        '                Next k
        '                Exit For
        '            End If
        '        Next j
        '    Next i
        '    If imageURL = "" Then '***********Find other zones
        '        For i = 1 To 13 'outer squres
        '            If i = 1 Then
        '                minXX = -94
        '                minYY = 29.7 '*******
        '                maxXX = -93.1
        '                maxYY = 30.64
        '            ElseIf i = 2 Then
        '                minXX = -94
        '                minYY = 30.64
        '                maxXX = -93.1
        '                maxYY = 31.44
        '            ElseIf i = 3 Then
        '                minXX = -94
        '                minYY = 31.44
        '                maxXX = -93.1
        '                maxYY = 32.24
        '            ElseIf i = 4 Then
        '                minXX = -94
        '                minYY = 32.24
        '                maxXX = -93.1
        '                maxYY = 33.01
        '            ElseIf i = 5 Then
        '                minXX = -93.1
        '                minYY = 29.5 '*******
        '                maxXX = -92.1
        '                maxYY = 29.84
        '            ElseIf i = 6 Then
        '                minXX = -92.1
        '                minYY = 29.04
        '                maxXX = -91.2
        '                maxYY = 29.84
        '            ElseIf i = 7 Then
        '                minXX = -91.2
        '                minYY = 29.04
        '                maxXX = -90.2
        '                maxYY = 29.84
        '            ElseIf i = 8 Then
        '                minXX = -90.2
        '                minYY = 29.04
        '                maxXX = -89.27
        '                maxYY = 29.84
        '            ElseIf i = 9 Then
        '                minXX = -90.2
        '                minYY = 29.84
        '                maxXX = -89.27
        '                maxYY = 30.64
        '            ElseIf i = 10 Then
        '                minXX = -90.2
        '                minYY = 30.64
        '                maxXX = -89.27
        '                maxYY = 31.44
        '            ElseIf i = 11 Then
        '                minXX = -93.1
        '                minYY = 32.24
        '                maxXX = -92.1
        '                maxYY = 33.01
        '            ElseIf i = 12 Then
        '                minXX = -92.1
        '                minYY = 32.24
        '                maxXX = -91.2
        '                maxYY = 33.01
        '            ElseIf i = 13 Then
        '                minXX = -91.2
        '                minYY = 32.24
        '                maxXX = -90.2
        '                maxYY = 33.01
        '            End If
        '            sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
        '            sAXLText = sAXLText & "<REQUEST>"
        '            sAXLText = sAXLText & "<GET_FEATURES outputmode=""newxml"" geometry=""false"">"
        '            sAXLText = sAXLText & "<LAYER id=""2"" /> " 'search for information of csects
        '            sAXLText = sAXLText & "<SPATIALQUERY where=""" & exp & """ subfields=""#ALL#"">"
        '            sAXLText = sAXLText & "<SPATIALFILTER relation=""area_intersection"">"
        '            sAXLText = sAXLText & "<ENVELOPE minx=""" & minXX & """ miny=""" & minYY & """ maxx=""" & maxXX & """ maxy=""" & maxYY & """ />"
        '            sAXLText = sAXLText & "</SPATIALFILTER>"
        '            sAXLText = sAXLText & "</SPATIALQUERY>"
        '            sAXLText = sAXLText & "</GET_FEATURES></REQUEST></ARCXML>"
        '            response = conArcIMS.Send(sAXLText, "Query")
        '            axlResponse.LoadXml(response)
        '            'Save the value of cesect
        '            CsectNodelist = axlResponse.GetElementsByTagName("FIELD")

        '            Dim j, num As Integer
        '            num = CsectNodelist.Count / 34 - 1
        '            For j = 0 To num
        '                midsec = CsectNodelist(2 + 34 * j).Attributes("value").Value
        '                If midsec = CscetNum Then
        '                    CsectNodelist = Nothing
        '                    Dim k As Integer
        '                    For k = 1 To 9
        '                        If k = 1 Then
        '                            maxXX = minXX + 0.4
        '                            maxYY = minYY + 0.3
        '                        ElseIf k = 2 Or k = 3 Or k = 5 Or k = 6 Or k = 8 Or k = 9 Then
        '                            minXX = maxXX
        '                            maxXX = maxXX + 0.4
        '                        ElseIf k = 4 Or k = 7 Then
        '                            minXX = maxXX - 1.2
        '                            maxXX = maxXX - 0.8
        '                            minYY = maxYY
        '                            maxYY = maxYY + 0.3
        '                        End If
        '                        sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
        '                        sAXLText = sAXLText & "<REQUEST>"
        '                        sAXLText = sAXLText & "<GET_FEATURES outputmode=""newxml"" geometry=""false"">"
        '                        sAXLText = sAXLText & "<LAYER id=""2"" /> " 'search for information of csects
        '                        sAXLText = sAXLText & "<SPATIALQUERY where=""" & exp & """ subfields=""#ALL#"">"
        '                        sAXLText = sAXLText & "<SPATIALFILTER relation=""area_intersection"">"
        '                        sAXLText = sAXLText & "<ENVELOPE minx=""" & minXX & """ miny=""" & minYY & """ maxx=""" & maxXX & """ maxy=""" & maxYY & """ />"
        '                        sAXLText = sAXLText & "</SPATIALFILTER>"
        '                        sAXLText = sAXLText & "</SPATIALQUERY>"
        '                        sAXLText = sAXLText & "</GET_FEATURES></REQUEST></ARCXML>"
        '                        response = conArcIMS.Send(sAXLText, "Query")
        '                        axlResponse.LoadXml(response)
        '                        'Save the value of cesect
        '                        CsectNodelist = axlResponse.GetElementsByTagName("FIELD")
        '                        Dim l, num2 As Integer
        '                        num2 = CsectNodelist.Count / 34 - 1
        '                        For l = 0 To num2
        '                            midsec = CsectNodelist(2 + 34 * j).Attributes("value").Value
        '                            If midsec = CscetNum Then
        '                                Dim iWidth As Integer = 150
        '                                Dim iHeight As Integer = 150
        '                                Dim sService2 As String = System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPSERVICE2")
        '                                Dim exp2 As String
        '                                conArcIMS.ServiceName = sService2

        '                                'layer 0:DistrictBoundary
        '                                'layer 1:URBANIZED
        '                                'layer 2:section2000
        '                                'layer 3:section2000
        '                                'layer 0:DistrictBoundary
        '                                'layer 1:URBANIZED
        '                                'layer 2:section2000
        '                                'layer 3:section2000
        '                                exp2 = "HWY_NUM='" & TxtHW.SelectedRow.Text & "'"
        '                                sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
        '                                sAXLText = sAXLText & "<REQUEST><GET_IMAGE><PROPERTIES>"
        '                                sAXLText = sAXLText & "<IMAGESIZE width=""" & iWidth & """ height=""" & iHeight & """/>"
        '                                sAXLText = sAXLText & "<ENVELOPE minx=""" & minXX & """ miny=""" & minYY & """ maxx=""" & maxXX & """ maxy=""" & maxYY & """ />"
        '                                sAXLText = sAXLText & "<LEGEND display=""false"" />"
        '                                sAXLText = sAXLText & "<LAYERLIST>"
        '                                sAXLText = sAXLText & "<LAYERDEF id=""0"" visible=""false"" />"
        '                                sAXLText = sAXLText & "<LAYERDEF id=""3"" visible=""true"" >"
        '                                sAXLText = sAXLText & "<SPATIALQUERY where=""" & exp2 & """ />"
        '                                sAXLText = sAXLText & "<SIMPLERENDERER >"
        '                                sAXLText = sAXLText & "<SIMPLELINESYMBOL  type=""solid"" color=""255,0,0"" width=""4.5"" />"
        '                                sAXLText = sAXLText & "</SIMPLERENDERER> </LAYERDEF> </LAYERLIST>"
        '                                sAXLText = sAXLText & "</PROPERTIES></GET_IMAGE></REQUEST></ARCXML>"

        '                                axlResponse.LoadXml(conArcIMS.Send(sAXLText))
        '                                If axlResponse.GetElementsByTagName("OUTPUT").Count = 1 Then
        '                                    Dim strDestination, strSource As String
        '                                    Dim nodeOutput As System.Xml.XmlNodeList = axlResponse.GetElementsByTagName("OUTPUT")

        '                                    imageURL = nodeOutput(0).Attributes("url").Value
        '                                    imageURL = Replace(imageURL, "http://dghj5sc1/output/", "")
        '                                    strSource = "C:\ARCIMS\output\" & imageURL
        '                                    strDestination = "C:\Inetpub\wwwroot\crashSafe\mapImages\" & imageURL
        '                                    FileCopy(strSource, strDestination)
        '                                    Image4.ImageUrl = "mapImages/" & imageURL
        '                                    Image4.Visible = True
        '                                End If

        '                                sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
        '                                sAXLText = sAXLText & "<REQUEST><GET_IMAGE><PROPERTIES>"
        '                                sAXLText = sAXLText & "<IMAGESIZE width=""" & iWidth & """ height=""" & iHeight & """/>"
        '                                sAXLText = sAXLText & "<ENVELOPE minx=""" & minXX & """ miny=""" & minYY & """ maxx=""" & maxXX & """ maxy=""" & maxYY & """ />"
        '                                sAXLText = sAXLText & "<LEGEND display=""false"" />"
        '                                sAXLText = sAXLText & "<LAYERLIST>"
        '                                sAXLText = sAXLText & "<LAYERDEF id=""0"" visible=""false"" />"
        '                                sAXLText = sAXLText & "<LAYERDEF id=""3"" visible=""true"" >"
        '                                sAXLText = sAXLText & "<SPATIALQUERY where=""" & exp & """ />"
        '                                sAXLText = sAXLText & "<SIMPLERENDERER >"
        '                                sAXLText = sAXLText & "<SIMPLELINESYMBOL  type=""solid"" color=""255,0,0"" width=""4.5"" />"
        '                                sAXLText = sAXLText & "</SIMPLERENDERER> </LAYERDEF> </LAYERLIST>"
        '                                sAXLText = sAXLText & "</PROPERTIES></GET_IMAGE></REQUEST></ARCXML>"

        '                                axlResponse.LoadXml(conArcIMS.Send(sAXLText))
        '                                If axlResponse.GetElementsByTagName("OUTPUT").Count = 1 Then
        '                                    Dim strDestination, strSource As String
        '                                    Dim nodeOutput As System.Xml.XmlNodeList = axlResponse.GetElementsByTagName("OUTPUT")

        '                                    imageURL = nodeOutput(0).Attributes("url").Value
        '                                    imageURL = Replace(imageURL, "http://dghj5sc1/output/", "")
        '                                    strSource = "C:\ARCIMS\output\" & imageURL
        '                                    strDestination = "C:\Inetpub\wwwroot\crashSafe\mapImages\" & imageURL
        '                                    FileCopy(strSource, strDestination)
        '                                    Image5.ImageUrl = "mapImages/" & imageURL
        '                                    Image5.Visible = True
        '                                End If
        '                            End If
        '                        Next l
        '                    Next k
        '                    Exit For
        '                End If
        '            Next j
        '        Next i
        '    End If
        'End Sub

        Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
            Dim resources As IEnumerable = MapResourceManager1.GetResources
            Dim res As IGISResource = Nothing
            Dim gres As ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource = Nothing

            For Each res In resources
                Dim resname As String = res.Name
                If resname = "Buffer" Or resname = "Selection" Or resname = "QuerySelection" Then
                    gres = CType(res, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)
                    gres.Graphics.Clear()
                 End If
            Next

            Response.Redirect("options.aspx")
        End Sub

        Private Sub BtnExplain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExplain.Click
            Dim a As String
            a = BtnExplain.Text
            If a = ">" Then
                Panel15.Visible = True
                BtnExplain.Text = "<"
            End If
            If a = "<" Then
                Panel15.Visible = False
                BtnExplain.Text = ">"
            End If
        End Sub

        Private Sub loadMap(ByVal exp)
            Dim cstype As Type = Me.GetType()
            Dim mri As MapResourceItem
            Dim mridefault As MapResourceItem
            Dim mapSP As MapServerProxy
            Dim FeaidSet As New ESRI.ArcGIS.ADF.ArcGISServer.FIDSet
            Dim queryfilter As New ESRI.ArcGIS.ADF.Web.QueryFilter
            ' Dim spatialfilter As New ESRI.ArcGIS.ADF.Web.SpatialFilter()
            Dim mrl As ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapResourceLocal
            Dim qf As ESRI.ArcGIS.ADF.Web.DataSources.IQueryFunctionality

            Dim mapName As String
            Dim layerName As String = ""
            Dim mapDpt As New ESRI.ArcGIS.ADF.ArcGISServer.MapDescription
            Dim color As New ESRI.ArcGIS.ADF.ArcGISServer.RgbColor
            Dim itemNum As Integer = 0
            Dim rdef As String = Nothing
            'very important,mapresource mannager and map control must be initialized before used
            MapResourceManager1.Initialize()
            layerName = "section2000"

            Dim item As MapResourceItem
            Dim def As New GISResourceItemDefinition
            Dim res As New ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapResourceLocal

            itemNum = 3
            item = MapResourceManager1.ResourceItems(itemNum)

            mri = item
            Dim n As Integer
            For n = 1 To 2
                Dim ctrlname As String = ""
                ctrlname = "Map" & n
                Dim mapctrl As ESRI.ArcGIS.ADF.Web.UI.WebControls.Map = CType(Page.FindControl(ctrlname), ESRI.ArcGIS.ADF.Web.UI.WebControls.Map)


                mapctrl.InitializeFunctionalities()
                mapctrl.InitializeFunctionality(mri)

                Dim mf As ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapFunctionality
                mf = CType(mapctrl.GetFunctionality(itemNum), _
                   ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapFunctionality)

                mrl = CType(mf.MapResource, _
                   ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapResourceLocal)
                mapSP = mrl.MapServerProxy()
                qf = CType(mrl.CreateFunctionality(GetType(ESRI.ArcGIS.ADF.Web.DataSources.IQueryFunctionality), Nothing), ESRI.ArcGIS.ADF.Web.DataSources.IQueryFunctionality)
                qf.Initialize()

                queryfilter.WhereClause = exp

                'queryfilter.WhereClause = "CRASH_HOUR >= 4 AND <= 23"

                Dim id As String
                id = GetLayerId(layerName, qf)
                mf.DisplaySettings.ImageDescriptor.TransparentBackground = True

                mapDpt = mf.MapDescription

                Dim i As Integer
                For i = 0 To mapDpt.LayerDescriptions.Length - 1
                    Dim layer As LayerDescription = mapDpt.LayerDescriptions(i)
                    If layer.SelectionSymbol.GetType Is GetType(ESRI.ArcGIS.ADF.ArcGISServer.SimpleMarkerSymbol) Then
                        layer.Visible = False
                        'datatable.Columns(i).DataType Is GetType(ESRI.ArcGIS.ADF.Web.Geometry.Geometry)
                    End If
                Next

                mapName = mrl.MapServer.MapName(0)
                Dim lids() As Object = Nothing
                Dim flds As String() = qf.GetFields(Nothing, id)

                Dim scoll As New ESRI.ArcGIS.ADF.StringCollection(flds)
                queryfilter.SubFields = scoll

                Dim datatable As System.Data.DataTable = qf.Query(Nothing, id, queryfilter)

                Dim drs_s As DataRowCollection = datatable.Rows

                Dim shpind As Integer = -1
                Dim j As Integer
                For j = 0 To datatable.Columns.Count - 1
                    If datatable.Columns(j).DataType Is GetType(ESRI.ArcGIS.ADF.Web.Geometry.Geometry) Then
                        shpind = j
                        Exit For
                    End If
                Next j
                ' Dim mapctrl As ESRI.ArcGIS.ADF.Web.UI.WebControls.Map = CType(Map1, ESRI.ArcGIS.ADF.Web.UI.WebControls.Map)

                Dim gfc_s As IEnumerable = mapctrl.GetFunctionalities()
                'Dim gfc_s As IEnumerable = mapctrl.GetFunctionalities()
                Dim gResource_s As ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource = Nothing
                Dim gfunc_s As IGISFunctionality
                For Each gfunc_s In gfc_s
                    If Not gfunc_s.Resource.Name = "Accident" Then
                        gResource_s = CType(gfunc_s.Resource, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)
                        gResource_s.Graphics.Clear()

                    End If
                Next gfunc_s

                For Each gfunc_s In gfc_s
                    If gfunc_s.Resource.Name = "QuerySelection" Then
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

                    Dim ge_s As New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, System.Drawing.Color.Red)

                    ge_s.Symbol.Transparency = 30.0
                    gselectionlayer.Add(ge_s)
                Next dr_s

                gResource_s.DisplaySettings.DisplayInTableOfContents = True
                gResource_s.DisplaySettings.Visible = True


                'Map1.Extent = gselectionlayer.FullExtent()
                If n = 1 Then
                    mapctrl.Extent = Map1.GetFullExtent()

                Else : mapctrl.Extent = gselectionlayer.FullExtent()
                End If
                mapctrl.Visible = True

                If mapctrl.ImageBlendingMode = ImageBlendingMode.WebTier Then
                    mapctrl.Refresh()
                Else
                    If mapctrl.ImageBlendingMode = ImageBlendingMode.Browser Then
                        mapctrl.RefreshResource(gResource_s.Name)
                        mapctrl.RefreshResource(mri.Name)
                    End If
                End If
            Next


        End Sub

        Private Function ChangeExp(ByVal OriExp As String, ByVal year As String) As String
            Dim NewExp As String
            If year = 2004 Then
                NewExp = Replace(OriExp, "HOUR", "HOUR_")
            Else 'for 2001 to 2003
                NewExp = Replace(OriExp, "ACC_DATE", "CRASH_DATE")
                NewExp = Replace(NewExp, "PARISH", "PARISH_CD")
                NewExp = Replace(NewExp, "CITY", "CITY_CD")
                NewExp = Replace(NewExp, "POPULATION", "POP_CD")
                NewExp = Replace(NewExp, "WEEKDAY", "DAY_OF_WK")
                NewExp = Replace(NewExp, "HOUR", "CRASH_HOUR")
                NewExp = Replace(NewExp, "HWY_NUM", "ROUTE")
                NewExp = Replace(NewExp, "HWY_TYPE", "HWY_TYPE_C")
                NewExp = Replace(NewExp, "TYPE_VEH1", "VEH_TYPE_C")
                NewExp = Replace(NewExp, "TYPE_VEH2", "VEH_TYPE_1")
                NewExp = Replace(NewExp, "VEH1_SPEED", "EST_SPEED1")
                NewExp = Replace(NewExp, "VEH2_SPEED", "EST_SPEED2")
                NewExp = Replace(NewExp, "POST_SPEED", "POSTED_SPE")
                NewExp = Replace(NewExp, "NUM_KILLED", "NUM_TOT_KI")
                NewExp = Replace(NewExp, "NUM_INJURE", "NUM_TOT_IN")
                NewExp = Replace(NewExp, "ALIGNMENT", "ALIGNMENT_")
                NewExp = Replace(NewExp, "SURF_COND", "SURF_COND_")
                NewExp = Replace(NewExp, "SURF_TYPE", "SURF_TYPE_")
                NewExp = Replace(NewExp, "TYPE_ROAD", "ROAD_TYPE_")
                NewExp = Replace(NewExp, "ROAD_COND", "ROAD_COND_")
                NewExp = Replace(NewExp, "LIGHTING", "LIGHTING_C")
                NewExp = Replace(NewExp, "WEATHER", "WEATHER_CD")
                NewExp = Replace(NewExp, "KIND_LOC", "LOC_TYPE_C")
                NewExp = Replace(NewExp, "TRAF_CTRL", "TRAFF_CNTL")
                NewExp = Replace(NewExp, "COND_VEH1", "VEH_COND_C")
                NewExp = Replace(NewExp, "COND_VEH2", "VEH_COND_1")
                NewExp = Replace(NewExp, "COND_DRIV1", "DR_COND_CD")
                NewExp = Replace(NewExp, "COND_DRIV2", "DR_COND__1")
                NewExp = Replace(NewExp, "COND_PED", "PED_COND")
                NewExp = Replace(NewExp, "VIOLATION1", "VIOLATIONS")
                NewExp = Replace(NewExp, "VIOLATION2", "VIOLATIO_1")
                NewExp = Replace(NewExp, "DIRECTION1", "TRAVEL_D_1")
                NewExp = Replace(NewExp, "DIRECTION2", "TRAVEL_D_2")
                NewExp = Replace(NewExp, "INTER", "INTERSECTI")
                NewExp = Replace(NewExp, "COMPUTER", "CRASH_NUM")
                NewExp = Replace(NewExp, "SEVERITY", "SEVERITY_C")
                NewExp = Replace(NewExp, "NUM_INJ4", "NUM_INJB")
                NewExp = Replace(NewExp, "NUM_INJ5", "NUM_INJC")
                NewExp = Replace(NewExp, "NUM_INJ6", "NUM_INJD")
                NewExp = Replace(NewExp, "NUM_INJ7", "NUM_INJE")
                NewExp = Replace(NewExp, "NUM_INJ8", "CRASH")
            End If
            Return NewExp
        End Function

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

        Private Sub getselfeature(ByVal where As String)
            Dim mri As MapResourceItem
            Dim mapSP As MapServerProxy
            Dim FeaidSet As New ESRI.ArcGIS.ADF.ArcGISServer.FIDSet
            Dim queryfilter As New ESRI.ArcGIS.ADF.Web.QueryFilter
            ' Dim spatialfilter As New ESRI.ArcGIS.ADF.Web.SpatialFilter()
            Dim mrl As ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapResourceLocal
            Dim qf As ESRI.ArcGIS.ADF.Web.DataSources.IQueryFunctionality

            Dim mapName As String
            Dim layerName As String = ""
            Dim mapDpt As New ESRI.ArcGIS.ADF.ArcGISServer.MapDescription
            Dim color As New ESRI.ArcGIS.ADF.ArcGISServer.RgbColor
            Dim itemNum As Integer = 0
            Dim rdef As String = Nothing
            'very important,mapresource mannager and map control must be initialized before used
            MapResourceManager1.Initialize()
            layerName = "section2000"

            Dim item As MapResourceItem
            Dim def As New GISResourceItemDefinition
            Dim res As New ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapResourceLocal

            itemNum = 3
            item = MapResourceManager1.ResourceItems(itemNum)

            mri = item
            Dim n As Integer = 2

            Dim ctrlname As String = ""
            ctrlname = "Map" & n
            Dim mapctrl As ESRI.ArcGIS.ADF.Web.UI.WebControls.Map = CType(Page.FindControl(ctrlname), ESRI.ArcGIS.ADF.Web.UI.WebControls.Map)

            mapctrl.InitializeFunctionalities()
            mapctrl.InitializeFunctionality(mri)

            Dim mf As ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapFunctionality
            mf = CType(mapctrl.GetFunctionality(itemNum), _
               ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapFunctionality)

            mrl = CType(mf.MapResource, _
               ESRI.ArcGIS.ADF.Web.DataSources.ArcGISServer.MapResourceLocal)
            mapSP = mrl.MapServerProxy()
            qf = CType(mrl.CreateFunctionality(GetType(ESRI.ArcGIS.ADF.Web.DataSources.IQueryFunctionality), Nothing), ESRI.ArcGIS.ADF.Web.DataSources.IQueryFunctionality)
            qf.Initialize()
            'CSECT.Text


            queryfilter.WhereClause = where

            Dim id As String
            id = GetLayerId(layerName, qf)
            mf.DisplaySettings.ImageDescriptor.TransparentBackground = True

            mapDpt = mf.MapDescription

            Dim i As Integer
            For i = 0 To mapDpt.LayerDescriptions.Length - 1
                Dim layer As LayerDescription = mapDpt.LayerDescriptions(i)
                If layer.SelectionSymbol.GetType Is GetType(ESRI.ArcGIS.ADF.ArcGISServer.SimpleMarkerSymbol) Then
                    layer.Visible = False
                    'datatable.Columns(i).DataType Is GetType(ESRI.ArcGIS.ADF.Web.Geometry.Geometry)
                End If
            Next

            mapName = mrl.MapServer.MapName(0)
            Dim lids() As Object = Nothing
            Dim flds As String() = qf.GetFields(Nothing, id)

            Dim scoll As New ESRI.ArcGIS.ADF.StringCollection(flds)
            queryfilter.SubFields = scoll

            Dim datatable As System.Data.DataTable = qf.Query(Nothing, id, queryfilter)

            Dim drs_s As DataRowCollection = datatable.Rows

            Dim shpind As Integer = -1
            Dim j As Integer
            For j = 0 To datatable.Columns.Count - 1
                If datatable.Columns(j).DataType Is GetType(ESRI.ArcGIS.ADF.Web.Geometry.Geometry) Then
                    shpind = j
                    Exit For
                End If
            Next j
            ' Dim mapctrl As ESRI.ArcGIS.ADF.Web.UI.WebControls.Map = CType(Map1, ESRI.ArcGIS.ADF.Web.UI.WebControls.Map)

            Dim gfc_s As IEnumerable = mapctrl.GetFunctionalities()
            Dim gResource_s As ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource = Nothing
            Dim gfunc_s As IGISFunctionality
            For Each gfunc_s In gfc_s
                If gfunc_s.Resource.Name = "Selection" Or gfunc_s.Resource.Name = "Buffer" Then
                    gResource_s = CType(gfunc_s.Resource, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)
                    gResource_s.Graphics.Clear()

                End If
            Next gfunc_s

            For Each gfunc_s In gfc_s
                If gfunc_s.Resource.Name = "Selection" Then
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
                Dim ge_s As New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, System.Drawing.Color.Green)
                ge_s.Symbol.Transparency = 30.0
                gselectionlayer.Add(ge_s)
            Next dr_s

            gResource_s.DisplaySettings.DisplayInTableOfContents = True
            gResource_s.DisplaySettings.Visible = True

            'Map1.Extent = gselectionlayer.FullExtent()
            Dim areaextent As New ESRI.ArcGIS.ADF.Web.Geometry.Envelope
            Dim area As ESRI.ArcGIS.ADF.Web.Geometry.Envelope = gselectionlayer.FullExtent()
            areaextent.XMin = area.XMin - (area.Width)
            areaextent.XMax = area.XMax + (area.Width)
            areaextent.YMin = area.YMin - (area.Height)
            areaextent.YMax = area.YMax + (area.Height)
            mapctrl.Extent = areaextent

            Dim gResource2 As ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource = Nothing

            For Each gfunc_s In gfc_s
                If gfunc_s.Resource.Name = "Buffer" Then
                    gResource2 = CType(gfunc_s.Resource, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)
                    Exit For
                End If
            Next gfunc_s

            If gResource2 Is Nothing Then
                Throw New Exception("Selection Graphics layer not in MapResourceManager")
            End If
            Dim gbufferlayer As ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer = Nothing

            Dim dt As System.Data.DataTable
            For Each dt In gResource2.Graphics.Tables
                If TypeOf dt Is ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer Then
                    gbufferlayer = CType(dt, ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer)
                    Exit For
                End If
            Next dt

            If gbufferlayer Is Nothing Then
                gbufferlayer = New ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer()
                gResource2.Graphics.Tables.Add(gbufferlayer)
            End If

            Dim bufferextent As New ESRI.ArcGIS.ADF.Web.Geometry.Envelope
            Dim centerp As New ESRI.ArcGIS.ADF.Web.Geometry.Point
            centerp = GetCenterPoint(CType(areaextent, ESRI.ArcGIS.ADF.Web.Geometry.Geometry))
            'Dim maxi As Double = Max(areaextent.Width, areaextent.Height)
            bufferextent.XMin = centerp.X - 0.14
            bufferextent.XMax = centerp.X + 0.14
            bufferextent.YMin = centerp.Y - 0.14
            bufferextent.YMax = centerp.Y + 0.14

            Dim geo As ESRI.ArcGIS.ADF.Web.Geometry.Geometry = CType(bufferextent, ESRI.ArcGIS.ADF.Web.Geometry.Geometry)
            Dim gebuffer As New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geo, System.Drawing.Color.Yellow)
            gebuffer.Symbol.Transparency = 90.0
            gselectionlayer.Add(gebuffer)



            If mapctrl.ImageBlendingMode = ImageBlendingMode.WebTier Then
                mapctrl.Refresh()
            Else
                If mapctrl.ImageBlendingMode = ImageBlendingMode.Browser Then
                    mapctrl.RefreshResource(gResource_s.Name)
                    mapctrl.RefreshResource(mri.Name)
                End If
            End If

            Map1.Refresh()


        End Sub

        Protected Sub cmbCsect_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCsect.SelectedIndexChanged
            Map1.Visible = False
            Map2.Visible = False
            GridView1.Visible = False
        End Sub

        Protected Sub cmbCsect_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCsect.TextChanged
            Map1.Visible = False
            Map2.Visible = False
            GridView1.Visible = False
        End Sub
    End Class

End Namespace
