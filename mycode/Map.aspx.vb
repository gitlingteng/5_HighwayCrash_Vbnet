Imports System.Data.OleDb
Imports System

Imports System.timers
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
Imports ESRI.ArcGIS.ADF.ArcGISServer.PointN


Namespace Crashsafe


    Partial Class MapPage
        Inherits System.Web.UI.Page
        'Implements System.Web.UI.ICallbackEventHandler
        'Protected WithEvents Label6 As System.Web.UI.WebControls.Label

        Dim cstype As Type = Me.GetType()
        Public sADFCallBackFunctionInvocation As String

        Private Returnstring As String = ""
        Dim streetlay, highwaylay As LayerDescription
        

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
                'Dim login As String
                'login = Session("succ")
                'If login = "" Then ' fake user
                '    Response.Redirect("LoginIn.aspx")

                'End If

                getSession()

                showmap = 0 'used to control map showing size
                If Label3.Text = "For map control" Then
                    Labelnew.Text = Label3.Text
                    'Exit Sub
                End If

                Dim TYPE As String
                TYPE = Request("TYPE")
                If TYPE = "mapcontrol" Then
                    QueryType = 9 'map control
                End If

                If DrOp.Items.Count > 0 Then
                    Labelnew.Text = Label3.Text
                    If QueryType = 3 Or QueryType = 7 Then 'black and sliding window
                        Label4.Text = caption
                    End If
                    If QueryType = 5 Then 'by partial highway with compound query
                        Label4.Text = caption & " (Top 10 crash positions be shown)"
                    End If
                    If DrOp.SelectedItem.Text <> "Whole" Then
                        ShowWholemap = True
                    Else
                        ShowWholemap = False
                    End If
                    ' Exit Sub
                End If

                ShowWholemap = False

                ' if the map show Csect
                If QueryType = 3 Or QueryType = 4 Then
                    IsCsect = True
                Else
                    IsCsect = False
                End If

                Session("IsCsect") = IsCsect

                If QueryType = 8 Then
                    DrOp.Visible = False
                    Btn5.Visible = False
                    DropYear.Visible = False
                End If

                If QueryType <> 9 Then 'IF IT IS NOT THE map control
                    Call InitF()
                Else
                    DrOp.Visible = False
                    Btn5.Visible = False
                    DropYear.Visible = False

                    CheckBox1.Visible = True
                    Button2.Visible = True
                    cmbYear.Visible = True

                    indi.year = 2000
                    Labelnew.Text = "For map control(Year:" & indi.year & ")"
                    queryString = ""
                    'Exit Sub
                    'mapcontrol............
                End If
                If QueryType = 3 Then 'blackspot     
                    If optRateChosen Then 'if choose the times crash rate
                        queryString = queryStringNum
                    Else 'if choose the top number of crashes
                        DropYear.Visible = True 'different to others  
                        If DropYear.SelectedItem.Text = "Rank by number of crashes" Then
                            queryString = queryStringNum
                        Else
                            queryString = queryStringRate
                        End If
                    End If
                End If


               
                exp = queryString
                exp = Replace(exp, "[", "")
                exp = Replace(exp, "]", "")
                exp = Replace(exp, "WHERE", "")

                Dim layerName As String = Nothing
                loadResource(layerName)
                Dim sResource As String = ""

                ' exp = "HWY_NUM = '0010' AND (MILE_POST >=0 AND MILE_POST <249.73 )  AND (HOUR >= '04' AND HOUR <= '11')"
                Session("whole") = 1
                loadMap(exp, sResource, layerName)
                Session("layerName") = layerName
                Session("resName") = sResource

            End If
        End Sub

        Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
            If QueryType = 9 Then 'mapcnotrl
                Response.Redirect("options.aspx")
            Else
                'mri.DisplaySettings.DisplayInTableOfContents = False
                'mri.DisplaySettings.Visible = False

                Response.Write("<script language ='javascript'>window.close('Map.aspx?');</script>")
            End If
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
            Dim layerName As String = Nothing
            loadResource(layerName)
            Session("layerName") = layerName
            Dim sResource As String = ""

            ' exp = "HWY_NUM = '0010' AND (MILE_POST >=0 AND MILE_POST <249.73 )  AND (HOUR >= '04' AND HOUR <= '11')"
            loadMap("", sResource, layerName)


        End Sub

        Private Sub InitF()
            Dim j As Integer

            CheckBox1.Visible = False
            Button2.Visible = False
            Button4.Text = "Close"

            If mutipleYear = True Then
                If indi.year = 1999 Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('Sorry about that we do not have the crash map of 1999!')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    YearNum = YearNum - 1
                    indi.year = indi.year + 1
                End If
                cmbYear.Visible = True
                cmbYear.Items.Clear()
                For j = 0 To YearNum
                    cmbYear.Items.Add(indi.year + j)
                Next j
            Else
                cmbYear.Visible = False
            End If

            DropYear.Visible = False

            If QueryType = 1 Then 'by time
                DrOp.Visible = True
                Btn5.Visible = True
                If SubQueryType = 1 Then 'by month   
                    Labelnew.Text = caption & ".Search by month of the year."
                    Label3.Text = Labelnew.Text
                    DrOp.Items.Add("Whole")
                    DrOp.Items.Add("January")
                    DrOp.Items.Add("February")
                    DrOp.Items.Add("March")
                    DrOp.Items.Add("April")
                    DrOp.Items.Add("May")
                    DrOp.Items.Add("June")
                    DrOp.Items.Add("July")
                    DrOp.Items.Add("August")
                    DrOp.Items.Add("September")
                    DrOp.Items.Add("October")
                    DrOp.Items.Add("November")
                    DrOp.Items.Add("December")
                ElseIf SubQueryType = 2 Then 'by day             
                    Labelnew.Text = caption & ".Search by day of the week."
                    Label3.Text = Labelnew.Text
                    DrOp.Items.Add("Whole")
                    DrOp.Items.Add("Mon")
                    DrOp.Items.Add("Tue")
                    DrOp.Items.Add("Wed")
                    DrOp.Items.Add("Thu")
                    DrOp.Items.Add("Fri")
                    DrOp.Items.Add("Sat")
                    DrOp.Items.Add("Sun")
                ElseIf SubQueryType = 3 Then 'by hour                
                    Labelnew.Text = caption & ".Search by hour of the day."
                    Label3.Text = Labelnew.Text
                    DrOp.Items.Add("Whole")
                    DrOp.Items.Add("00")
                    DrOp.Items.Add("01")
                    DrOp.Items.Add("02")
                    DrOp.Items.Add("03")
                    DrOp.Items.Add("04")
                    DrOp.Items.Add("05")
                    DrOp.Items.Add("06")
                    DrOp.Items.Add("07")
                    DrOp.Items.Add("08")
                    DrOp.Items.Add("09")
                    DrOp.Items.Add("10")
                    DrOp.Items.Add("11")
                    DrOp.Items.Add("12")
                    DrOp.Items.Add("13")
                    DrOp.Items.Add("14")
                    DrOp.Items.Add("15")
                    DrOp.Items.Add("16")
                    DrOp.Items.Add("17")
                    DrOp.Items.Add("18")
                    DrOp.Items.Add("19")
                    DrOp.Items.Add("20")
                    DrOp.Items.Add("21")
                    DrOp.Items.Add("22")
                    DrOp.Items.Add("23")
                End If
            ElseIf QueryType = 2 Then  'by CraCha
                DrOp.Visible = True
                Btn5.Visible = True
                If SubQueryType = 1 Then 'by Type of Coll   
                    Labelnew.Text = caption & ".Search by type of Coll."
                    Label3.Text = Labelnew.Text
                    DrOp.Items.Add("Whole")
                    DrOp.Items.Add("A.Non-Collision")
                    DrOp.Items.Add("B.Rear End")
                    DrOp.Items.Add("C.Head-on")
                    DrOp.Items.Add("D.Right Angle")
                    DrOp.Items.Add("E.Left Turn1")
                    DrOp.Items.Add("F.Left Turn2")
                    DrOp.Items.Add("G.Left Turn3")
                    DrOp.Items.Add("H.Right Turn1")
                    DrOp.Items.Add("I.Right Turn2")
                    DrOp.Items.Add("J.SideSwipe(Same dir.)")
                    DrOp.Items.Add("K.SideSwipe(Opposite dir.)")
                    DrOp.Items.Add("L.Other")
                ElseIf SubQueryType = 2 Then 'by Type of Acc             
                    Labelnew.Text = caption & ".Search by type of accidents."
                    Label3.Text = Labelnew.Text
                    DrOp.Items.Add("Whole")
                    DrOp.Items.Add("A.Running off Roadway")
                    DrOp.Items.Add("B.Overturning on Roadway")
                    DrOp.Items.Add("C.Collision with Pedestrian")
                    DrOp.Items.Add("D.Collision With Other motor vehicle in traffic")
                    DrOp.Items.Add("E.Collision with Parked Vehicle")
                    DrOp.Items.Add("F.Collision with Train")
                    DrOp.Items.Add("G.Collision with Bicyclist")
                    DrOp.Items.Add("H.Collison with Animal")
                    DrOp.Items.Add("I.Collison with Fixed Object")
                    DrOp.Items.Add("J.Collison with Other Object")
                    DrOp.Items.Add("K.Other non-collison on Road")
                ElseIf SubQueryType = 3 Then 'by Point of Impact               
                    Labelnew.Text = caption & ".Search by point of impact."
                    Label3.Text = Labelnew.Text
                    DrOp.Items.Add("Whole")
                    DrOp.Items.Add("A.On roadway")
                    DrOp.Items.Add("B.Shoulder")
                    DrOp.Items.Add("C.Median")
                    DrOp.Items.Add("D.Beyond shoulder - Left")
                    DrOp.Items.Add("E.Beyond shoulder - Right")
                    DrOp.Items.Add("F.Off Roadway")
                    DrOp.Items.Add("G.Gore")
                    DrOp.Items.Add("H.Unknown")
                    DrOp.Items.Add("I.Other")
                ElseIf SubQueryType = 4 Then 'by Violation of 1st Vehicle   
                    Labelnew.Text = caption & ".Search by violation of 1st vehicle."
                    Label3.Text = Labelnew.Text
                    DrOp.Items.Add("Whole")
                    DrOp.Items.Add("A.Exceeding stated speed limit")
                    DrOp.Items.Add("B.Exceeding safe speed limit")
                    DrOp.Items.Add("C.Failure to yield")
                    DrOp.Items.Add("D.Following too close")
                    DrOp.Items.Add("E.Driving left of center")
                    DrOp.Items.Add("F.Cutting in, improper pass")
                    DrOp.Items.Add("G.Failure to signal")
                    DrOp.Items.Add("H.Made wide right turn")
                    DrOp.Items.Add("I.Cut corner on left turn")
                    DrOp.Items.Add("J.Turned from wrong lane")
                    DrOp.Items.Add("K.Other improper turning")
                    DrOp.Items.Add("L.Disregarded traffic control")
                    DrOp.Items.Add("M.Improper starting")
                    DrOp.Items.Add("N.Improper parking")
                    DrOp.Items.Add("O.Failure to set out flags")
                    DrOp.Items.Add("P.Failed to dim headlights")
                    DrOp.Items.Add("Q.Vehicle condition")
                    DrOp.Items.Add("R.Driver condition")
                    DrOp.Items.Add("S.Careless operation")
                    DrOp.Items.Add("T.Unknown Violations")
                    DrOp.Items.Add("U.No violation")
                    DrOp.Items.Add("V.Other")
                End If
            ElseIf QueryType = 3 Then  'by blackspot
                DropDownList1.Visible = False
                DrOp.Visible = True
                Btn5.Visible = True
                Labelnew.Text = caption1
                Label4.Text = caption
                Label3.Text = Labelnew.Text
                DrOp.Items.Add("Whole")
                Dim number As String
                Dim i As Integer
                If optRateChosen Then 'if choose the times crash rate
                    For i = 1 To HowManyResult(0)
                        number = CStr(i)
                        If intersection = "0" Then 'if segment
                            DrOp.Items.Add("Top " & number & " of number of crashes in sections ")
                        ElseIf intersection = "1" Then  'if intersections
                            DrOp.Items.Add("Top " & number & " of number of crashes in intersections ")
                        Else 'Total
                            DrOp.Items.Add("Top " & number & " of number of crashes in intersections and sections")
                        End If
                    Next i
                Else 'if choose the top number of crashes
                    DropYear.Visible = True
                    DropYear.Items.Add("Rank by number of crashes")
                    DropYear.Items.Add("Rank by crash rate")

                    For i = 1 To HowManyResult(0)
                        number = CStr(i)
                        If intersection = "0" Then 'if segment
                            DrOp.Items.Add("Top " & number & " of number of crashes in sections ")
                        ElseIf intersection = "1" Then  'if intersections
                            DrOp.Items.Add("Top " & number & " of number of crashes in intersections ")
                        Else 'Total
                            DrOp.Items.Add("Top " & number & " of number of crashes in intersections and sections")
                        End If
                    Next i
                End If

            ElseIf QueryType = 4 Then  'by partial highway
                DropDownList1.Visible = False
                DrOp.Visible = True
                Btn5.Visible = True
                Labelnew.Text = newcaption
                Label3.Text = Labelnew.Text
                DrOp.Items.Add("Whole")
                If HowManyResult(0) >= 20 Then
                    HowManyResult(0) = 20
                End If
                Dim number As String
                Dim i As Integer
                For i = 1 To HowManyResult(0)
                    number = CStr(i)
                    DrOp.Items.Add("Top " & number & " by crash rate")
                Next i
                For i = 1 To HowManyResult(0)
                    number = CStr(i)
                    DrOp.Items.Add("Top " & number & " by Number of Crashes")
                Next i

            ElseIf QueryType = 5 Then  'by partial highway with compound query
                DrOp.Visible = True
                Btn5.Visible = True
                Labelnew.Text = caption1
                Label4.Text = caption & " (Top 10 crash positions be shown)"
                Label3.Text = Labelnew.Text
                DrOp.Items.Add("Whole")

                Dim number As String
                Dim i As Integer

                For i = 1 To HowManyResult(0)
                    number = CStr(i)
                    DrOp.Items.Add("Top " & number & " by Number of Crashes")
                Next i
            ElseIf QueryType = 6 Then  'by individual
                DrOp.Visible = False
                Btn5.Visible = False
                Labelnew.Text = caption
                Label3.Text = Labelnew.Text
            ElseIf QueryType = 7 Then  'by sliding window
                DrOp.Visible = True
                Btn5.Visible = True
                Labelnew.Text = "By sliding window query. " & caption1
                Label4.Text = caption
                Label3.Text = Labelnew.Text
                DrOp.Items.Add("Whole")
                If HowManyResult(0) >= 20 Then
                    HowManyResult(0) = 20
                End If
                Dim number As String
                Dim i As Integer
                For i = 1 To HowManyResult(0)
                    number = CStr(i)
                    DrOp.Items.Add("Top " & number & " by crash rate")
                Next i
                For i = 1 To HowManyResult(0)
                    number = CStr(i)
                    DrOp.Items.Add("Top " & number & " by Number of Crashes")
                Next i
		ElseIf QueryType = 0 Then
			 Labelnew.Text = "By urban area query. " & caption
		End If


		End Sub

        Private Sub Btn5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn5.Click

            Try

                Dim i As Integer
                Dim subexp As String

                GridView1.Visible = False

                showmap = 1
                If DrOp.SelectedItem.Text <> "Whole" Then
                    ShowWholemap = True
                    Session("whole") = 0
                Else
                    ShowWholemap = False
                    Session("datatable") = Session("wholeselect")
                    Session("whole") = 1
                    ' Session("exp") = exp
                    ClearGraphic()

                    Exit Sub
                End If

                getSession()

                subQueryString = queryString

                If QueryType = 1 Then 'by time            
                    If SubQueryType = 1 Then 'by month'******************mistack can not found in here'**************
                        Dim accDate, endDate, j, l As Date
                        Dim newtime As String

                        accDate = DateSerial(CInt(indi.year), DrOp.SelectedIndex, 1)
                        endDate = DateAdd("m", 1, accDate)
                        'newtime = Format(accDate, "yyyy-M-d") & " 0:00:00" '
                        'newtime = "2000-1-28 0:00:00"
                        'j = Format(accDate, "yyyy-M-d") '
                        'l = Format(accDate, "R") 'R, rFormats the date and time as Greenwich Mean Time (GMT)
                        'subQueryString = subQueryString & " AND [ACC_DATE] = '" & newtime & "'"
                        subQueryString = subQueryString & " AND [ACC_DATE] >= date '" & Format(accDate, "Short Date") & "' AND [ACC_DATE] < date '" & Format(endDate, "Short Date") & "'"

                    ElseIf SubQueryType = 2 Then 'by day of week

                        indi.WeekDay = CStr(DrOp.SelectedIndex)
                        Dim weeksql As String = generateFieldSql("WEEKDAY", indi.WeekDay)
                        If indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003 Then
                            Dim wkDay As String = ""
                            wkDay = changWeekday(indi.WeekDay)
                            weeksql = generateFieldSql("WEEKDAY", wkDay)
                        End If

                        subQueryString = subQueryString & weeksql


                    ElseIf SubQueryType = 3 Then 'by hour of the day   
                        If indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003 Then
                            subQueryString &= " AND [HOUR] = " & CInt(DrOp.SelectedItem.Text)
                        Else : subQueryString &= " AND [HOUR] = '" & DrOp.SelectedItem.Text & "'"
                        End If
                        ' subQueryString = subQueryString & " AND [HOUR] =" & "'" & DrOp.SelectedItem.Text & "'"
                    End If
                End If
                If QueryType = 2 Then  'by CraCha          
                    If SubQueryType = 1 Then 'by Type of Coll   
                        subQueryString = subQueryString & generateFieldSql("TYPE_COLL", Left(DrOp.SelectedItem.Text, 1))
                    ElseIf SubQueryType = 2 Then 'by Type of Acc           
                        subQueryString = subQueryString & generateFieldSql("TYPE_ACC", Left(DrOp.SelectedItem.Text, 1))
                    ElseIf SubQueryType = 3 Then 'by Point of Impact             
                        subQueryString = subQueryString & generateFieldSql("PT_IMPACT", Left(DrOp.SelectedItem.Text, 1))
                    ElseIf SubQueryType = 4 Then 'by Violation of 1st Vehicle   
                        subQueryString = subQueryString & generateFieldSql("VIOLATION1", Left(DrOp.SelectedItem.Text, 1))
                    End If
                End If
                If QueryType = 3 Then  'by blackspot  
                    i = DrOp.SelectedIndex - 1
                    'subQueryString = subQueryString
                    If optRateChosen Then 'if choose the times crash rate

							subQueryString = "WHERE CSECT='" & BlackNum(i).Csect & "' AND LOGMI_FROM>" & BlackNum(i).accLogmifrom - 0.5 & " AND LOGMI_TO<" & BlackNum(i).accLogmito + 0.5
							  'subQueryString = "WHERE CSECT='" & BlackNum(i).Csect & "' AND LOGMI_FROM=" & BlackNum(i).accLogmifrom & " AND LOGMI_TO=" & BlackNum(i).accLogmito

					Else 'if choose the top number of crashes               
						If DropYear.SelectedItem.Text = "Rank by crash rate" Then

								subQueryString = "WHERE CSECT='" & BlackRate(i).Csect & "' AND LOGMI_FROM>" & BlackRate(i).accLogmifrom - 0.5 & " AND LOGMI_TO<" & BlackRate(i).accLogmito + 0.5

						Else 'Rank by number of crashes

								subQueryString = "WHERE CSECT='" & BlackNum(i).Csect & "' AND LOGMI_FROM>" & BlackNum(i).accLogmifrom - 0.5 & " AND LOGMI_TO<" & BlackNum(i).accLogmito + 0.5

						End If
					End If
				End If
                If QueryType = 4 Then  'by parhighway  
                    Dim j As Integer
                    j = (DrOp.Items.Count - 1) / 2
                    i = DrOp.SelectedIndex - 1
                    If i < j Then 'rank by rate

                        subQueryString = "WHERE CSECT='" & BlackRate(i).Csect & "'"

                    Else 'rank by number
                        i = i - j
                        'subQueryString = "WHERE CSECT='" & BlackNum(i).Csect.Insert(3, "-") & "' AND MIPOST_FR=" & BlackNum(i).accLogmifrom & " AND MIPOST_TO=" & BlackNum(i).accLogmito
                        subQueryString = "WHERE CSECT='" & BlackNum(i).Csect & "'"
                    End If
                End If
                If QueryType = 5 Then  'by total highway based on crash  
                    i = DrOp.SelectedIndex - 1
                    subQueryString = BlackNum(i).accComputer
                    If indi.year = "2001" Or indi.year = "2002" Or indi.year = "2003" Then

                        subQueryString = Replace(subQueryString, "'", "")
                    End If
                End If
                If QueryType = 7 Then  'by sliding window
                    Dim j As Integer
                    j = (DrOp.Items.Count - 1) / 2
                    i = DrOp.SelectedIndex - 1

                    If i < j Then 'rank by rate
                        subQueryString = "HWY_NUM ='" & indi.WayNum & "' AND MILE_POST >= " & BlackRate(i).accLogmifrom & " AND MILE_POST < " & BlackRate(i).accLogmito
                    Else 'rank by number
                        i = i - j
                        subQueryString = "HWY_NUM = '" & indi.WayNum & "' AND MILE_POST >=" & BlackNum(i).accLogmifrom & " AND MILE_POST < " & BlackNum(i).accLogmito
                    End If
                End If

                subexp = subQueryString
                subexp = Replace(subexp, "[", "")
                subexp = Replace(subexp, "]", "")
                subexp = Replace(subexp, "WHERE", "")
                Dim layer As String = Session("layerName")
                Dim sResourcename As String = "PartialSelect"
               
                '    WHERE (CSECT='827-32' AND LOGMI_FROM=0 AND LOGMI_TO=0.7) OR(CSECT='290-01' AND LOGMI_FROM=10.03 AND LOGMI_TO=10.26) OR(CSECT='817-20' AND LOGMI_FROM=0.32 AND LOGMI_TO=0.27) OR(CSECT='429-02' AND LOGMI_FROM=7.26 AND LOGMI_TO=7.48) OR (CSECT='004-30' AND LOGMI_FROM=0.36 AND LOGMI_TO=0.43)
                loadMap(subexp, sResourcename, layer)

            Catch b As Exception
                Map1.Refresh()
                System.Diagnostics.Debug.WriteLine(("Exception: " + b.Message))
                Return
            End Try


        End Sub

        Private Sub cmbYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbYear.SelectedIndexChanged
            If QueryType = 9 Then 'IF IT IS  map control            
                indi.year = cmbYear.SelectedItem.Text
                Labelnew.Text = "For map control(Year:" & indi.year & ")"
            Else
                indi.year = cmbYear.SelectedItem.Text
                'If cmbYear.SelectedItem.Text <> "2001" And cmbYear.SelectedItem.Text <> "2000" Then
                '    strscript = "<script language='javascript'>"
                '    strscript = strscript & "alert('Sorry, we only have shapes for 2000 and 2001 now.' )"
                '    strscript = strscript & "</script>"
                '  ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                'End If
            End If

            Session.Add("VALID_USER", True)
            'hvMapPage.Value = "MakeMap.aspx"
        End Sub

        Private Sub DropYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DropYear.SelectedIndexChanged


			 getSession()

            GridView1.Visible = False
            If QueryType = 3 Then  'by blackspot
                showmap = 1
                ShowWholemap = False
                DrOp.Items.Clear()
                DrOp.Items.Add("Whole")
                If DropYear.SelectedItem.Text = "Rank by number of crashes" Then
                    Dim number As String
                    Dim i As Integer
                    For i = 1 To HowManyResult(0)
                        number = CStr(i)
                        If intersection = "0" Then 'if segment
                            DrOp.Items.Add("Top " & number & " of number of crashes in sections ")
                        ElseIf intersection = "1" Then  'if intersections
                            DrOp.Items.Add("Top " & number & " of number of crashes in intersections ")
                        Else 'Total
                            DrOp.Items.Add("Top " & number & " of number of crashes in intersections and sections")
                        End If
                    Next i
                    queryString = queryStringNum
                Else
                    Dim number As String
                    Dim i As Integer
                    For i = 1 To HowManyResult(0)
                        number = CStr(i)
                        If intersection = "0" Then 'if segment
                            DrOp.Items.Add("Top " & number & " of crash rate in sections ")
                        ElseIf intersection = "1" Then  'if intersections
                            DrOp.Items.Add("Top " & number & " of crash rate in intersections ")
                        Else 'Total
                            DrOp.Items.Add("Top " & number & " of crash rate in intersections and sections")
                        End If
                    Next i
                    queryString = queryStringRate
                End If
            End If

			Dim sResource As String = ""
            '   Dim exp As String = "WHERE (CSECT='854-23' AND LOGMI_FROM=12.77 AND LOGMI_TO=12.83) OR(CSECT='805-06' AND LOGMI_FROM=0 AND LOGMI_TO=0.05)  OR (CSECT='801-30' AND LOGMI_FROM=0 AND LOGMI_TO=0.24)"
            Dim exp As String = queryString
            exp = Replace(exp, "[", "")
            exp = Replace(exp, "]", "")
            exp = Replace(exp, "WHERE", "")

            Dim layerName As String = "section2000"
            loadMap(exp, sResource, layerName)

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

                If QueryType <> 9 Then 'for special exp for special year(because in the map, the data are different)
                    If optstateChosen Then
                        layerName = "Accidents2000"
                    ElseIf optDistChosen Then
                        layerName = "a2000D" & distNo

                    ElseIf optsUrbanChosen Then


                    End If

                Else
                    layerName = "Accidents2000"
                End If
                rdef = "(default)@ac2000"
            ElseIf indi.year = "2001" Then

                If QueryType <> 9 Then 'for special exp for special year(because in the map, the data are different)
                    If optstateChosen Then
                        layerName = "Crashes_State_2001"
                    ElseIf optDistChosen Then
                        layerName = "a2001D" & distNo
                    ElseIf optsUrbanChosen Then


                    End If

                Else
                    layerName = "Crashes_State_2001"
                End If
                rdef = "(default)@ac2001"
            ElseIf indi.year = "2002" Then

                If QueryType <> 9 Then 'for special exp for special year(because in the map, the data are different)
                    If optstateChosen Then
                        layerName = "Crashes_State_2002"
                    ElseIf optDistChosen Then
                        layerName = "a2002D" & distNo
                    ElseIf optsUrbanChosen Then


                    End If
                Else
                    layerName = "Crashes_State_2002"

                End If
                rdef = "(default)@ac2002"
            ElseIf indi.year = "2003" Then

                If QueryType <> 9 Then 'for special exp for special year(because in the map, the data are different)
                    If optstateChosen Then
                        layerName = "Crashes_State_2003"
                    ElseIf optDistChosen Then
                        layerName = "a2003D" & distNo
                    ElseIf optsUrbanChosen Then


                    End If

                Else
                    layerName = "Crashes_State_2003"
                End If
                rdef = "(default)@ac2003"
            ElseIf indi.year = "2004" Then

                If QueryType <> 9 Then 'for special exp for special year(because in the map, the data are different)
                    If optstateChosen Then
                        layerName = "Crashes_State_2004"
                    ElseIf optDistChosen Then
                        layerName = "a2004D" & distNo
                    ElseIf optsUrbanChosen Then


                    End If
                Else
                    layerName = "Crashes_State_2004"

                End If

                rdef = "(default)@ac2004"
            ElseIf indi.year = "2005" Then

                If QueryType <> 9 Then 'for special exp for special year(because in the map, the data are different)
                    If optstateChosen Then
                        layerName = "ACC2005"
                    ElseIf optDistChosen Then
                        layerName = "a2005D" & distNo
                    ElseIf optsUrbanChosen Then


                    End If

                Else
                    layerName = "ACC2005"
                End If

                rdef = "(default)@ac2005"
            ElseIf indi.year = "2006" Then

                If QueryType <> 9 Then 'for special exp for special year(because in the map, the data are different)
                    ' exp = ChangeExp(exp, 2004)
                    If optstateChosen Then
                        layerName = "ACC2006"
                    ElseIf optDistChosen Then
                        layerName = "a2006D" & distNo
                    ElseIf optsUrbanChosen Then
                        ' layerName = "La2006"
                        layerName = "la2006city"
                    End If

                Else
                    layerName = "ACC2006"
                End If
                rdef = "(default)@ac2006"

            End If

            If QueryType = 3 Or QueryType = 4 Then 'blackspot or totalhighway
                layerName = "csect2000"
                IsCsect = True

                'Showstreet.Visible = False
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
                ' exp =" 1=1  AND ROUTE = '0010'  AND DAY_OF_WK='1'"
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
                Dim strid, csectid As String
                id = GetLayerId(layerName, qf)
                strid = GetLayerId("section2000", qf)
                csectid = GetLayerId("csect2000", qf)
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
                streetlay = mapDpt.LayerDescriptions(CInt(csectid))
                highwaylay = mapDpt.LayerDescriptions(CInt(strid))
				'clear the section layer's difinition expression

If Not (QueryType = 3 Or QueryType = 4) Then
	 streetlay.DefinitionExpression = ""
End If

                Session("streetlay") = streetlay
                Session("highwaylay") = highwaylay

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

					'mapName = mrl.MapServer.MapName(0)
					'Dim lids() As Object = Nothing
					'Dim flds As String() = qf.GetFields(Nothing, id)

					'Dim scoll As New ESRI.ArcGIS.ADF.StringCollection(flds)
					'queryfilter.SubFields = scoll
					'' queryfilter.ReturnADFGeometries = True

					'Dim datatbl As System.Data.DataTable = qf.Query(Nothing, id, queryfilter)

					'Session("datatable") = datatbl
					'If sResourcename = "QuerySelection" Then
					'    Session("wholeselect") = datatbl
					'End If
					'Dim drs_s As DataRowCollection = datatbl.Rows

					'Dim shpind As Integer = -1
					'Dim j As Integer
					'For j = 0 To datatbl.Columns.Count - 1
					'    If datatbl.Columns(j).DataType Is GetType(ESRI.ArcGIS.ADF.Web.Geometry.Geometry) Then
					'        shpind = j
					'        Exit For
					'    End If
					'Next j
					'' Dim mapctrl As ESRI.ArcGIS.ADF.Web.UI.WebControls.Map = CType(Map1, ESRI.ArcGIS.ADF.Web.UI.WebControls.Map)



					'For Each gfunc_s In gfc_s
					'    If gfunc_s.Resource.Name = sResourcename Then
					'        gResource_s = CType(gfunc_s.Resource, ESRI.ArcGIS.ADF.Web.DataSources.Graphics.MapResource)
					'        Exit For
					'    End If
					'Next gfunc_s

					'If gResource_s Is Nothing Then
					'    Throw New Exception("Selection Graphics layer not in MapResourceManager")
					'End If
					'Dim gselectionlayer As ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer = Nothing

					'Dim dt_s As System.Data.DataTable
					'For Each dt_s In gResource_s.Graphics.Tables
					'    If TypeOf dt_s Is ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer Then
					'        gselectionlayer = CType(dt_s, ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer)
					'        Exit For
					'    End If
					'Next dt_s

					'If gselectionlayer Is Nothing Then
					'    gselectionlayer = New ESRI.ArcGIS.ADF.Web.Display.Graphics.ElementGraphicsLayer()
					'    gResource_s.Graphics.Tables.Add(gselectionlayer)
					'End If

					'Dim dr_s As DataRow
					'For Each dr_s In drs_s
					'    Dim geom_s As ESRI.ArcGIS.ADF.Web.Geometry.Geometry = CType(dr_s(shpind), ESRI.ArcGIS.ADF.Web.Geometry.Geometry)
					'    Dim ge_s As ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement = Nothing

				   mapName = mrl.MapServer.MapName(0)
					Dim lids() As Object = Nothing
					Dim flds As String() = qf.GetFields(Nothing, id)

					Dim scoll As New ESRI.ArcGIS.ADF.StringCollection(flds)
					queryfilter.SubFields = scoll
					' queryfilter.ReturnADFGeometries = True


					Dim qfilter2 As New ESRI.ArcGIS.ADF.ArcGISServer.QueryFilter
					qfilter2.WhereClause = exp

					Dim count As Integer
					count = mapSP.QueryFeatureCount(mapName, CInt(id), qfilter2)

					Dim queryset As New RecordSet
					queryset = mapSP.QueryFeatureData(mapName, CInt(id), qfilter2)


					Dim datatbl As New System.Data.DataTable
					Dim columnlength, recordlength As Integer
					columnlength = queryset.Fields.FieldArray.Length
					recordlength = queryset.Records.Length
				   Dim k, m As Integer


					For k = 0 To columnlength - 1
					Dim name As String
					name = queryset.Fields.FieldArray(k).Name
					'Dim coltype As New ESRI.ArcGIS.ADF.ArcGISServer.esriFieldType
					'coltype = queryset.Fields.FieldArray(k).Type

					Dim col As New DataColumn
					col.ColumnName = name
					'col.DataType = GetType(coltype)
					datatbl.Columns.Add(col)
					Next

 Dim geofild As ESRI.ArcGIS.ADF.ArcGISServer.Field = queryset.Fields.FieldArray(1)
 Dim geodef As String = geofild.GeometryDef.SpatialReference.WKT

For m = 0 To recordlength - 1
Dim rec As New Object
rec = queryset.Records(m).Values

datatbl.Rows.Add(queryset.Records(m).Values)
Next



					Session("datatable") = datatbl
					If sResourcename = "QuerySelection" Then
						Session("wholeselect") = datatbl
					End If
					Dim drs_s As Record() = queryset.Records

					Dim shpind As Integer = -1
					Dim j As Integer
					For j = 0 To columnlength - 1
						If datatbl.Columns(j).ColumnName = "Shape" Then
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

				   Dim z As Integer = 0
					Dim dr_s As Record
					For Each dr_s In drs_s

					'Dim geom_s As ESRI.ArcGIS.ADF.Web.Geometry.Geometry = CType(dr_s(shpind), ESRI.ArcGIS.ADF.Web.Geometry.Geometry)
					'	Dim ge_s As New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, System.Drawing.Color.Green)
					'	Dim cenpoint, testpoint As New ESRI.ArcGIS.ADF.Web.Geometry.Point
					'	cenpoint = GetCenterPoint(geom_s)
					'	 testpoint.M = 11.0
					'	 testpoint.X = -90
					'	 testpoint.Y = 30
					'	 testpoint.Z = -1
					'	 testpoint.SpatialReference = cenpoint.SpatialReference


					'	Dim geometrytest As ESRI.ArcGIS.ADF.Web.Geometry.Geometry = CType(cenpoint, ESRI.ArcGIS.ADF.Web.Geometry.Geometry)
					'	Dim geoanothertest As ESRI.ArcGIS.ADF.Web.Geometry.Geometry = CType(testpoint, ESRI.ArcGIS.ADF.Web.Geometry.Geometry)



					'	 ge_s.Symbol.Transparency = 30.0

						 Dim pointN As ESRI.ArcGIS.ADF.ArcGISServer.PointN
					   If TypeOf (dr_s.Values(shpind)) Is ESRI.ArcGIS.ADF.ArcGISServer.PointN Then
						   pointN = dr_s.Values(shpind)
							z = z + 1
						Else
						   Dim multipoint As ESRI.ArcGIS.ADF.ArcGISServer.MultipointN = dr_s.Values(shpind)
							If Not (multipoint.PointArray.Length = 0) Then
							   pointN = multipoint.PointArray(0)
							   z = z + 1

							   Else

							   Continue For

							End If


					   End If


						Dim webpoint As New ESRI.ArcGIS.ADF.Web.Geometry.Point

					 webpoint.M = pointN.M
					 webpoint.X = pointN.X
					 webpoint.Y = pointN.Y
					 webpoint.Z = pointN.Z

					 Dim coord As New ESRI.ArcGIS.ADF.Web.SpatialReference.DefinitionSpatialReferenceInfo(geodef)
					 Dim spref As New ESRI.ArcGIS.ADF.Web.SpatialReference.SpatialReference
					 spref.CoordinateSystem = coord
					 webpoint.SpatialReference = spref

					 Dim geom_s As ESRI.ArcGIS.ADF.Web.Geometry.Geometry = CType(webpoint, ESRI.ArcGIS.ADF.Web.Geometry.Geometry)


					 Dim ge_s As ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement = Nothing


						If sResourcename = "QuerySelection" Then


							If QueryType = 3 Then
								Dim sym As New ESRI.ArcGIS.ADF.Web.Display.Symbol.SimpleLineSymbol

								sym.Width = 10.0
								sym.Color = Drawing.Color.Blue


								ge_s = New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, sym)
							Else
								ge_s = New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, System.Drawing.Color.Blue)


							End If


						Else		   'sResource ="PartialSelection"

							If QueryType = 3 Then		 'blackspot
								Dim sym As New ESRI.ArcGIS.ADF.Web.Display.Symbol.SimpleLineSymbol

								sym.Width = 9.0
								' sym.FillType = ESRI.ArcGIS.ADF.Web.Display.Symbol.PolygonFillType.Solid

								sym.Color = Drawing.Color.DeepPink



								ge_s = New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, sym)
							Else
								ge_s = New ESRI.ArcGIS.ADF.Web.Display.Graphics.GraphicElement(geom_s, System.Drawing.Color.Yellow)
							End If
						End If
						ge_s.Symbol.Transparency = 0.0
						gselectionlayer.Add(ge_s)

						'timer control trying
						'Dim myTimer As System.Timers.Timer


						'myTimer = New System.Timers.Timer
						'myTimer.Interval = 5000

						'myTimer.Enabled = True
						'myTimer.Start()

						'gselectionlayer.Clear()
						'myTimer.Stop()
						'myTimer.Enabled = False
						'gselectionlayer.Add(ge_s)


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

            Dim linelay As LayerDescription = Session("highwaylay")

            If Showstreet.Checked Then
                linelay.Visible = True
                Map1.Refresh()
            Else
                linelay.Visible = False
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
