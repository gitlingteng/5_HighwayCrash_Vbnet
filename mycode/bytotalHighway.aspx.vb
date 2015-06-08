Imports Microsoft.Office.Interop.Owc11
Imports System.Data.OleDb


Namespace Crashsafe


    Partial Class bytotalHighway
        Inherits System.Web.UI.Page

        '  global variable

        Dim strscript As String
        Dim optDistChosen As Boolean
        Dim optstateChosen As Boolean
        Dim optsUrbanChosen As Boolean
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

        Dim wholeselect As DataTable
        Dim BlackNum(800) As BlackMapshow
        Dim BlackRate(800) As BlackMapshow




        'local variables
        Dim showmap As Integer
        Dim ShowWholemap As Boolean
        Dim IsCsect As Boolean
        Dim ShowAcc As Boolean
        Dim OptMile As Boolean
        Dim OptDiv As Boolean
        Dim HigNum As String
        Public statemdb As String = "mdbdata\acc"
        Public rows(10) As Integer
        Dim times As Integer


        'common parameters
        Public rtValue As Boolean
        Private condsql As String
        Public AppPath As String = "C:\data"

        'These parameters are used in unit length model
        'Public xlApp As New Microsoft.Office.Interop.Excel.Application()
        Public xlBook As Microsoft.Office.Interop.Excel.Workbook
        Public xlSheet As Microsoft.Office.Interop.Excel.Worksheet
        Public xlsheet2 As Microsoft.Office.Interop.Excel.Worksheet
        Public xlsheet3 As Microsoft.Office.Interop.Excel.Worksheet
        Public timesUnit As Integer
        Public DivValue As Double
        Dim cstype As Type = Me.GetType()
        Public Structure AccInfo
            Public accrate As Double
            Public accnum As Integer
            Public accpos As Integer
        End Structure
        Public results(8, 2000) As AccInfo 'this array means can search 8 years data and at most 2000 parts of a HW

        'These parameters are used in by crash (by compound query)
        Public accDate As Date
        Public Month As String

        Public RoadAlgn As String
        Public AccType As String
        Public TotalResults As Integer = 0
        '***********************************************
        Public Structure AccInfoCrash
            Public accCsect As String 'Control Section
            Public accMilePost As String 'MilePost
            Public accNum As Integer 'Number of Crashes
            Public accFatalNum As Integer 'FatalInjuries
            Public accCriticalNum As Integer 'CriticalInjuries
            Public accSeriousNum As Integer 'SeriousInjuries
            Public accSevereNum As Integer 'SevereInjuries
            Public accModerateNum As Integer 'ModerateInjuries
            Public accMinorNum As Integer 'MinorInjuries
        End Structure
        Public resultsCrash(8, 5000) As AccInfoCrash 'this array means can search 8 years data and at most 5000 parts of a HW
        '***********************************************
        'These parameters are used in by control section (by particular highway)
        Public Structure AccInfoSect
            Public accCsect As String
            Public accRate As Double
            Public accSum As Integer
            Public accFrom As Double
            Public accTo As Double
        End Structure
        Public resultsSect(8, 5000) As AccInfoSect 'this array means can search 8 years data and at most 5000 parts of a HW
        Protected WithEvents BtnHelp1 As System.Web.UI.HtmlControls.HtmlInputButton
        Protected WithEvents labWeekDay As System.Web.UI.WebControls.Label



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

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            ' 'Put user code to initialize the page here
            'Dim login As String
            'login = Session("succ")
            'If login = "" Then ' fake user
            '    Response.Redirect("LoginIn.aspx")
            'End If


            optstateChosen = Session("optstateChosen")
            optDistChosen = Session("optDistChosen")
            optsUrbanChosen = Session("optsUrbanChosen")
            distNo = Session("distNo")
            QueryType = Session("QueryType")
            ShowAcc = Session("ShowAcc")


            If Rbmile.Checked Then
                OptMile = True
                OptDiv = False
            End If

        End Sub

        Private Sub Settrue(ByVal ok As Integer)
            If ok = 0 Then 'by crash
                ctrlSect1.Visible = False
                ctrlSect.Visible = False
                cmbMonth.Enabled = True
                cmbWeekday.Enabled = True
                cmbHourFrom.Enabled = True
                cmbHourTo.Enabled = True
                cmbWeather.Enabled = True
                cmbSufaceType.Enabled = True
                fstSpeed.Enabled = True
                cmbSufaceCond.Enabled = True
                cmbViolation1.Enabled = True
                cmbRoadwayAlign.Enabled = True
                cmbTypeofColl.Enabled = True
                cmb1driver.Enabled = True
                cmbRoadCond.Enabled = True
                cmbTypeofAcc.Enabled = True
                '
                labHourFrom.Enabled = True
                labHourTo.Enabled = True
                labMonth.Enabled = True
                labWeek.Enabled = True
                labWeather.Enabled = True
                labSufTyp.Enabled = True
                labSpeed.Enabled = True
                labSufCon.Enabled = True
                labVio.Enabled = True
                labAlign.Enabled = True
                labColl.Enabled = True
                labDriv.Enabled = True
                labRoadCon.Enabled = True
                labAcc.Enabled = True
                '
                Rbmile.Enabled = False
                labText1.Enabled = False
                labText2.Enabled = False
                RbDiv.Enabled = False
                labText3.Enabled = False
                'TxtPart.Enabled = False
                labText4.Enabled = False
                TxtMile.Enabled = False

                '
                Button2.Visible = False
                Button3.Visible = False
                labLogmile.Enabled = False
                lblLogFrom.Enabled = False
                LabInter.Enabled = False
                cmbInter.Enabled = False
            ElseIf ok = 1 Or ok = 2 Then 'by control section and unit length
                Button2.Visible = True

                Button2.Text = ">"
                cmbMonth.Enabled = False
                cmbWeekday.Enabled = False
                cmbHourFrom.Enabled = False
                cmbHourTo.Enabled = False
                cmbWeather.Enabled = False
                cmbSufaceType.Enabled = False
                fstSpeed.Enabled = False
                cmbSufaceCond.Enabled = False
                cmbViolation1.Enabled = False
                cmbRoadwayAlign.Enabled = False
                cmbTypeofColl.Enabled = False
                cmb1driver.Enabled = False
                cmbRoadCond.Enabled = False
                cmbTypeofAcc.Enabled = False
                '
                labHourFrom.Enabled = False
                labHourTo.Enabled = False
                labMonth.Enabled = False
                labWeek.Enabled = False
                labWeather.Enabled = False
                labSufTyp.Enabled = False
                labSpeed.Enabled = False
                labSufCon.Enabled = False
                labVio.Enabled = False
                labAlign.Enabled = False
                labColl.Enabled = False
                labDriv.Enabled = False
                labRoadCon.Enabled = False
                labAcc.Enabled = False
                ctrlSect1.Visible = False
                Button2.Visible = False
                If ok = 1 Then 'by section control  
                    Button2.Visible = False
                    Button3.Visible = False
                    ctrlSect.Visible = False
                    Rbmile.Enabled = False
                    labText1.Enabled = False
                    labText2.Enabled = False
                    RbDiv.Enabled = False
                    labText3.Enabled = False
                    'TxtPart.Enabled = False
                    labText4.Enabled = False
                    TxtMile.Enabled = False
                    '
                    labLogmile.Enabled = False
                    lblLogFrom.Enabled = False
                    LabInter.Enabled = False
                    cmbInter.Enabled = False
                ElseIf ok = 2 Then 'by unit length
                    Button2.Visible = True
                    Rbmile.Enabled = True
                    labText1.Enabled = True
                    labText2.Enabled = True
                    RbDiv.Enabled = True
                    labText3.Enabled = True
                    'TxtPart.Enabled = False
                    labText4.Enabled = True
                    TxtMile.Enabled = True
                    '
                    LabInter.Enabled = True
                    cmbInter.Enabled = True
                End If
            End If
        End Sub

        Private Sub radControl_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radControl.CheckedChanged
            'base on control section
            Settrue(1)
            labSect.ForeColor = System.Drawing.Color.Red
            labCrash.ForeColor = System.Drawing.Color.Black
            labUnit.ForeColor = System.Drawing.Color.Black
            radControl.Checked = True
            radCrash.Checked = False
            radUnit.Checked = False
            QueryType = 4 'by control section(by partial highway)
        End Sub

        Private Sub radCrash_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radCrash.CheckedChanged
            'base on crash
            Settrue(0)
            labSect.ForeColor = System.Drawing.Color.Black
            labCrash.ForeColor = System.Drawing.Color.Red
            labUnit.ForeColor = System.Drawing.Color.Black
            radControl.Checked = False
            radCrash.Checked = True
            radUnit.Checked = False
            QueryType = 5 'by crash(by compound query)
        End Sub

        Private Sub radUnit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radUnit.CheckedChanged
            'base on unit length
            Settrue(2)
            labSect.ForeColor = System.Drawing.Color.Black
            labCrash.ForeColor = System.Drawing.Color.Black
            labUnit.ForeColor = System.Drawing.Color.Red
            radControl.Checked = False
            radCrash.Checked = False
            radUnit.Checked = True
            QueryType = 7 'by unit length(by sliding window)
        End Sub

        Private Sub YearToFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YearToFrom.Click
            Dim a As String
            a = YearToFrom.Text
            If a = ">" Then
                cmbYear2.Visible = True
                cmbYear2.Enabled = True
                YearToFrom.Text = "<"
            End If
            If a = "<" Then
                cmbYear2.Visible = False
                cmbYear2.Enabled = False
                YearToFrom.Text = ">"
            End If
        End Sub

        Private Sub Button1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.ServerClick
            If cmbYear.SelectedIndex <> 0 And Len(txtHwyNum.Text) <> 0 Then
                If ctrlSect.Enabled And ctrlSect.SelectedIndex <> 0 Then 'by unit length(by sliding window)
                    'select control section for highway
                    Dim accConn As ADODB.Connection
                    Dim rst As ADODB.Recordset
                    Dim accDB As String
                    Dim condsql As String

                    accConn = New ADODB.Connection()
                    'rst = New ADODB.Recordset
                    accDB = "C:\DATA\mdbdata\section2004.mdb"
                    With accConn
                        'Telling ADO to use JOLT Here
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .Open(accDB)
                    End With

                    condsql = "SELECT MIN(MIPOST_FR),MAX(MIPOST_TO) FROM section2004 WHERE HWY_NUM='" & GetTrueHnum(txtHwyNum.Text) & "'" & "AND [CSECT]='" & ctrlSect.SelectedItem.Text & "'"
                    'condsql = "SELECT MIN[MIPOST_FR] as min_mark  FROM section2004"
                    rst = New ADODB.Recordset()
                    rst.Open(condsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    With rst
                        If .RecordCount < 1 Then
                            Exit Sub
                        Else
                            rst.MoveFirst()
                            txtMileFrom.Text = .Fields(0).Value
                            txtMileTo.Text = .Fields(1).Value
                        End If
                    End With

                    accConn.Close()
                    rst = Nothing

                Else
                    txtMileFrom.Text = "0"
                    indi.year = cmbYear.SelectedItem.Text
                    txtMileTo.Text = CStr(GetMileTo(indi.year, GetTrueHnum(txtHwyNum.Text)))
                End If
            End If
        End Sub


        Private Function GetMileTo(ByVal year As String, ByVal num As String) As Double
            'System table
            Dim tempSql As String
            'ADO Objects Used
            Dim secConn As ADODB.Connection
			Dim rst As ADODB.Recordset
			Dim dbyear As Integer

            On Error GoTo GetMileTo_Err

            secConn = New ADODB.Connection()
            Dim dbfConnect As String
            dbfConnect = "Driver={Microsoft dBase Driver (*.dbf)};" & _
              "Dbq=" & AppPath & "\GisPrj;" & _
              "DefaultDir=" & AppPath & ";" & _
              "Uid=Admin;Pwd=;"
            secConn.Open(dbfConnect)

			If indi.year = "2006" Then
				 dbyear = 2005
		   Else
			   dbyear = indi.year
		 End If

			tempSql = "SELECT  MAX([Mipost_to]) " & _
			   "FROM sec" & dbyear & ".dbf" & _
			   " WHERE [HWY_NUM] = '" & num & "' "
            rst = New ADODB.Recordset()
            rst.Open(tempSql, secConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            With rst
                If (.RecordCount < 1) Or IsDBNull(.Fields(0).Value) Then
                    GetMileTo = 0
                Else
                    GetMileTo = .Fields(0).Value
                End If
            End With

GetMileTo_Exit:
            If Not (rst Is Nothing) Then
                If (rst.State And ConnectionState.Open) = ConnectionState.Open Then
                    rst.Close()
                End If
                rst = Nothing
            End If
            If Not (secConn Is Nothing) Then
                If (secConn.State And ConnectionState.Open) = ConnectionState.Open Then
                    secConn.Close()
                End If
                secConn = Nothing
            End If
            Exit Function
GetMileTo_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error Number and Description' & Err.Num & Err.Description)"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            GetMileTo = 0
            Resume GetMileTo_Exit
        End Function

        Private Sub Rbmile_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rbmile.CheckedChanged
            If Rbmile.Checked Then
                TxtMile.Enabled = True
                TxtPart.Enabled = False
                RbDiv.Checked = False
                TxtPart.Text = ""
                TxtMile.Text = "5"
                TxtPart.BackColor = System.Drawing.Color.Gainsboro
                TxtMile.BackColor = System.Drawing.Color.White
                OptMile = True
                OptDiv = False
            End If
        End Sub

        Private Sub RbDiv_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbDiv.CheckedChanged
            If RbDiv.Checked Then
                TxtMile.Enabled = False
                TxtPart.Enabled = True
                Rbmile.Checked = False
                TxtMile.Text = ""
                TxtMile.BackColor = System.Drawing.Color.Gainsboro
                TxtPart.BackColor = System.Drawing.Color.White
                OptDiv = True
                OptMile = False
            End If
        End Sub

        Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click

            'multiple year
            If cmbYear2.Enabled = True Then
                If CInt(cmbYear.SelectedIndex.ToString) >= CInt(cmbYear2.SelectedIndex.ToString) And cmbYear2.SelectedIndex.ToString <> "0" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The YearFrom can not be greater than or equal to YearTo!' )"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                End If
                mutipleYear = True
                YearNum = cmbYear2.SelectedIndex.ToString - cmbYear.SelectedIndex.ToString
                If cmbYear2.SelectedIndex = 0 Then
                    mutipleYear = False
                    YearNum = 0
                End If
            Else
                mutipleYear = False
                YearNum = 0
            End If
            'Check the year value
            If Len(cmbYear.SelectedItem.Text) < 4 Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please enter year!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            Else
                indi.year = cmbYear.SelectedItem.Text
            End If
            'Highway number
            If Len(txtHwyNum.Text) = 0 Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please enter highway number!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            Else
                indi.WayNum = GetTrueHnum(txtHwyNum.Text)
                HigNum = GetTrueHnum(txtHwyNum.Text) 'used in information.aspx
            End If
            ' MilePost
            If Len(txtMileFrom.Text) > 0 And IsNumeric(txtMileFrom.Text) Then
                indi.MileFrom = txtMileFrom.Text
            Else
                indi.MileFrom = 0
            End If
            If Len(txtMileTo.Text) > 0 And IsNumeric(txtMileTo.Text) Then
                indi.MileTo = txtMileTo.Text
            Else
                indi.MileTo = GetMileTo(indi.year, indi.WayNum)
            End If
            If (indi.MileTo = 0) Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('No Record in Segmetation table for this Highway Number.' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
            If (CInt(indi.MileTo) < CInt(indi.MileFrom)) Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('The value of start point is less than end point.' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
            txtMileFrom.Text = CStr(indi.MileFrom)
            txtMileTo.Text = CStr(indi.MileTo)

            rtValue = False

            If radControl.Checked Then 'at first,it is by control section be chosen
                QueryType = 4 'by control section(by partial highway)
            End If
            If radCrash.Checked Then 'at first,it is by control section be chosen
                QueryType = 5 'based on crash(by partial highway)
            End If
            If radUnit.Checked Then
                QueryType = 7 'base on selected unit length(by partial highway)
            End If

            If QueryType = 4 Then 'by control section(by partial highway)
                Dim i As Integer
				If indi.year = "1999" Or indi.year = "2006" Then
					strscript = "<script language='javascript'>"
					strscript = strscript & "alert('No section data for this year!')"
					strscript = strscript & "</script>"
					ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
				   Exit Sub
				End If

                For i = 0 To YearNum
                    Call GenerateSectReport(indi.year + i, i)
                Next i

                If rtValue Then
                    Call drawSectChart()
                End If

                Session("queryString") = queryString
                Session("QueryType") = QueryType
                Session("SubQueryType") = SubQueryType
               
                Session("indi") = indi
                Session("caption") = caption
                Session("caption1") = caption1
                Session("HowManyResult") = HowManyResult
                Session("BlackNum") = BlackNum
                Session("BlackRate") = BlackRate
               

                If rtValue Then
                    NoShow = True
                    Session("NoShow") = NoShow

                    Response.Write("<script languge='javascript'>window.open('repSpcHway.aspx')</script>;")
                    'Response.Redirect("repSpcHway.aspx")
                Else
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('No results found!')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                End If
            ElseIf QueryType = 5 Then 'by crash (by compound query)
                Dim i As Integer
                'Hour
                If cmbHourFrom.SelectedIndex > 0 Then
                    indi.Hourfrom = cmbHourFrom.SelectedItem.text
                Else
                    indi.Hourfrom = 0
                End If
                If cmbHourTo.SelectedIndex > 0 Then
                    indi.Hourto = cmbHourTo.SelectedItem.Text
                Else
                    indi.Hourto = 23
                End If
                If CInt(indi.Hourfrom) > CInt(indi.Hourto) Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('Please make sure that the HourFrom is less than HourTo!' )"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                End If
                '*******************************
                If cmbWeekday.SelectedIndex > 0 Then
                    indi.WeekDay = cmbWeekday.SelectedValue
                Else
                    indi.WeekDay = ""
                End If
                If cmbWeather.SelectedIndex > 0 Then
                    indi.Weather = IndexToAsc(cmbWeather)
                Else
                    indi.Weather = ""
                End If

                If cmbSufaceType.SelectedIndex > 0 Then
                    indi.SurfaceType = IndexToAsc(cmbSufaceType)
                Else
                    indi.SurfaceType = ""
                End If


                If Len(fstSpeed.Text) > 0 Then
                    indi.Veh1Speed = fstSpeed.Text
                Else
                    indi.Veh1Speed = ""
                End If


                If cmbSufaceCond.SelectedIndex > 0 Then
                    indi.SurfaceCond = IndexToAsc(cmbSufaceCond)
                Else
                    indi.SurfaceCond = ""
                End If

                If cmbViolation1.SelectedIndex > 0 Then
                    indi.Violation1 = IndexToAsc(cmbViolation1)
                Else
                    indi.Violation1 = ""
                End If

                If cmb1driver.SelectedIndex > 0 Then
                    indi.Driver1Cond = IndexToAsc(cmb1driver)
                Else
                    indi.Driver1Cond = ""
                End If

                If cmbTypeofColl.SelectedIndex > 0 Then
                    indi.CollType = IndexToAsc(cmbTypeofColl)
                Else
                    indi.CollType = ""
                End If

                If cmbRoadCond.SelectedIndex > 0 Then
                    indi.RoadCond = IndexToAsc(cmbRoadCond)
                Else
                    indi.RoadCond = ""
                End If

                '******************
                If cmbRoadwayAlign.SelectedIndex > 0 Then
                    RoadAlgn = IndexToAsc(cmbRoadwayAlign)
                Else
                    RoadAlgn = ""
                End If

                If cmbTypeofAcc.SelectedIndex > 0 Then
                    AccType = IndexToAsc(cmbTypeofAcc)
                Else
                    AccType = ""
                End If

                '*****
                For i = 0 To YearNum
                    'Month
                    If cmbMonth.SelectedIndex > 0 Then
                        accDate = DateSerial(CInt(indi.year + i), cmbMonth.SelectedIndex, 1)
                        Month = CStr(cmbMonth.SelectedIndex)
                        'Month1 = cmbMonth.Text
                    Else
                        'accDate = Empty'don't know how to make it be empty
                        Month = ""
                    End If
                    Call genCrashResults(indi.year + i, i)
                Next i
                '

             

                If rtValue Then
                    Call drawCrashChart()
                Else
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('No records found for the selections made.')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                End If
                Session("queryString") = queryString
                Session("QueryType") = QueryType
                Session("SubQueryType") = SubQueryType

                Session("indi") = indi
                Session("caption") = caption
                Session("caption1") = caption1
                Session("HowManyResult") = HowManyResult
                Session("BlackNum") = BlackNum
                Session("BlackRate") = BlackRate



                If rtValue Then
                    NoShow = True
                    Session("NoShow") = NoShow
                    Response.Write("<script languge='javascript'>window.open('repComQry.aspx')</script>;")
                    'Response.Redirect("repComQry.aspx")
                Else
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('Something worng happened when the program try to create the chart!')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                End If
            ElseIf QueryType = 7 Then  'unit length(sliding window)
                'chose the data that meet the demand 
                If ctrlSect.SelectedIndex > 0 Then
                    indi.Csect = ctrlSect.SelectedItem.Text
                Else
                    indi.Csect = ""

                End If

                Dim i As Integer
                For i = 0 To YearNum
                    Call GenerateUnitReport(indi.year + i, i)
                Next i

               


                If rtValue Then
                    Call drawUnitChart()
                Else
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('No results found!')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                End If

                Session("queryString") = queryString
                Session("QueryType") = QueryType
                Session("SubQueryType") = SubQueryType

                Session("indi") = indi
                Session("caption") = caption
                Session("caption1") = caption1
                Session("HowManyResult") = HowManyResult
                Session("BlackNum") = BlackNum
                Session("BlackRate") = BlackRate


                If rtValue Then
                    NoShow = True
                    Session("NoShow") = NoShow
                    Response.Write("<script languge='javascript'>window.open('repPartiWin.aspx')</script>;")
                    'Response.Redirect("repPartiwin.aspx")
                Else
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('No results found!')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                End If
            End If
        End Sub

        '*************************************
        Private Function GenerateUnitReport(ByVal year As Integer, ByVal num As Integer)
            Dim strSource As String
            Dim strDestination As String

            On Error GoTo Proc_Err

            'xlApp = CreateObject("Microsoft.Office.Interop.Excel.Application")
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application()

            xlApp.Visible = False
            strSource = AppPath & "\QryByPartition.xls"

            Dim ChartName As String = Rnd()
            'strDestination = AppPath & "\tempByParti" & year & ".xls"
            strDestination = AppPath & "\tempByParti" & ChartName & year & ".xls"
            Session(CStr(year) & "partwin") = strDestination

            FileCopy(strSource, strDestination)
            xlBook = xlApp.Workbooks.Open(strDestination)
            xlSheet = xlBook.Worksheets(1)
            xlsheet2 = xlBook.Worksheets.Add
            xlsheet3 = xlBook.Worksheets.Add

            With xlsheet2
                .Cells(1, 1).value = "Ranked by Crash Rate"
                .Cells(1, 2).value = "Crash Rate"
                .Cells(1, 3).value = "Number of Crashes"
                .Cells(1, 4).value = "Mile from"
                .Cells(1, 5).value = "Mile to"
            End With

            With xlsheet3
                .Cells(1, 1).value = "Ranked by Num of Acc"
                .Cells(1, 2).value = "Number of Crashes"
                .Cells(1, 3).value = "Crash Rate"
                .Cells(1, 4).value = "Mile from"
                .Cells(1, 5).value = "Mile to"
            End With


            Call CompoundUnitQuery(num, xlBook, xlSheet, xlsheet2, xlsheet3)

            xlBook.Save()
            xlApp.ActiveWorkbook.Saved = True
            xlApp.DisplayAlerts = False
            xlApp.ActiveWorkbook.Close()

            'close the workbook
            xlApp.Workbooks.Close()
            xlApp.DisplayAlerts = True

            'close the Microsoft.Office.Interop.Excel
            xlApp.Quit()
            xlApp = Nothing
            xlSheet = Nothing
            xlsheet2 = Nothing
            xlsheet3 = Nothing

            xlBook = Nothing
            GC.Collect()
Proc_Exit:
            'MousePointer = vbNormal
            Exit Function
Proc_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error'& Err.Description )"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)

            If Not IsNothing(xlApp) Then
                xlApp.Quit()
            End If
            xlApp = Nothing
            xlSheet = Nothing
            xlsheet2 = Nothing
            xlsheet3 = Nothing
            xlBook = Nothing

            Resume Proc_Exit
        End Function

        '*************************************
        Public Sub CompoundUnitQuery(ByVal num As Integer, ByVal xlBook As Microsoft.Office.Interop.Excel.Workbook, ByVal xlSheet As Microsoft.Office.Interop.Excel.Worksheet, ByVal xlSheet2 As Microsoft.Office.Interop.Excel.Worksheet, ByVal xlSheet3 As Microsoft.Office.Interop.Excel.Worksheet)
            On Error GoTo Proc_Err

            Dim i, j As Integer 'used in "for loop"
            Dim mileType As String
            mileType = "MILE_POST"

            Dim Conn As ADODB.Connection
            Dim rst As ADODB.Recordset

            Conn = New ADODB.Connection()
            rst = New ADODB.Recordset() 'store the data of top crash rate
            connString = "\" & statemdb & indi.year + num & ".mdb"
            With Conn
                'Telling ADO to use JOLT Here
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Open(AppPath & connString)
            End With

            'get the sql sentence
            If indi.Csect = "" Then
                condsql = " WHERE [HWY_NUM] = '" & indi.WayNum & "'"
            Else
                condsql = " WHERE [HWY_NUM] = '" & indi.WayNum & "'AND [CSECT] = '" & indi.Csect & "'"
            End If
            BlackNum(0).Csect = indi.Csect 'actual here we store the highway number, and used it in map show

            '*****************************Maybe Have Problem
            'condsql = condsql & " AND (" & mileType & " >= " & indi.MileFrom & _
            '          " AND " & mileType & " <= " & indi.MileTo & ")"
            queryString = condsql

            If cmbInter.SelectedIndex <> 0 Then
                condsql = condsql & " AND (" & mileType & " BETWEEN " & indi.MileFrom & _
                          " AND " & indi.MileTo & ")" & "AND [INTER]='" & cmbInter.SelectedIndex - 1 & "'"

                queryString = queryString & " AND (" & mileType & " >= " & indi.MileFrom & _
                          " AND " & mileType & " < " & indi.MileTo & ")" & "AND [INTER]='" & cmbInter.SelectedIndex - 1 & "'"
            Else
                condsql = condsql & " AND (" & mileType & " BETWEEN " & indi.MileFrom & _
                                  " AND " & indi.MileTo & ")"

                queryString = queryString & " AND (" & mileType & " >= " & indi.MileFrom & _
                          " AND " & mileType & " < " & indi.MileTo & ")"


            End If
            '

            '*to get the number of acc,ask dr sun what is the demand in the search****************************************
            condsql = "SELECT [MILE_POST],[CSECT], [LOGMI_FROM],[LOGMI_TO],[ADT], count(*) " & _
             "FROM accidents" & indi.year + num & condsql & _
              " group BY [MILE_POST],[CSECT],[LOGMI_FROM],[LOGMI_TO],[ADT]"

            rst.Open(condsql, Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'rstAdt.Open(sqlAdt, ConnAdt, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            'calculate the crash rate and Number of Crashof every partitional part
            timesUnit = 0
            Dim cycle As Integer
            Dim incValue = 1
            If OptMile Then
                DivValue = CDbl(TxtMile.Text)
                timesUnit = CInt((CDbl(indi.MileTo) - CDbl(indi.MileFrom)) / DivValue)
                If timesUnit > 2000 Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The number of divided parts is too large,please enlarge the miles of each part.')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    GoTo Proc_Exit
                End If
            Else
                timesUnit = CInt(TxtPart.Text)
                If timesUnit > 2000 Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The number of divided parts can not be greater than 2000,please enter a less number.')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    GoTo Proc_Exit
                End If
                DivValue = CInt((CDbl(indi.MileTo) - CDbl(indi.MileFrom)) / timesUnit)
            End If

            'Dim ADT(times + 1) As Double
            Dim accInfo(timesUnit + 1) As AccInfo
            For i = 1 To timesUnit + 1
                accInfo(i).accnum = 0
                accInfo(i).accpos = i
                accInfo(i).accrate = 0
            Next
            Dim milepost, addNum, addDistant, lastLogmileTo, thisLogmileTo, ADT
            Dim thisCsect, lastCsect As String
            Dim Singlecrashrate As Double
            Dim nextSum As Boolean = False
            Dim NumofAcc, lastCycle As Integer
            thisCsect = " 0"
            lastCsect = "0"
            addNum = 0
            lastLogmileTo = 0
            thisLogmileTo = 0
            Singlecrashrate = 0
            With rst
                If .RecordCount < 1 Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('No records found for the selections made!')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    rtValue = False
                    GoTo Proc_Exit
                Else
                    rst.MoveFirst()
                    addNum = 1
                    Do While (Not .EOF)
                        '  condsql = "SELECT [MILE_POST],[CSECT], [LOGMI_FROM],[LOGMI_TO],[ADT], count(*) " & _
                        '"FROM accidents" & indi.year + num & condsql & _
                        ' " group BY [MILE_POST]"

                        On Error GoTo Proc_Err
                        milepost = CDbl(.Fields(0).Value)  '[MILE_POST]
                        For cycle = incValue To timesUnit 'Dim incValue = 1
                            If ((milepost - indi.MileFrom) < DivValue * cycle) And ((milepost - indi.MileFrom) >= DivValue * (cycle - 1)) Then
                                accInfo(cycle).accnum = .Fields(5).Value + accInfo(cycle).accnum '[count(*)]
                                NumofAcc = .Fields(5).Value

                                thisCsect = .Fields(1).Value '[Csect]
                                thisLogmileTo = .Fields(3).Value '[LOGMILE_TO]
                                If lastCsect = thisCsect And thisLogmileTo = lastLogmileTo Then
                                    'caclulate average ADT
                                    ADT = .Fields(4).Value
                                    addNum = addNum + NumofAcc 'this number no use in here right now,maybe can be used for future
                                Else
                                    If addNum <> 1 Or (addNum = 1 And nextSum) Then
                                        Singlecrashrate = ADT * addDistant + Singlecrashrate
                                    End If
                                    addNum = NumofAcc
                                End If
                                nextSum = True 'to control the rocords which only have one accident

                                addDistant = .Fields(3).Value - .Fields(2).Value '[LOGMILE_TO]-[LOGMILE_FROM]
                                If CDbl(addDistant) < 0 Then
                                    addDistant = CDbl(addDistant) * (-1)
                                End If
                                lastLogmileTo = .Fields(3).Value '[LOGMILE_TO]
                                lastCsect = thisCsect
                                incValue = cycle
                                Exit For
                            Else
                                If (CInt(indi.year) Mod 4) = 0 Then
                                    Singlecrashrate = (ADT * addDistant + Singlecrashrate) * 366
                                    accInfo(cycle).accrate = Format(accInfo(cycle).accnum * 1000000 / Singlecrashrate, "0.00")     '***********************
                                Else
                                    Singlecrashrate = (ADT * addDistant + Singlecrashrate) * 365
                                    accInfo(cycle).accrate = Format(accInfo(cycle).accnum * 1000000 / Singlecrashrate, "0.00")  '***********************
                                End If
                                thisCsect = " 0"
                                lastCsect = "0"
                                addNum = 0
                                lastLogmileTo = 0
                                thisLogmileTo = 0
                                Singlecrashrate = 0
                                nextSum = False
                                addNum = 1
                            End If
                        Next cycle
                        .MoveNext()
                    Loop
                    rtValue = True
                End If
            End With

            'the rate of last cycle add in here
            If (CInt(indi.year) Mod 4) = 0 Then
                Singlecrashrate = (ADT * addDistant + Singlecrashrate) * 366
                accInfo(timesUnit).accrate = Format(accInfo(timesUnit).accnum * 1000000 / Singlecrashrate, "0.00")
            Else
                Singlecrashrate = (ADT * addDistant + Singlecrashrate) * 365
                accInfo(timesUnit).accrate = Format(accInfo(timesUnit).accnum * 1000000 / Singlecrashrate, "0.00")
            End If

            HowManyResult(0) = timesUnit ' used in map show

            'store the result to the array "results"
            For i = 1 To timesUnit
                results(num, i).accrate = accInfo(i).accrate 'Crash Rate
                results(num, i).accnum = accInfo(i).accnum 'Number of Crashes
                results(num, i).accpos = accInfo(i).accpos 'Position of accident
            Next

            'sorting the result by acc number,for selecting  top of acc number
            Dim CompareRate, GetRate, MidRate, ChangeRate As Double
            Dim CompareNum, GetNum, MidNum, ChangeNum, MidPos, ChangePos As Integer
            For i = 1 To timesUnit - 1
                CompareNum = accInfo(1).accnum
                MidPos = accInfo(1).accpos
                MidRate = accInfo(1).accrate
                For j = 2 To timesUnit + 1 - i
                    If CompareNum < accInfo(j).accnum Then
                        GetNum = accInfo(j).accnum
                        ChangePos = accInfo(j).accpos
                        ChangeRate = accInfo(j).accrate

                        accInfo(j).accnum = CompareNum
                        accInfo(j).accpos = MidPos
                        accInfo(j).accrate = MidRate

                        accInfo(j - 1).accnum = GetNum
                        accInfo(j - 1).accpos = ChangePos
                        accInfo(j - 1).accrate = ChangeRate
                    Else
                        CompareNum = accInfo(j).accnum
                        MidPos = accInfo(j).accpos
                        MidRate = accInfo(j).accrate
                    End If
                Next
            Next

            'add the decreasesing results of acc number to the excle table
            Dim rowcount As Integer = 3
            For i = 1 To timesUnit
                xlSheet3.Cells(rowcount, 1) = i 'Ranked by Num of Acc
                xlSheet3.Cells(rowcount, 2) = accInfo(i).accnum 'Number of Crashes
                xlSheet3.Cells(rowcount, 3) = accInfo(i).accrate  'Crash Rate
                xlSheet3.Cells(rowcount, 4) = (accInfo(i).accpos - 1) * DivValue + indi.MileFrom 'Mile from
                xlSheet3.Cells(rowcount, 5) = accInfo(i).accpos * DivValue + indi.MileFrom 'Mile to
                rowcount = rowcount + 1
                BlackNum(i - 1).accLogmifrom = (accInfo(i).accpos - 1) * DivValue + indi.MileFrom 'Mile from
                BlackNum(i - 1).accLogmito = accInfo(i).accpos * DivValue + indi.MileFrom 'Mile to
            Next


            'sorting the result by acc rate,for selecting  top of acc number
            For i = 1 To timesUnit - 1
                CompareRate = accInfo(1).accrate
                MidPos = accInfo(1).accpos
                MidNum = accInfo(1).accnum
                For j = 2 To timesUnit + 1 - i
                    If CompareRate < accInfo(j).accrate Then
                        GetRate = accInfo(j).accrate
                        ChangePos = accInfo(j).accpos
                        ChangeNum = accInfo(j).accnum

                        accInfo(j).accrate = CompareRate
                        accInfo(j).accpos = MidPos
                        accInfo(j).accnum = MidNum

                        accInfo(j - 1).accrate = GetRate
                        accInfo(j - 1).accpos = ChangePos
                        accInfo(j - 1).accnum = ChangeNum
                    Else
                        CompareRate = accInfo(j).accrate
                        MidPos = accInfo(j).accpos
                        MidNum = accInfo(j).accnum
                    End If
                Next
            Next
            'add the decreasesing results of acc rate to the excle table
            rowcount = 3
            For i = 1 To timesUnit
                xlSheet2.Cells(rowcount, 1) = i 'Ranked by Num of Acc
                xlSheet2.Cells(rowcount, 2) = accInfo(i).accrate  'Crash Rate
                xlSheet2.Cells(rowcount, 3) = accInfo(i).accnum 'Number of Crashes
                xlSheet2.Cells(rowcount, 4) = (accInfo(i).accpos - 1) * DivValue + indi.MileFrom 'Mile from
                xlSheet2.Cells(rowcount, 5) = accInfo(i).accpos * DivValue + indi.MileFrom 'Mile to
                rowcount = rowcount + 1
                BlackRate(i - 1).accLogmifrom = (accInfo(i).accpos - 1) * DivValue + indi.MileFrom 'Mile from
                BlackRate(i - 1).accLogmito = accInfo(i).accpos * DivValue + indi.MileFrom  'Mile to
            Next

Proc_Exit:
            If Not (Conn Is Nothing) Then
                Conn.Close()
            End If
            Conn = Nothing
            rst = Nothing

            Exit Sub
Proc_Err:
            Resume Proc_Exit
        End Sub

        Private Sub drawUnitChart()
            Dim objCSpace1 As Microsoft.Office.Interop.Owc11.ChartSpace = New Microsoft.Office.Interop.Owc11.ChartSpaceClass()
            Dim objCSpace2 As Microsoft.Office.Interop.Owc11.ChartSpace = New Microsoft.Office.Interop.Owc11.ChartSpaceClass()
            Dim objCSpace3 As Microsoft.Office.Interop.Owc11.ChartSpace = New Microsoft.Office.Interop.Owc11.ChartSpaceClass()
            Dim objChart1, objChart2, objChart3
            Dim hwy_class_name As String
            Dim chart_title As String
            Dim vv(YearNum), uu(YearNum), xxsum, rangStr As String
            Dim sum(timesUnit) As Double
            Dim i, j, k As Integer

            On Error GoTo Proc_Err

            caption = ""
            objChart1 = objCSpace1.Charts.Add(0)
            objChart1.Type = Microsoft.Office.Interop.Owc11.ChartChartTypeEnum.chChartTypeSmoothLine

            objChart2 = objCSpace2.Charts.Add(0)
            objChart2.Type = Microsoft.Office.Interop.Owc11.ChartChartTypeEnum.chChartTypeSmoothLine

            objChart3 = objCSpace3.Charts.Add(0)
            objChart3.Type = Microsoft.Office.Interop.Owc11.ChartChartTypeEnum.chChartTypeSmoothLine

            For j = 0 To timesUnit
                sum(j) = 0
            Next
            For i = 0 To YearNum
                rangStr = ""

                'milepost and crash rate
                For j = 0 To timesUnit
                    If j = 0 Then
                        rangStr = indi.MileFrom & vbTab
                        vv(i) &= 0 & vbTab 'crash rate
                        uu(i) &= 0 & vbTab 'Number of Crash
                    Else
                        'rangStr &= CStr(Format((  j * DivValue - (DivValue / 2)), "0.00")) & vbTab 'milepost
                        rangStr &= CStr(Format((indi.MileFrom + j * DivValue - (DivValue / 2)), "0.00")) & vbTab 'milepost

                        vv(i) &= CStr(results(i, j).accrate) & vbTab 'crash rate
                        uu(i) &= CStr(results(i, j).accnum) & vbTab  'Number of Crash
                    End If

                    sum(j) += results(i, j).accnum
                Next

                objChart1.SeriesCollection.Add(i)
                objChart2.SeriesCollection.Add(i)
            Next i
            xxsum = ""
            For j = 0 To timesUnit
                xxsum &= CStr(sum(j)) & vbTab
            Next


            With objChart1
                .HasLegend = True

                For i = 0 To YearNum
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimSeriesNames, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, indi.year + i)
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimValues, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, vv(i))
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimCategories, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, rangStr)
                Next i

                .SeriesCollection(0).Name = "=""Crash Rate"""

                .HasTitle = True
                If mutipleYear = True Then
                    chart_title = vbTab & "Crash Rate for year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & " Section No: " & indi.Csect & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo & "with every " & DivValue & "miles"
                Else
                    chart_title = vbTab & "Crash Rate for year " & indi.year & " on Hwy: " & indi.WayNum & " Section No: " & indi.Csect & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo & "with every " & DivValue & "miles"
                End If
                .Title.Caption = chart_title

                .Axes(0).HasTitle = True
                .Axes(1).HasTitle = True
                .Axes(0).Title.Caption = "MilePost"
                .Axes(1).Title.Caption = "Crash Rate"

            End With

            With objChart2
                .HasLegend = True

                For i = 0 To YearNum
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimSeriesNames, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, indi.year + i)
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimValues, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, uu(i))
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimCategories, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, rangStr)
                Next i

                .SeriesCollection(0).Name = "=""Number of Crashes"""

                .HasTitle = True
                If mutipleYear = True Then
                    chart_title = vbTab & "Number of crashes form year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & " Section No: " & indi.Csect & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo
                Else
                    chart_title = vbTab & "Number of crashes for year " & indi.year & " on Hwy: " & indi.WayNum & " Section No: " & indi.Csect & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo
                End If
                .Title.Caption = chart_title

                If mutipleYear = True Then
                    caption1 = "Number of crashes and crash rate form year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & ""
                Else
                    caption1 = "Number of crashes and crash rate for year " & indi.year & " on Hwy: " & indi.WayNum & ""
                End If
                caption = "MilePost: " & indi.MileFrom & "---" & indi.MileTo & "" & "" & " with every " & DivValue & " miles" & "(" & timesUnit & " parts)"
                If cmbInter.SelectedIndex = 2 Then
                    caption = caption & ", in Intersection selected"
                ElseIf cmbInter.SelectedIndex = 1 Then
                    caption = caption & " in Segmentation selected"
                Else
                    caption = caption & " in Total(Segmentation & Intersection) selected"
                End If

                .Axes(0).HasTitle = True
                .Axes(1).HasTitle = True
                .Axes(0).Title.Caption = "MilePost"
                .Axes(1).Title.Caption = "Number of Crashes"

            End With


            Dim total As String = "Total"
            With objChart3
                .HasLegend = True
                .SeriesCollection.Add()


                .SeriesCollection(0).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimSeriesNames, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, indi.year)
                .SeriesCollection(0).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimValues, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, xxsum)
                .SeriesCollection(0).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimCategories, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, rangStr)


                .SeriesCollection(0).Name = "=""Total Number of Crashes"""

                .HasTitle = True

                chart_title = vbTab & "Total Number of crashes form year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & " Section No: " & indi.Csect & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo

                .Title.Caption = chart_title


                'caption1 = "Total Number of crashes and crash rate form year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & ""

                caption = "MilePost: " & indi.MileFrom & "---" & indi.MileTo & "" & "" & " with every " & DivValue & " miles" & "(" & timesUnit & " parts)"
                If cmbInter.SelectedIndex = 2 Then
                    caption = caption & ", in Intersection selected"
                ElseIf cmbInter.SelectedIndex = 1 Then
                    caption = caption & " in Segmentation selected"
                Else
                    caption = caption & " in Total(Segmentation & Intersection) selected"
                End If

                .Axes(0).HasTitle = True
                .Axes(1).HasTitle = True
                .Axes(0).Title.Caption = "MilePost"
                .Axes(1).Title.Caption = "Number of Crashes"

            End With

            'Now a chart is ready to export to a GIF.
            Dim ChartName1 As String = Rnd() & ".gif"
            Dim strAbsolutePath As String = Server.MapPath(".") & "\" & ChartName1
            Dim strRelativePath1 As String = "./" & ChartName1
            objCSpace1.ExportPicture(strAbsolutePath, "GIF", 900, 450)
            Session("relPath") = strRelativePath1

            showlargechar = False 'don't used large chart
            Dim ChartName2 As String = Rnd() & ".gif"
            Dim strAbsolutePath2 As String = Server.MapPath(".") & "\" & ChartName2
            Dim strRelativePath2 As String = "./" & ChartName2
            objCSpace2.ExportPicture(strAbsolutePath2, "GIF", 900, 450)

            Dim ChartName3 As String = Rnd() & ".gif"
            Dim strAbsolutePath3 As String = Server.MapPath(".") & "\" & ChartName3
            Dim strRelativePath3 As String = "./" & ChartName3
            objCSpace3.ExportPicture(strAbsolutePath3, "GIF", 900, 450)
            Session("relPath3") = strRelativePath3

            Session("relPath") = strRelativePath1
            Session("relPath2") = strRelativePath2
            Session("relPath3") = strRelativePath3

Proc_Exit:
            rtValue = True
            objChart1 = Nothing
            Exit Sub
Proc_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error Number and Description' & Err.Num & Err.Description)"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            rtValue = False
            objChart1 = Nothing
            Exit Sub
        End Sub

        '************************
        Private Sub genCrashResults(ByVal year As Integer, ByVal num As Integer)
            Dim rowCount As Integer
            Dim colCount As Integer
            Dim strSource As String
            Dim strDestination As String

            rowCount = 3
            colCount = 5

            On Error GoTo Proc_Err
            'xlApp = CreateObject("Microsoft.Office.Interop.Excel.Application")
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application()
            xlApp.Visible = False
            strSource = AppPath & "\QryByComQy.xls"
            Dim ChartName As String = Rnd()
            strDestination = AppPath & "\tempByComQy " & ChartName & year & ".xls"
            Session(CStr(year) & "comqryname") = strDestination
            'strDestination = AppPath & "\tempByComQy" & year & ".xls"
            FileCopy(strSource, strDestination)
            xlBook = xlApp.Workbooks.Open(strDestination)
            xlSheet = xlBook.Worksheets(1)

            Call CrashQuery2(num, xlBook, xlSheet)

            '***           
            xlBook.Save()
            xlApp.ActiveWorkbook.Saved = True
            xlApp.DisplayAlerts = False
            xlApp.ActiveWorkbook.Close()
            'close the workbook
            xlApp.Workbooks.Close()
            xlApp.DisplayAlerts = True
            'close the Microsoft.Office.Interop.Excel
            xlApp.Quit()
            xlApp = Nothing
            xlSheet = Nothing
            'xlsheet3 = Nothing
            xlBook = Nothing
            GC.Collect()
Proc_Exit:
            Exit Sub
Proc_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error' & Err.Description)"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)


            If Not IsNothing(xlApp) Then
                xlApp.Quit()
            End If
            xlApp = Nothing
            xlSheet = Nothing
            xlBook = Nothing

            Resume Proc_Exit
        End Sub

        Private Sub drawCrashChart()
            Dim objCSpace As ChartSpace = New ChartSpaceClass()
            'Dim objChart As WCChart
            Dim objChart
            Dim hwy_class_name As String
            Dim chart_title As String
            Dim vv(YearNum), rangStr, midvalue(9) As String
            Dim i, j, k As Integer

            On Error GoTo Proc_Err

            caption = ""
            objChart = objCSpace.Charts.Add(0)
            objChart.Type = ChartChartTypeEnum.chChartTypeSmoothLine

            Dim n As Integer
            n = CInt((indi.MileTo - indi.MileFrom) / 10)
            rangStr = indi.MileFrom & vbTab
            For i = 1 To 9
                rangStr = rangStr & (n * i + indi.MileFrom) & vbTab
                midvalue(i) = "0"
            Next i
            rangStr = rangStr & indi.MileTo

            midvalue(0) = "0"

            For i = 0 To YearNum
                For j = 1 To 9
                    midvalue(j) = "0"
                Next
                For k = 0 To (rows(i) - 1)
                    If resultsCrash(i, k).accMilePost < n + indi.MileFrom Then
                        midvalue(1) = 1 + midvalue(1)
                    End If
                    If resultsCrash(i, k).accMilePost < 2 * n + indi.MileFrom Then
                        midvalue(2) = 1 + midvalue(2)
                    End If
                    If resultsCrash(i, k).accMilePost < 3 * n + indi.MileFrom Then
                        midvalue(3) = 1 + midvalue(3)
                    End If
                    If resultsCrash(i, k).accMilePost < 4 * n + indi.MileFrom Then
                        midvalue(4) = 1 + midvalue(4)
                    End If
                    If resultsCrash(i, k).accMilePost < 5 * n + indi.MileFrom Then
                        midvalue(5) = 1 + midvalue(5)
                    End If
                    If resultsCrash(i, k).accMilePost < 6 * n + indi.MileFrom Then
                        midvalue(6) = 1 + midvalue(6)
                    End If
                    If resultsCrash(i, k).accMilePost < 7 * n + indi.MileFrom Then
                        midvalue(7) = 1 + midvalue(7)
                    End If
                    If resultsCrash(i, k).accMilePost < 8 * n + indi.MileFrom Then
                        midvalue(8) = 1 + midvalue(8)
                    End If
                    If resultsCrash(i, k).accMilePost < 9 * n + indi.MileFrom Then
                        midvalue(9) = 1 + midvalue(9)
                    End If
                Next
                vv(i) = midvalue(0) & vbTab & midvalue(1) & vbTab & _
                    midvalue(2) & vbTab & midvalue(3) & vbTab & _
                    midvalue(4) & vbTab & midvalue(5) & vbTab & _
                    midvalue(6) & vbTab & midvalue(7) & vbTab & _
                    midvalue(8) & vbTab & midvalue(9) & vbTab & _
                    CStr(TotalResults)
                objChart.SeriesCollection.Add(i)
            Next i

            With objChart
                .HasLegend = True
                For i = 0 To YearNum
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimSeriesNames, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, indi.year + i)
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimValues, ChartSpecialDataSourcesEnum.chDataLiteral, vv(i))
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimCategories, ChartSpecialDataSourcesEnum.chDataLiteral, rangStr)
                Next i

                .SeriesCollection(0).Name = "=""Number of Crashes"""

                .HasTitle = True
                If mutipleYear = True Then
                    chart_title = vbTab & "Accumulated Number of Crashes for the year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo & vbLf
                Else
                    chart_title = vbTab & "Accumulated Number of Crashes for the year " & indi.year & " on Hwy: " & indi.WayNum & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo & vbLf
                End If

                caption1 = chart_title
                If cmbMonth.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Month: " & cmbMonth.SelectedItem.Text
                    caption = "--- Month: " & cmbMonth.SelectedItem.Text
                End If

                If cmbWeekday.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Weekday: " & cmbWeekday.SelectedItem.Text
                    caption = caption & " --- Weekday: " & cmbWeekday.SelectedItem.Text
                End If

                If cmbHourFrom.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Hour From: " & cmbHourFrom.SelectedItem.Text & ":00"
                    caption = caption & " --- Hour From: " & cmbHourFrom.SelectedItem.Text & ":00"
                Else
                    chart_title = chart_title & ", Hour From: " & "00" & ":00"
                    caption = caption & " --- Hour From: " & "00" & ":00"
                End If

                If cmbHourTo.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Hour To: " & cmbHourTo.SelectedItem.Text & vbLf & ":00"
                    caption = caption & " --- Hour To: " & cmbHourTo.SelectedItem.Text & vbLf & ":00"
                Else
                    chart_title = chart_title & ", Hour To: " & "23" & ":59"
                    caption = caption & " --- Hour To: " & "23" & ":59"
                End If

                If cmbWeather.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Weather Condition: " & cmbWeather.SelectedItem.Text
                    caption = caption & " --- Weather Condition: " & cmbWeather.SelectedItem.Text
                End If

                If Len(fstSpeed.Text) > 0 Then
                    chart_title = chart_title & ",  First Vehicle speed: " & fstSpeed.Text
                    caption = caption & " ---  First Vehicle speed: " & fstSpeed.Text
                End If

                If cmbViolation1.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Violation of Vehicle1: " & cmbViolation1.SelectedItem.Text
                    caption = caption & " --- Violation of Vehicle1: " & cmbViolation1.Items(cmbViolation1.SelectedIndex).Text
                End If

                If cmb1driver.SelectedIndex > 0 Then
                    chart_title = chart_title & ", First Driver's Condition: " & cmb1driver.SelectedItem.Text
                    caption = caption & " --- First Driver's Condition: " & cmb1driver.SelectedItem.Text
                End If

                If cmbRoadCond.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Road Condition: " & cmbRoadCond.SelectedItem.Text
                    caption = caption & " --- Road Condition: " & cmbRoadCond.SelectedItem.Text
                End If

                If cmbSufaceType.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Surface Condition: " & cmbSufaceType.SelectedItem.Text
                    caption = caption & " --- Surface Condition: " & cmbSufaceType.SelectedItem.Text
                End If

                If cmbSufaceCond.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Surface Condition: " & cmbSufaceCond.SelectedItem.Text
                    caption = caption & " --- Surface Condition: " & cmbSufaceCond.SelectedItem.Text
                End If

                If cmbRoadwayAlign.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Roadway Alignment: " & cmbRoadwayAlign.SelectedItem.Text
                    caption = caption & " ---  Roadway Alignment: " & cmbRoadwayAlign.SelectedItem.Text
                End If

                If cmbTypeofColl.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Type of Collision: " & cmbTypeofColl.SelectedItem.Text
                    caption = caption & " ---  Type of Collision: " & cmbTypeofColl.SelectedItem.Text
                End If

                If cmbTypeofAcc.SelectedIndex > 0 Then
                    chart_title = chart_title & ", Type of Accident: " & cmbTypeofAcc.SelectedItem.Text
                    caption = caption & " ---  Type of Accident: " & cmbTypeofAcc.SelectedItem.Text
                End If
                .Title.Caption = chart_title

                .Axes(0).HasTitle = True
                .Axes(1).HasTitle = True
                .Axes(0).Title.Caption = "MilePost"
                .Axes(1).Title.Caption = "Number of crashes"

            End With

            showlargechar = False 'don't used large chart
            Dim ChartName As String = Rnd() & ".gif"
            Dim strAbsolutePath As String = Server.MapPath(".") & "\" & ChartName
            objCSpace.ExportPicture(strAbsolutePath, "GIF", 900, 450)

            Dim strRelativePath As String = "./" & ChartName
            Session("relPath") = strRelativePath
            rtValue = True
Proc_Exit:
            objChart = Nothing
            Exit Sub
Proc_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error Number and Description' & Err.Num & Err.Description)"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            rtValue = False
            Resume Proc_Exit
        End Sub

        Public Sub CrashQuery2(ByVal num As Integer, ByVal xlBook As Microsoft.Office.Interop.Excel.Workbook, ByVal xlSheet As Microsoft.Office.Interop.Excel.Worksheet)
            'System table
            Dim accDB As String
            Dim accConn As ADODB.Connection
            Dim rst, rstcomputer1 As ADODB.Recordset
            'Dim t
            '
            Dim accRateTemp As Double
            Dim yearDate As String
            Dim rowcount As Integer
            Dim tempsql2 As String

            On Error GoTo Proc_Err

            Call generateCrashSql2()


            tempsql2 = condsql
            condsql = "SELECT [CSECT], [MILE_POST], count(*), " & _
                "SUM([NUM_KILLED]), SUM([NUM_INJ2]) , SUM([NUM_INJ3]) , SUM([NUM_INJ4]) , SUM([NUM_INJ5]) , SUM([NUM_INJ6]) " & _
                "FROM accidents" & indi.year + num & condsql & " GROUP BY [CSECT], [MILE_POST]" & _
                " ORDER BY [MILE_POST]"

            '****************************************
            accConn = New ADODB.Connection()
            rst = New ADODB.Recordset()

            accDB = "\" & statemdb & indi.year + num & ".mdb"
            With accConn
                'Telling ADO to use JOLT Here
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Open(AppPath & accDB)
            End With

            rowcount = 7
            TotalResults = 0
            rst = New ADODB.Recordset()
            rst.Open(condsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            With rst
                If .RecordCount < 1 Then
                    rtValue = False
                    GoTo Proc_Exit
                Else
                    xlSheet.Cells(6, 2) = indi.Hourfrom
                    rst.MoveFirst()
                    Do While (Not .EOF)
                        On Error GoTo Proc_Err
                        resultsCrash(num, rowcount - 7).accCsect = .Fields(0).Value  '[CSECT]  
                        resultsCrash(num, rowcount - 7).accMilePost = .Fields(1).Value    '[MilePost]  
                        resultsCrash(num, rowcount - 7).accNum = .Fields(2).Value    'count(*) 
                        resultsCrash(num, rowcount - 7).accFatalNum = .Fields(3).Value    'FatalInjuries
                        resultsCrash(num, rowcount - 7).accCriticalNum = .Fields(4).Value    'CriticalInjuries
                        resultsCrash(num, rowcount - 7).accSeriousNum = .Fields(5).Value     'SeriousInjuries
                        resultsCrash(num, rowcount - 7).accSevereNum = .Fields(6).Value     'SevereInjuries
                        resultsCrash(num, rowcount - 7).accModerateNum = .Fields(7).Value     'ModerateInjuries
                        resultsCrash(num, rowcount - 7).accMinorNum = .Fields(8).Value      'MinorInjuries
                        TotalResults = CInt(.Fields(2).Value) + TotalResults
                        rowcount = rowcount + 1
                        .MoveNext()
                    Loop
                    xlSheet.Cells(rowcount, 2) = indi.Hourto
                    xlSheet.Cells(rowcount, 4) = xlSheet.Cells(rowcount - 1, 4).value
                    rtValue = True
                End If
            End With

            'Sorting results by number of crashes
            Dim i, j As Integer
            Dim MidMilePost, ChangeMilePost, MidCsect, ChangeCsect As String
            Dim MidFatalNum, ChangeFatalNum, MidCriticalNum, ChangeCriticalNum, MidSeriousNum, ChangeSeriousNum As Integer
            Dim MidSevereNum, ChangeSevereNum, MidModerateNum, ChangeModerateNum, MidMinorNum, ChangeMinorNum As Integer
            Dim CompareSum, GetSum, MidSum, ChangeSum As Integer
            times = rowcount - 7

            Session("times") = times

            rows(num) = times 'store the number which will be used in drawing the chart
            For i = 0 To times - 2
                CompareSum = resultsCrash(num, 0).accNum
                MidCsect = resultsCrash(num, 0).accCsect
                MidMilePost = resultsCrash(num, 0).accMilePost
                MidFatalNum = resultsCrash(num, 0).accFatalNum
                MidCriticalNum = resultsCrash(num, 0).accCriticalNum
                MidSeriousNum = resultsCrash(num, 0).accSeriousNum
                MidSevereNum = resultsCrash(num, 0).accSevereNum
                MidModerateNum = resultsCrash(num, 0).accModerateNum
                MidMinorNum = resultsCrash(num, 0).accMinorNum

                For j = 1 To times - i
                    If CompareSum < resultsCrash(num, j).accNum Then
                        GetSum = resultsCrash(num, j).accNum
                        ChangeCsect = resultsCrash(num, j).accCsect
                        ChangeMilePost = resultsCrash(num, j).accMilePost
                        ChangeFatalNum = resultsCrash(num, j).accFatalNum
                        ChangeCriticalNum = resultsCrash(num, j).accCriticalNum
                        ChangeSeriousNum = resultsCrash(num, j).accSeriousNum
                        ChangeSevereNum = resultsCrash(num, j).accSevereNum
                        ChangeModerateNum = resultsCrash(num, j).accModerateNum
                        ChangeMinorNum = resultsCrash(num, j).accMinorNum

                        resultsCrash(num, j).accNum = CompareSum
                        resultsCrash(num, j).accCsect = MidCsect
                        resultsCrash(num, j).accMilePost = MidMilePost
                        resultsCrash(num, j).accFatalNum = MidFatalNum
                        resultsCrash(num, j).accCriticalNum = MidCriticalNum
                        resultsCrash(num, j).accSeriousNum = MidSeriousNum
                        resultsCrash(num, j).accSevereNum = MidSevereNum
                        resultsCrash(num, j).accModerateNum = MidModerateNum
                        resultsCrash(num, j).accMinorNum = MidMinorNum

                        resultsCrash(num, j - 1).accNum = GetSum
                        resultsCrash(num, j - 1).accCsect = ChangeCsect
                        resultsCrash(num, j - 1).accMilePost = ChangeMilePost
                        resultsCrash(num, j - 1).accFatalNum = ChangeFatalNum
                        resultsCrash(num, j - 1).accCriticalNum = ChangeCriticalNum
                        resultsCrash(num, j - 1).accSeriousNum = ChangeSeriousNum
                        resultsCrash(num, j - 1).accSevereNum = ChangeSevereNum
                        resultsCrash(num, j - 1).accModerateNum = ChangeModerateNum
                        resultsCrash(num, j - 1).accMinorNum = ChangeMinorNum
                    Else
                        CompareSum = resultsCrash(num, j).accNum
                        MidCsect = resultsCrash(num, j).accCsect
                        MidMilePost = resultsCrash(num, j).accMilePost
                        MidFatalNum = resultsCrash(num, j).accFatalNum
                        MidCriticalNum = resultsCrash(num, j).accCriticalNum
                        MidSeriousNum = resultsCrash(num, j).accSeriousNum
                        MidSevereNum = resultsCrash(num, j).accSevereNum
                        MidModerateNum = resultsCrash(num, j).accModerateNum
                        MidMinorNum = resultsCrash(num, j).accMinorNum
                    End If
                Next
            Next

            'add the decreasesing results of acc number to the excle table
            Dim tempsql As String

            rowcount = 2
            If times < 30 Then
                If times < 10 Then
                    HowManyResult(0) = times 'used in map show
                Else
                    HowManyResult(0) = 10 'used in map show
                End If
                For j = 0 To times - 1
                    xlSheet.Cells(rowcount, 1) = j + 1 'Ranked by Num of Acc
                    xlSheet.Cells(rowcount, 2) = resultsCrash(num, j).accCsect
                    xlSheet.Cells(rowcount, 3) = resultsCrash(num, j).accMilePost
                    xlSheet.Cells(rowcount, 4) = resultsCrash(num, j).accNum
                    xlSheet.Cells(rowcount, 5) = resultsCrash(num, j).accFatalNum
                    xlSheet.Cells(rowcount, 6) = resultsCrash(num, j).accCriticalNum
                    xlSheet.Cells(rowcount, 7) = resultsCrash(num, j).accSeriousNum
                    xlSheet.Cells(rowcount, 8) = resultsCrash(num, j).accSevereNum
                    xlSheet.Cells(rowcount, 9) = resultsCrash(num, j).accModerateNum
                    xlSheet.Cells(rowcount, 10) = resultsCrash(num, j).accMinorNum
                    rowcount = rowcount + 1
                    ''********                
                    ''get the value of accident's Computer to show acc map
                    'If (num = 0) Then 'now only show one years map'**************************************************************
                    '    rstcomputer1 = New ADODB.Recordset()
                    '    tempsql = "SELECT [COMPUTER] FROM accidents" & indi.year + num & tempsql2 & " AND [CSECT]='" & resultsCrash(num, j).accCsect & "' AND [MILE_POST]= " & _
                    '              resultsCrash(num, j).accMilePost
                    '    rstcomputer1.Open(tempsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    '    With rstcomputer1
                    '        .MoveFirst()
                    '        BlackNum(j).accComputer = "COMPUTER='" & .Fields(0).Value & "'"
                    '        If (j = 0) Then
                    '            queryString = BlackNum(0).accComputer
                    '        Else
                    '            queryString = queryString & " OR " & BlackNum(j).accComputer
                    '        End If
                    '    End With
                    '    rstcomputer1.Close()
                    '    rstcomputer1 = Nothing
                    'End If
                    ''get the value of accident's Computer to show acc map done               
                Next
            Else
                HowManyResult(0) = 10 'used in map show
                For j = 0 To 29
                    xlSheet.Cells(rowcount, 1) = j + 1 'Ranked by Num of Acc
                    xlSheet.Cells(rowcount, 2) = resultsCrash(num, j).accCsect
                    xlSheet.Cells(rowcount, 3) = resultsCrash(num, j).accMilePost
                    xlSheet.Cells(rowcount, 4) = resultsCrash(num, j).accNum
                    xlSheet.Cells(rowcount, 5) = resultsCrash(num, j).accFatalNum
                    xlSheet.Cells(rowcount, 6) = resultsCrash(num, j).accCriticalNum
                    xlSheet.Cells(rowcount, 7) = resultsCrash(num, j).accSeriousNum
                    xlSheet.Cells(rowcount, 8) = resultsCrash(num, j).accSevereNum
                    xlSheet.Cells(rowcount, 9) = resultsCrash(num, j).accModerateNum
                    xlSheet.Cells(rowcount, 10) = resultsCrash(num, j).accMinorNum
                    rowcount = rowcount + 1

                    ''get the value of accident's Computer to show acc map 
                    If (num = 0) Then 'now only show one years map'**************************************************************
                        rstcomputer1 = New ADODB.Recordset()
                        tempsql = "SELECT [COMPUTER] FROM accidents" & indi.year + num & tempsql2 & " AND [CSECT]='" & resultsCrash(num, j).accCsect & "' AND [MILE_POST]= " & _
                                  resultsCrash(num, j).accMilePost
                        rstcomputer1.Open(tempsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        With rstcomputer1
                            .MoveFirst()
                            BlackNum(j).accComputer = "COMPUTER='" & .Fields(0).Value & "'"
                            'If (j < 10) Then
                            '    If (j = 0) Then
                            '        queryString &= "(" & BlackNum(0).accComputer & ")"
                            '    Else
                            '        queryString &= " OR (" & BlackNum(j).accComputer & ")"
                            '    End If
                            'End If
                        End With
                        rstcomputer1.Close()
                        rstcomputer1 = Nothing
                    End If
                Next
            End If

            accConn.Close()
Proc_Exit:
            If Not (accConn Is Nothing) Then
                If (accConn.State And ConnectionState.Open) = ConnectionState.Open Then
                    accConn.Close()
                End If
                accConn = Nothing
            End If
            If Not (rst Is Nothing) Then
                If (rst.State And ConnectionState.Open) = ConnectionState.Open Then
                    rst.Close()
                End If
                rst = Nothing
            End If
            rstcomputer1 = Nothing
            Exit Sub
Proc_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error Number and Description')"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            Resume Proc_Exit
        End Sub

        '************************
        Private Sub generateCrashSql2()

            Dim endDate As Date

            condsql = " WHERE [HWY_NUM] = '" & indi.WayNum & "'"

            endDate = DateAdd("m", 1, accDate)



            condsql = condsql & _
                  generateFieldSql("WEATHER", indi.Weather) & _
                  generateFieldSql("SURF_TYPE", indi.SurfaceType) & _
                  generateFieldSql("SURF_COND", indi.SurfaceCond) & _
                  generateFieldSql("VIOLATION1", indi.Violation1) & _
                  generateFieldSql("ROAD_COND", indi.RoadCond) & _
                  generateFieldSql("TYPE_ACC", AccType) & _
                  generateFieldSql("VEH1_SPEED", indi.Veh1Speed) & _
                  generateFieldSql("COND_DRIV1", indi.Driver1Cond) & _
                  generateFieldSql("TYPE_COLL", indi.CollType) & _
                  generateFieldSql("ALIGNMENT", RoadAlgn)

            queryString = condsql
            'generateFieldSql("TYPE_COLL", indi.CollType) & _

            'Dim collsql As String = generateFieldSql("TYPE_COLL", indi.CollType)
            'Dim mapcollsql As String
            'condsql = condsql & collsql
            'If Not (indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003) Then
            '    mapcollsql = collsql
            'Else
            '    mapcollsql = generateFieldSql("MAN_COLL_C", indi.CollType)
            'End If

            'queryString = queryString & mapcollsql


            Dim weeksql As String = generateFieldSql("WEEKDAY", indi.WeekDay)
            condsql = condsql & weeksql
            If indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003 Then
                Dim wkDay As String = ""
                wkDay = changWeekday(indi.WeekDay)
                weeksql = generateFieldSql("WEEKDAY", wkDay)
            End If


            queryString = queryString & weeksql

            If Len(txtMileFrom.Text) > 0 And IsNumeric(txtMileFrom.Text) Then
                condsql = condsql & " AND ([MILE_POST] BETWEEN " & CDbl(indi.MileFrom)
                queryString = queryString & " AND ([MILE_POST] >=" & CDbl(indi.MileFrom)
            End If
            If Len(txtMileTo.Text) > 0 And IsNumeric(txtMileTo.Text) Then
                condsql = condsql & " AND " & CDbl(indi.MileTo) & " ) "
                queryString = queryString & " AND [MILE_POST] <" & CDbl(indi.MileTo) & " ) "

            End If


            If (Month <> "") Then
                condsql = condsql & "AND ([ACC_DATE] BETWEEN #" & Format(accDate, "Short Date") & _
                "# AND #" & Format(endDate, "Short Date") & "#)"
                queryString = queryString & "AND [ACC_DATE] >= date '" & Format(indi.Accdata, "Short Date") & "' AND [ACC_DATE] < date '" & Format(endDate, "Short Date") & "' "

            End If

            'check if it is a day or night
            If CInt(indi.Hourfrom) < CInt(indi.Hourto) Then
                '   condsql = condsql & " AND ([HOUR] BETWEEN '" & cmbHourFrom.Items(CInt(indi.Hourfrom) + 1).Text & "' AND '" & cmbHourTo.Items(CInt(indi.Hourto) + 1).Text & "')"

                condsql = condsql & " AND ([HOUR] BETWEEN '" & CInt(indi.Hourfrom) & "' AND '" & CInt(indi.Hourto) & "')"

                If indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003 Then
                    queryString &= " AND ([HOUR] >= " & CInt(indi.Hourfrom) & " AND [HOUR] <= " & CInt(indi.Hourto) & ")"
                Else : queryString &= " AND ([HOUR] >= '" & CStr(indi.Hourfrom) & "' AND [HOUR] <= '" & CStr(indi.Hourto) & "')"
                End If
            ElseIf CInt(indi.Hourfrom) = CInt(indi.Hourto) Then
                condsql = condsql & " AND [HOUR] = '" & CInt(indi.Hourfrom) & "'"
                If indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003 Then
                    queryString &= " AND [HOUR] = " & CInt(indi.Hourfrom)
                Else : queryString &= " AND [HOUR] = '" & CStr(indi.Hourfrom) & "'"
                End If
            End If


        End Sub

        Private Function GenerateSectReport(ByVal year As Integer, ByVal num As Integer)
            Dim strSource As String
            Dim strDestination As String

            On Error GoTo Proc_Err

            'xlApp = CreateObject("Microsoft.Office.Interop.Excel.Application")
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application()
            xlApp.Visible = False
            strSource = AppPath & "\QryByParHighWay.xls"
            Dim ChartName As String = Rnd()
            strDestination = AppPath & "\tempByParHW" & ChartName & year & ".xls"
            Session(CStr(year) & "byparhw") = strDestination

            FileCopy(strSource, strDestination)
            xlBook = xlApp.Workbooks.Open(strDestination)
            xlSheet = xlBook.Worksheets(1)
            xlsheet2 = xlBook.Worksheets.Add
            xlsheet3 = xlBook.Worksheets.Add

            With xlsheet2
                .Cells(1, 1).value = "Ranked by Crash Rate"
                .Cells(1, 2).value = "Control Section"
                .Cells(1, 3).value = "Crash Rate"
                .Cells(1, 4).value = "Number of Crashes"
                .Cells(1, 5).value = "Mile from"
                .Cells(1, 6).value = "Mile to"
            End With
            With xlsheet3
                .Cells(1, 1).value = "Ranked by Num of Acc"
                .Cells(1, 2).value = "Control Section"
                .Cells(1, 3).value = "Number of Crashes"
                .Cells(1, 4).value = "Crash Rate"
                .Cells(1, 5).value = "Mile from"
                .Cells(1, 6).value = "Mile to"
            End With

            Call CompoundSectQuery(num, xlBook, xlSheet, xlsheet2, xlsheet3)

            xlBook.Save()
            xlApp.ActiveWorkbook.Saved = True
            xlApp.DisplayAlerts = False
            xlApp.ActiveWorkbook.Close()

            'close the workbook
            xlApp.Workbooks.Close()
            xlApp.DisplayAlerts = True

            'close the Microsoft.Office.Interop.Excel
            xlApp.Quit()
            xlApp = Nothing
            xlSheet = Nothing
            xlsheet2 = Nothing
            xlsheet3 = Nothing
            xlBook = Nothing
            GC.Collect()

Proc_Exit:
            'MousePointer = vbNormal
            Exit Function
Proc_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error" & Err.Description & "')"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)

            If Not IsNothing(xlApp) Then
                xlApp.Quit()
            End If
            xlApp = Nothing
            xlSheet = Nothing
            xlsheet2 = Nothing
            xlsheet3 = Nothing
            xlBook = Nothing

            Resume Proc_Exit
        End Function

        Public Sub CompoundSectQuery(ByVal num As Integer, ByVal xlBook As Microsoft.Office.Interop.Excel.Workbook, ByVal xlSheet As Microsoft.Office.Interop.Excel.Worksheet, ByVal xlSheet2 As Microsoft.Office.Interop.Excel.Worksheet, ByVal xlSheet3 As Microsoft.Office.Interop.Excel.Worksheet)
            'System table
            Dim i, j As Integer

            Dim accConn As ADODB.Connection
            Dim rst As ADODB.Recordset
            Dim accRateTemp As Double

            On Error GoTo Proc_Err

            accConn = New ADODB.Connection()
            rst = New ADODB.Recordset()
            accConn = New ADODB.Connection()
            Dim Connect As String
            Connect = "Driver={Microsoft dBase Driver (*.dbf)};" & _
             "Dbq=" & AppPath & "\GisPrj;" & _
             "DefaultDir=" & AppPath & "\GisPrj;" & _
             "Uid=Admin;Pwd=;"
            accConn.Open(Connect)

            Dim mileType_fr As String
            Dim mileType_to As String

            mileType_fr = "MIPOST_FR"
            mileType_to = "MIPOST_TO"

            condsql = " WHERE [HWY_NUM] = '" & indi.WayNum & "'"

            '*****************************Maybe Have Problem
            condsql = condsql & " AND ((" & mileType_fr & " >= " & indi.MileFrom & _
     " AND " & mileType_fr & " <= " & indi.MileTo & ") OR (" & mileType_to & " > " & indi.MileFrom & _
     " AND " & mileType_to & " <= " & indi.MileTo & "))"

            '
            condsql = "SELECT [CSECT], [MIPOST_FR], [TOT_PERMVM], SUM([TOT_ACC]),[MIPOST_TO]" & _
       "FROM sec" & indi.year + num & ".dbf" & condsql & " GROUP BY [Csect], [MIPOST_FR], [TOT_PERMVM],[MIPOST_TO]" & _
     " ORDER BY [MIPOST_FR]"
            rst = New ADODB.Recordset()
            rst.Open(condsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            Dim rowCount = 7, row = 7
            Dim midvalue As Integer
            With rst
                If .RecordCount < 1 Then
                    rtValue = False
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('No records found for the selections made.')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    GoTo Proc_Exit
                Else
                    'xlSheet.Cells(6, 2) = Qry2.From
                    rtValue = True
                    rst.MoveFirst()
                    Do While (Not .EOF)
                        On Error GoTo Proc_Err
                        resultsSect(num, rowCount - 7).accCsect = .Fields(0).Value '[CSECT]
                        resultsSect(num, rowCount - 7).accFrom = .Fields(1).Value '[MIPOST_FR]'milepost from
                        resultsSect(num, rowCount - 7).accTo = .Fields(4).Value   '[MIPOST_TO]'milepost to
                        resultsSect(num, rowCount - 7).accSum = .Fields(3).Value    'SUM([TOT_ACC])
                        resultsSect(num, rowCount - 7).accRate = .Fields(2).Value   '[TOT_PERMVM],crash_rate                 

                        rowCount = rowCount + 1
                        .MoveNext()
                    Loop
                    xlSheet.Cells(rowCount, 4) = xlSheet.Cells(rowCount - 1, 4)
                End If
            End With

            Dim accInfo(rowCount - 7) As AccInfoSect
            For i = 0 To rowCount - 7
                accInfo(i).accCsect = resultsSect(num, i).accCsect
                accInfo(i).accFrom = resultsSect(num, i).accFrom
                accInfo(i).accRate = resultsSect(num, i).accRate
                accInfo(i).accSum = resultsSect(num, i).accSum
                accInfo(i).accTo = resultsSect(num, i).accTo
            Next

            'Sorting results by Number of Crash
            Dim CompareRate, GetRate, MidRate, ChangeRate, MidFrom, MidTo, ChangeFrom, ChangeTo As Double
            Dim MidCsect, ChangeCsect As String
            Dim CompareSum, GetSum, MidSum, ChangeSum, timesSect As Integer
            timesSect = rowCount - 7
            rows(num) = rowCount 'store the number which will be used in drawing the chart
            For i = 0 To timesSect - 2
                CompareSum = accInfo(0).accSum
                MidCsect = accInfo(0).accCsect
                MidFrom = accInfo(0).accFrom
                MidTo = accInfo(0).accTo
                MidRate = accInfo(0).accRate
                For j = 1 To timesSect - i
                    If CompareSum < accInfo(j).accSum Then
                        GetSum = accInfo(j).accSum
                        ChangeCsect = accInfo(j).accCsect
                        ChangeRate = accInfo(j).accRate
                        ChangeFrom = accInfo(j).accFrom
                        ChangeTo = accInfo(j).accTo

                        accInfo(j).accSum = CompareSum
                        accInfo(j).accCsect = MidCsect
                        accInfo(j).accRate = MidRate
                        accInfo(j).accFrom = MidFrom
                        accInfo(j).accTo = MidTo

                        accInfo(j - 1).accSum = GetSum
                        accInfo(j - 1).accCsect = ChangeCsect
                        accInfo(j - 1).accRate = ChangeRate
                        accInfo(j - 1).accFrom = ChangeFrom
                        accInfo(j - 1).accTo = ChangeTo
                    Else
                        CompareSum = accInfo(j).accSum
                        MidCsect = accInfo(j).accCsect
                        MidFrom = accInfo(j).accFrom
                        MidTo = accInfo(j).accTo
                        MidRate = accInfo(j).accRate
                    End If
                Next
            Next
            'add the decreasesing results of acc number to the excle table
            rowCount = 3
            HowManyResult(0) = 0
            For i = 0 To timesSect - 1
                xlSheet3.Cells(rowCount, 1) = i + 1 'Ranked by Num of Acc
                xlSheet3.Cells(rowCount, 2) = accInfo(i).accCsect  'Csect
                xlSheet3.Cells(rowCount, 3) = accInfo(i).accSum  'number
                xlSheet3.Cells(rowCount, 4) = accInfo(i).accRate 'crash rate
                xlSheet3.Cells(rowCount, 5) = accInfo(i).accFrom  'from
                xlSheet3.Cells(rowCount, 6) = accInfo(i).accTo   'to
                rowCount = rowCount + 1
                '
                If accInfo(i).accSum <> 0 Then 'be used in map show
                    BlackNum(i).Csect = accInfo(i).accCsect 'total can save 800 items
                    'BlackNum(i).accLogmifrom = accInfo(i).accFrom
                    'BlackNum(i).accLogmito = accInfo(i).accTo
                End If
                '
                If HowManyResult(0) = 0 And accInfo(i).accSum = 0 Then 'used in map show
                    HowManyResult(0) = i
                    If i < 19 Then
                        queryString = queryString & " (CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "' AND MIPOST_FR=" & accInfo(i).accFrom & " AND MIPOST_TO=" & accInfo(i).accTo & ")"
                    End If
                End If
                If HowManyResult(0) = 0 And i < 20 Then
                    'If indi.year = 2000 Then

                    '    'used in map showing'search the accidents database
                    '    If i = 0 Then
                    '        'queryString = "WHERE (CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "' AND MIPOST_FR=" & accInfo(i).accFrom & " AND MIPOST_TO=" & accInfo(i).accTo & ") OR"
                    '        queryString = "WHERE CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "' OR "
                    '    ElseIf (i = timesSect - 1) Or (i = 19) Then
                    '        'queryString = queryString & " (CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "' AND MIPOST_FR=" & accInfo(i).accFrom & " AND MIPOST_TO=" & accInfo(i).accTo & ")"
                    '        queryString = queryString & " CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "'"
                    '    Else
                    '        'queryString = queryString & "(CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "' AND MIPOST_FR=" & accInfo(i).accFrom & " AND MIPOST_TO=" & accInfo(i).accTo & ") OR"
                    '        queryString = queryString & "CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "' OR "
                    '    End If


                    If i = 0 Then
                        'queryString = "WHERE (CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "' AND MIPOST_FR=" & accInfo(i).accFrom & " AND MIPOST_TO=" & accInfo(i).accTo & ") OR"
                        queryString = "WHERE CSECT='" & accInfo(i).accCsect & "' OR "
                    ElseIf (i = timesSect - 1) Or (i = 19) Then
                        'queryString = queryString & " (CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "' AND MIPOST_FR=" & accInfo(i).accFrom & " AND MIPOST_TO=" & accInfo(i).accTo & ")"
                        queryString = queryString & " CSECT='" & accInfo(i).accCsect & "'"
                    Else
                        'queryString = queryString & "(CSECT='" & accInfo(i).accCsect.Insert(3, "-") & "' AND MIPOST_FR=" & accInfo(i).accFrom & " AND MIPOST_TO=" & accInfo(i).accTo & ") OR"
                        queryString = queryString & "CSECT='" & accInfo(i).accCsect & "' OR "
                    End If


                End If
            Next

            If HowManyResult(0) = 0 Then
                HowManyResult(0) = rowCount - 7 'used in map show
            End If




            'sorting the result by acc rate
            For i = 0 To timesSect - 2
                CompareRate = accInfo(0).accRate
                MidCsect = accInfo(0).accCsect
                MidFrom = accInfo(0).accFrom
                MidTo = accInfo(0).accTo
                MidSum = accInfo(0).accSum
                For j = 1 To timesSect - i
                    If CompareRate < accInfo(j).accRate Then
                        GetRate = accInfo(j).accRate
                        ChangeCsect = accInfo(j).accCsect
                        ChangeSum = accInfo(j).accSum
                        ChangeFrom = accInfo(j).accFrom
                        ChangeTo = accInfo(j).accTo

                        accInfo(j).accRate = CompareRate
                        accInfo(j).accSum = MidSum
                        accInfo(j).accCsect = MidCsect
                        accInfo(j).accFrom = MidFrom
                        accInfo(j).accTo = MidTo

                        accInfo(j - 1).accRate = GetRate
                        accInfo(j - 1).accSum = ChangeSum
                        accInfo(j - 1).accCsect = ChangeCsect
                        accInfo(j - 1).accFrom = ChangeFrom
                        accInfo(j - 1).accTo = ChangeTo
                    Else
                        CompareRate = accInfo(j).accRate
                        MidCsect = accInfo(j).accCsect
                        MidFrom = accInfo(j).accFrom
                        MidTo = accInfo(j).accTo
                        MidSum = accInfo(j).accSum
                    End If
                Next
            Next
            'add the decreasesing results of acc rate to the excle table
            rowCount = 3
            For i = 0 To timesSect - 1
                xlSheet2.Cells(rowCount, 1) = i + 1 'Ranked by Num of Acc
                xlSheet2.Cells(rowCount, 2) = accInfo(i).accCsect  'Csect
                xlSheet2.Cells(rowCount, 3) = accInfo(i).accRate  'crash rate
                xlSheet2.Cells(rowCount, 4) = accInfo(i).accSum 'number
                xlSheet2.Cells(rowCount, 5) = accInfo(i).accFrom  'from
                xlSheet2.Cells(rowCount, 6) = accInfo(i).accTo   'to
                rowCount = rowCount + 1
                '
                If accInfo(i).accSum <> 0 Then 'be used in map show
                    BlackRate(i).Csect = accInfo(i).accCsect 'total can save 800 items
                    'BlackRate(i).accLogmifrom = accInfo(i).accFrom
                    'BlackRate(i).accLogmito = accInfo(i).accTo
                End If
            Next

            accConn.Close()
            rst = Nothing
Proc_Exit:
            If Not (accConn Is Nothing) Then
                If (accConn.State And ConnectionState.Open) = ConnectionState.Open Then
                    accConn.Close()
                End If
                accConn = Nothing
            End If

            rst = Nothing
            Exit Sub
Proc_Err:
            strscript = "<script language=javascript>alert('Error" & Err.Description & "');</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "clientscript", strscript)
            rtValue = False
            Resume Proc_Exit
        End Sub

        Private Sub drawSectChart()
            Dim objCSpace As Microsoft.Office.Interop.Owc11.ChartSpace = New Microsoft.Office.Interop.Owc11.ChartSpaceClass()
            Dim objCSpace2 As Microsoft.Office.Interop.Owc11.ChartSpace = New Microsoft.Office.Interop.Owc11.ChartSpaceClass()
            Dim objChart, objChart2
            Dim hwy_class_name As String
            Dim chart_title As String
            Dim vv(YearNum), uu(YearNum), rangStr As String
            Dim i, j, k As Integer

            On Error GoTo Proc_Err

            caption = ""
            objChart = objCSpace.Charts.Add(0)
            objChart.Type = Microsoft.Office.Interop.Owc11.ChartChartTypeEnum.chChartTypeSmoothLine

            objChart2 = objCSpace2.Charts.Add(0)
            objChart2.Type = Microsoft.Office.Interop.Owc11.ChartChartTypeEnum.chChartTypeSmoothLine

            For i = 0 To YearNum
                rangStr = ""
                'milepost and crash rate
                For k = 7 To (rows(i) - 1) Step 3
                    If (k > rows(i) - 1) Then
                        k = rows(i) - 1
                    End If
                    rangStr &= resultsSect(i, k - 7).accFrom & vbTab 'milepost
                    vv(i) &= resultsSect(i, k - 7).accRate & vbTab 'crash rate
                    uu(i) &= CStr(resultsSect(i, k - 7).accSum) & vbTab  'Number of Crash
                Next
                objChart.SeriesCollection.Add(i)
                objChart2.SeriesCollection.Add(i)
            Next i

            With objChart
                .HasLegend = True

                'some wrong in here **********************************************************
                For i = 0 To YearNum
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimSeriesNames, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, indi.year + i)
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimValues, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, vv(i))
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimCategories, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, rangStr)
                Next i

                .SeriesCollection(0).Name = "=""Crash Rate"""

                .HasTitle = True
                If mutipleYear = True Then
                    chart_title = vbTab & "Crash Rate for year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo
                Else
                    chart_title = vbTab & "Crash Rate for year " & indi.year & " on Hwy: " & indi.WayNum & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo
                End If
                .Title.Caption = chart_title

                If mutipleYear = True Then
                    caption1 = "Crash Rate for year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & ""
                Else
                    caption1 = "Crash Rate for year " & indi.year & " on Hwy: " & indi.WayNum & ""
                End If
                caption = "MilePost: " & indi.MileFrom & "---" & indi.MileTo & ""

                If mutipleYear = True Then
                    newcaption = "Top Ranked Crash Rate and Number of Crashes Form year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & " " & caption
                Else
                    newcaption = "Top Ranked Crash Rate and Number of Crahses in year " & indi.year & " on Hwy: " & indi.WayNum & " " & caption
                End If

                Session("newcaption") = newcaption


                .Axes(0).HasTitle = True
                .Axes(1).HasTitle = True
                .Axes(0).Title.Caption = "MilePost"
                .Axes(1).Title.Caption = "Crash Rate"

            End With

            With objChart2
                .HasLegend = True

                'some wrong in here **********************************************************
                For i = 0 To YearNum
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimSeriesNames, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, indi.year + i)
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimValues, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, uu(i))
                    .SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimCategories, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, rangStr)
                Next i

                .SeriesCollection(0).Name = "=""Number of Crash"""

                .HasTitle = True
                If mutipleYear = True Then
                    chart_title = vbTab & "Number of crash form year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo
                Else
                    chart_title = vbTab & "Number of crash for year " & indi.year & " on Hwy: " & indi.WayNum & " between MilePost: " & indi.MileFrom & " and " & indi.MileTo
                End If
                .Title.Caption = chart_title

                If mutipleYear = True Then
                    caption1 = "Number of crashes and crash rate for year " & indi.year & " to " & indi.year + YearNum & " on Hwy: " & indi.WayNum & ""
                Else
                    caption1 = "Number of crashes and crash rate for year " & indi.year & " on Hwy: " & indi.WayNum & ""
                End If
                caption = "MilePost: " & indi.MileFrom & "---" & indi.MileTo & ""

                .Axes(0).HasTitle = True
                .Axes(1).HasTitle = True
                .Axes(0).Title.Caption = "MilePost"
                .Axes(1).Title.Caption = "Number of crash"

            End With

            'Now a chart is ready to export to a GIF.
            Dim ChartName1 As String = Rnd() & ".gif"
            Dim strAbsolutePath As String = Server.MapPath(".") & "\" & ChartName1
            Dim strRelativePath1 As String = "./" & ChartName1
            objCSpace.ExportPicture(strAbsolutePath, "GIF", 900, 450)
            Session("relPath") = strRelativePath1

            showlargechar = False 'don't used large chart
            Dim ChartName2 As String = Rnd() & ".gif"
            Dim strAbsolutePath2 As String = Server.MapPath(".") & "\" & ChartName2
            Dim strRelativePath2 As String = "./" & ChartName2
            objCSpace2.ExportPicture(strAbsolutePath2, "GIF", 900, 450)

            Session("relPath") = strRelativePath1
            Session("relPath2") = strRelativePath2

Proc_Exit:
            rtValue = True
            objChart = Nothing
            Exit Sub
Proc_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error Number and Description' & Err.Num & Err.Description)"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            Resume Proc_Exit
        End Sub

        Private Sub CancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelButton.Click
            Response.Redirect("options.aspx")
        End Sub

        Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
            Dim b As String
            b = Button2.Text
            If b = ">" Then
                'check the highwaynumber
                If txtHwyNum.Text = "" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('Please enter highway number.')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                End If

                ctrlSect.Visible = True
                ctrlSect.Enabled = True
                ctrlSect1.Visible = True
                ctrlSect1.Enabled = True
                Button2.Text = "<"
                Button3.Visible = True

                'select control section for highway
                Dim accConn As ADODB.Connection
                Dim rst As ADODB.Recordset
                Dim accDB As String
                Dim condsql As String

                'Dim myConnection As OleDbConnection
                'Dim myDa1 As OleDbDataAdapter
                'Dim ds As DataSet
                'Dim sql As String

                'myConnection = New OleDbConnection("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA Source=C:\Data\section2004.mdb")
                'sql = "Select Distinct CSECT From section2004 Where "
                'myDa1 = New OleDbDataAdapter(sql, myConnection)
                'ds = New DataSet()
                'myDa1.Fill(ds, "AllTables")
                'ctrlSect.DataSource = ds
                'ctrlSect.DataSource = ds.Tables(0)
                'ctrlSect.DataTextField = ds.Tables(0).Columns("CSECT").ColumnName.ToString()
                'ctrlSect.DataValueField = ds.Tables(0).Columns("CSECT").ColumnName.ToString()
                'ctrlSect.DataBind()
                '
                ctrlSect.Items.Clear()
                ctrlSect.Items.Add("")

                accConn = New ADODB.Connection()
                'rst = New ADODB.Recordset
                accDB = "C:\DATA\mdbdata\section2004.mdb"
                With accConn
                    'Telling ADO to use JOLT Here
                    .Provider = "Microsoft.Jet.OLEDB.4.0"
                    .Open(accDB)
                End With

                condsql = "SELECT DISTINCT[CSECT] FROM section2004 WHERE HWY_NUM='" & GetTrueHnum(txtHwyNum.Text) & "'"
                rst = New ADODB.Recordset()
                rst.Open(condsql, accConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                With rst
                    If .RecordCount < 1 Then
                        Exit Sub
                    Else
                        rst.MoveFirst()
                        Do While (Not .EOF)
                            ctrlSect.Items.Add(.Fields(0).Value)
                            .MoveNext()
                        Loop
                    End If
                End With

                accConn.Close()
                rst = Nothing
            End If
            If b = "<" Then
                ctrlSect.Visible = False
                ctrlSect.Enabled = False
                ctrlSect1.Visible = False
                ctrlSect1.Enabled = False
                Button2.Text = ">"
                Button3.Visible = False
            End If
        End Sub

        Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

            Response.Write("<script language ='javascript'> window.open('HelpCSECTwin.aspx?','form','width=420,height=500,left=800,top=10,scrollbars = yes');</script>")

            'Dim conAuthors As OleDbConnection
            'Dim strSelect As String
            'Dim strSelect1 As String
            'Dim cmdSelect As OleDbCommand
            'Dim cmdSelect1 As OleDbCommand

            'conAuthors = New OleDbConnection("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA Source=C:\Data\scetion2004.mdb")
            '''''''strSelect = "Select MILEPOINT, LOGMIFROM,LOGMITO From scetion2004 Where CSECT=?   "
            ''''''LOOP' SEARCH EACH RESULTS THAT WE FOUND
            ''''''  FIND THE MINIMUM LOFGIFROM AND MAXIMU  LOGMITO
            ''''''MAXIMU(LOGMITO = LENTH)
            ''''''LENTH +

            ''''''END LOOP


            'cmdSelect = New OleDbCommand(strSelect, conAuthors)
            'strSelect1 = "Select CSECT From section2004"
            'cmdSelect = New OleDbCommand(strSelect1, conAuthors)

            'cmdSelect.Parameters.Add("@MilePost", txtMileTo.Text)
            'cmdSelect1.Parameters.Add("@CtrlSect", ctrlSect.SelectedItem.Value)
        End Sub


      


    End Class

End Namespace
