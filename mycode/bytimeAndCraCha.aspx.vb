Imports System.Data.OleDb
Imports microsoft.office.interop.owc11


Namespace Crashsafe


    Partial Class bytimeAndCraCha
        Inherits System.Web.UI.Page

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
''global variable
		Dim statemdb As String = "mdbdata\acc"
		Dim strscript As String
		Dim optDistChosen As Boolean
		Dim optstateChosen As Boolean
		Dim indi As StrIndividual
		Dim caption As String
		Dim caption1 As String
		Dim caption2 As String
		Dim queryString As String
		Dim distNo As String
		Dim mutipleYear As Boolean
		Dim YearNum As Integer
		Dim NoShow As Boolean
		Dim connString As String
		Dim showlargechar As Boolean

		Dim QueryType As Integer
		Dim SubQueryType As Integer
		
		'local variables

		Dim ShowAcc As Boolean
		Dim byMonthChosen As Boolean
		Dim byHourChosen As Boolean
		Dim byDayChosen As Boolean
		Dim byCollChosen As Boolean
		Dim byAccChosen As Boolean
		Dim byPt_ImpactChosen As Boolean
		Dim byViolChosen As Boolean
		Dim HigNum As String

		Dim Myconnection As New OleDbConnection()
		Public mdbtable As String = "accidents"

		''end of global variable
        Dim cstype As Type = Me.GetType()
        Private condsql As String
        Private indiTimeAndCraCha As StrIndividual

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
            ''  Put user code to initialize the page here
            'Dim login As String
            'login = Session("succ")
            'If login = "" Then ' fake user
            '    Response.Redirect("LoginIn.aspx")
            'End If

            'If RightLogin = 0 Then 'if no login, don't allow to use this webpage
            '    Response.Redirect("LoginIn.aspx")
            'End If

			getSession()

		 '1=By time & By Crash Characteristics,3=blackspot, 4=by partial highway,5=compound query

            If optDistChosen Then 'by District
				Label4.Text = "District " & distNo & " query"
				
				cmbYear.Items.Remove("1999")
				cmbYear2.Items.Remove("1999")

            Else
                Label4.Text = "whole state query"
			End If



        End Sub


        Private Sub cmdMonth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMonth.Click

            If cmbYear.SelectedIndex = 0 Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('You must select one year for your query!')"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

            If cmbYear2.Enabled = True Then
                If CInt(cmbYear.SelectedIndex.ToString) >= CInt(cmbYear2.SelectedIndex.ToString) And cmbYear2.SelectedIndex.ToString <> "0" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The YearFrom can not be greater than or equal to YearTo!')"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                End If
                'mutipleYear = True
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
            '
            indi.year = cmbYear.SelectedItem.Text
            'indi.year = cmbYear.SelectedItem.Text'***???
            byMonthChosen = True
            byHourChosen = False
            byDayChosen = False
            Call genTimeResults()
            QueryType = 1 'by time
            SubQueryType = 1 'by month
			NoShow = True

			saveintoSession()

            Response.Write("<script language='javascript'>window.open('rep1Month.aspx');</Script>")

        End Sub


        Private Sub cmdDay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDay.Click
      
            If cmbYear.SelectedIndex = 0 Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('You must select one year for your query!')"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

            If cmbYear2.Enabled = True Then
                If CInt(cmbYear.SelectedIndex.ToString) >= CInt(cmbYear2.SelectedIndex.ToString) And cmbYear2.SelectedIndex.ToString <> "0" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The YearFrom can not be greater than or equal to YearTo!')"
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
            '
            indi.year = cmbYear.SelectedItem.Text
            byDayChosen = True
            byHourChosen = False
            byMonthChosen = False
            Call genTimeResults()
            QueryType = 1 'by time
            SubQueryType = 2 'by day
			NoShow = True

	saveintoSession()

            Response.Write("<script language='javascript'>window.open('rep1Day.aspx');</Script>")

        End Sub

        Private Sub cmdHour_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHour.Click


            If cmbYear.SelectedIndex = 0 Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('You must select one year for your query!')"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

            If cmbYear2.Enabled = True Then
                If CInt(cmbYear.SelectedIndex.ToString) > CInt(cmbYear2.SelectedIndex.ToString) And cmbYear2.SelectedIndex.ToString <> "0" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The YearFrom can not be greater than or equal to YearTo!')"
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
            '
            indi.year = cmbYear.SelectedItem.Text
            byHourChosen = True
            byDayChosen = False
            byMonthChosen = False
            Call genTimeResults()
            QueryType = 1 'by time
            SubQueryType = 3 'by hour
			NoShow = True
				saveintoSession()
            Response.Write("<script language='javascript'>window.open('rep1Hour.aspx');</Script>")
        End Sub

        Private Sub genTimeResults()
            Dim tempSql As String
            Dim strValues(YearNum) As String 'used for chart created(value)
            Dim strCategory As String
            'Dim i As Integer
            Dim myReader As OleDbDataReader

            Dim cmd As OleDbCommand
            Dim row As Integer = 1
            Dim j As Integer
            Dim timeAdapter As OleDbDataAdapter
            Dim timeDs As DataSet
            timeDs = New DataSet()

            Call generateSql()

            If byMonthChosen Then
                Call genStrCat4Month(strCategory)
            ElseIf byDayChosen Then
                Call genStrCat4Day(strCategory)
            ElseIf byHourChosen Then
                Call genStrCat4Hour(strCategory)
            End If

            Dim cycle As Integer
            For cycle = 0 To YearNum
                If optDistChosen Then 'by District
                    '  connString = "Provider=Microsoft.Jet.OLEDB.4.0; User Id=; Password=; Data Source=C:\Data\" & statemdb & indi.year + cycle & "Dist" & distNo & ".mdb"
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\data\GisPrj\District; Extended Properties = DBASE III;"

                Else 'by State
                    ' connString = "Provider=Microsoft.Jet.OLEDB.4.0; User Id=; Password=; Data Source= C:\Data\accidents" & indi.year + cycle & ".mdb"
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;  Data Source= C:\data\" & statemdb & indi.year + cycle & ".mdb"

                End If

                'get the sql of check for every year
                If byMonthChosen Then
                    Call sqlStmt4Month(tempSql, indi.year + cycle)
                ElseIf byDayChosen Then
                    Call sqlStmt4Day(tempSql, indi.year + cycle)
                ElseIf byHourChosen Then
                    Call sqlStmt4Hour(tempSql, indi.year + cycle)
                End If

                If optDistChosen Then
                    tempSql = ChangeExp(tempSql, indi.year)
                End If


                Myconnection = Nothing
                Myconnection = New OleDbConnection(connString)
                myReader = Nothing
                cmd = Nothing
                timeAdapter = Nothing

                openDbConn()
                cmd = New OleDbCommand(tempSql, Myconnection)
                timeAdapter = New OleDbDataAdapter(tempSql, Myconnection)
                myReader = cmd.ExecuteReader()

                Dim monthDate As String
                Dim temp As Integer
                Dim rowCount As Integer
                Dim checkRowCount As Integer
                rowCount = 0
                While myReader.Read
                    rowCount = rowCount + 1
                End While
                myReader.Close()
                myReader = cmd.ExecuteReader()
                Try
                    If byMonthChosen Then
                        myReader.Read()
                        checkRowCount = 0
                        For j = 1 To 12
                            If checkRowCount = rowCount Then Exit For
                            Dim tempdate As Date
                            tempdate = DateSerial(CInt(cmbYear.SelectedItem.Text) + cycle, j, 1)
                            monthDate = Format(tempdate, "MM/yyyy")
                            If monthDate = myReader.GetValue(0).ToString() Then
                                strValues(cycle) += myReader.GetValue(1).ToString() + vbTab
                                checkRowCount += 1
                                myReader.Read()
                            Else
                                strValues(cycle) += "0" + vbTab
                            End If
                        Next
                        If cycle = YearNum Then
                            draw2TimeChart(strValues, strCategory, cycle)
                        End If
                    ElseIf byDayChosen Then
                        temp = 1
                        myReader.Read()
                        checkRowCount = 0
                        For j = 1 To 7
                            If checkRowCount = rowCount Then Exit For
							Dim wkday As String = myReader.GetValue(0).ToString

						   If optDistChosen Then

								 If indi.year = "2001" Or indi.year = "2002" Or indi.year = "2003" Then
								   wkday = changWeekdayback(wkday)
							  End If
						 End If

						While Not (temp = CInt(wkday))
								myReader.Read()
								If optDistChosen And (indi.year = "2001" Or indi.year = "2002" Or indi.year = "2003") Then
									  wkday = changWeekdayback(wkday)
								End If
						End While

						'	If temp = CInt(wkday) Then
								strValues(cycle) += myReader.GetValue(1).ToString() + vbTab
								checkRowCount += 1
								myReader.Read()
						'	Else
							'	strValues(cycle) += "0" + vbTab
							'End If
							temp += 1

						Next

                        If cycle = YearNum Then
                            draw2TimeChart(strValues, strCategory, cycle)
                        End If
                    ElseIf byHourChosen Then
                        Dim temp1 As String = "00"
                        myReader.Read()
                        checkRowCount = 0
                        For j = 1 To 24
                            If checkRowCount = rowCount Then Exit For
                            Dim test As String = CStr(temp)
                            If j = 2 Then
                                temp1 = "01"
                            ElseIf j = 3 Then
                                temp1 = "02"
                            ElseIf j = 4 Then
                                temp1 = "03"
                            ElseIf j = 5 Then
                                temp1 = "04"
                            ElseIf j = 6 Then
                                temp1 = "05"
                            ElseIf j = 7 Then
                                temp1 = "06"
                            ElseIf j = 8 Then
                                temp1 = "07"
                            ElseIf j = 9 Then
                                temp1 = "08"
                            ElseIf j = 10 Then
                                temp1 = "09"
                            ElseIf j = 11 Then
                                temp1 = "10"
                            ElseIf j = 12 Then
                                temp1 = "11"
                            ElseIf j = 13 Then
                                temp1 = "12"
                            ElseIf j = 14 Then
                                temp1 = "13"
                            ElseIf j = 15 Then
                                temp1 = "14"
                            ElseIf j = 16 Then
                                temp1 = "15"
                            ElseIf j = 17 Then
                                temp1 = "16"
                            ElseIf j = 18 Then
                                temp1 = "17"
                            ElseIf j = 19 Then
                                temp1 = "18"
                            ElseIf j = 20 Then
                                temp1 = "19"
                            ElseIf j = 21 Then
                                temp1 = "20"
                            ElseIf j = 22 Then
                                temp1 = "21"
                            ElseIf j = 23 Then
                                temp1 = "22"
                            ElseIf j = 24 Then
                                temp1 = "23"
                            End If
                            Dim a As String
							a = myReader.GetValue(0).ToString
							If a.Length = 1 Then
								a = "0" + a
							End If

							If temp1 = a Then
								strValues(cycle) += myReader.GetValue(1).ToString() + vbTab
								checkRowCount += 1
								myReader.Read()
							Else
								strValues(cycle) += "0" + vbTab
							End If
                            temp += 1
                        Next
                        If cycle = YearNum Then
                            draw2TimeChart(strValues, strCategory, cycle)
                        End If
                    End If
                    myReader.Close()

                    Dim Datafile As String
                    If optDistChosen Then 'by District
                        Datafile = "Accidents" & indi.year + cycle & "Dist" & distNo
                    Else 'by State
                        Datafile = "Accidents" & indi.year + cycle
                    End If
                    'timeAdapter.Fill(timeDs, "accidents2000")
                    timeAdapter.Fill(timeDs, Datafile)
                Catch e As Exception
                    Session("error") = e.Message
                    Response.Redirect("error.aspx")
                    Exit Sub
                End Try
                Myconnection.Close()
            Next cycle
            'Dim datas As String
            'datas = "datas" & cycle
            Session("datas") = timeDs

        End Sub

        Private Sub generateSql()
            condsql = " WHERE 1=1 "

            If cmbHwyClass.SelectedIndex > 0 Then
                condsql = condsql & generateFieldSql("HWY_CLASS", cmbHwyClass.SelectedIndex)
            End If
            If cmbAccClass.SelectedIndex > 0 Then
                condsql = condsql & generateFieldSql("ACC_CLASS", cmbAccClass.SelectedIndex)
            End If
            Try
                If Len(cmbFuncClass.SelectedItem.Text) > 0 Then
                    If cmbFuncClass.SelectedItem.Text = "Interstate" And cmbFuncClass.SelectedIndex = 2 Then
                        condsql = condsql & "AND FUN_CLASS = '01'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Prin. Arterial" And cmbFuncClass.SelectedIndex = 3 Then
                        condsql = condsql & "AND FUN_CLASS = '02'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Minor Arterial" And cmbFuncClass.SelectedIndex = 4 Then
                        condsql = condsql & "AND FUN_CLASS = '06'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Major Collector" Then
                        condsql = condsql & "AND FUN_CLASS = '07'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Minor Collector" Then
                        condsql = condsql & "AND FUN_CLASS = '08'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Local Systems" And cmbFuncClass.SelectedIndex = 7 Then
                        condsql = condsql & "AND FUN_CLASS = '09'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Interstate" And cmbFuncClass.SelectedIndex = 9 Then
                        condsql = condsql & "AND FUN_CLASS = '11'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Other Freeways" Then
                        condsql = condsql & "AND FUN_CLASS = '12'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Prin. Arterial" And cmbFuncClass.SelectedIndex = 11 Then
                        condsql = condsql & "AND FUN_CLASS = '14'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Minor Arterial" And cmbFuncClass.SelectedIndex = 12 Then
                        condsql = condsql & "AND FUN_CLASS = '16'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Collector" Then
                        condsql = condsql & "AND FUN_CLASS = '17'"
                    ElseIf cmbFuncClass.SelectedItem.Text = "Local Systems" And cmbFuncClass.SelectedIndex = 14 Then
                        condsql = condsql & "AND FUN_CLASS = '19'"
                    End If
                End If
            Catch e2 As Exception
                Session("error") = "Please select a different Function Class"
                Response.Redirect("error.aspx")
            End Try

            condsql = condsql & generateFieldSql("HWY_NUM", GetTrueHnum(txtHwyNum.Text)) & _
             generateFieldSql("CSECT", txtCSec.Text) & _
             generateFieldSql("WEATHER", Left(cmbWeather.SelectedItem.Text, 1)) & _
             generateFieldSql("TYPE_ROAD", Left(cmbRoadType.SelectedItem.Text, 1)) & _
             generateFieldSql("ROAD_COND", Left(cmbRoadCond.SelectedItem.Text, 1)) & _
             generateFieldSql("SURF_TYPE", Left(cmbSufaceType.SelectedItem.Text, 1)) & _
             generateFieldSql("SURF_COND", Left(cmbSufaceCond.SelectedItem.Text, 1)) & _
             generateFieldSql("KIND_LOC", Left(cmbKindofLoc.SelectedItem.Text, 1)) & _
             generateFieldSql("LIGHTING", Left(cmbLighting.SelectedItem.Text, 1)) & _
             generateFieldSql("ALIGNMENT", Left(cmbRoadwayAlign.SelectedItem.Text, 1)) & _
             generateFieldSql("TRAF_CTRL", Left(cmbTrafCtrl.SelectedItem.Text, 1)) & _
             generateFieldSql("COND_DRIV1", Left(cmb1driver.SelectedItem.Text, 1)) & _
             generateFieldSql("MOVEMENT1", Left(cmbMov1.SelectedItem.Text, 1)) & _
             generateFieldSql("MOVEMENT2", Left(cmbMov2.SelectedItem.Text, 1))

            
            If Not (byCollChosen) Or QueryType = 1 Then
                condsql = condsql & generateFieldSql("TYPE_COLL", Left(cmbTypeofColl2.SelectedItem.Text, 1))
            End If

            If Not (byAccChosen) Or QueryType = 1 Then
                condsql = condsql & generateFieldSql("TYPE_ACC", Left(cmbTypeofAcc.SelectedItem.Text, 1))
            End If

            If Not (byViolChosen) Or QueryType = 1 Then
                condsql = condsql & generateFieldSql("VIOLATION1", Left(cmbViolation1.SelectedItem.Text, 1))
            End If

            If cmbPedestrian.SelectedIndex > 0 Then
			
			 Dim fvalue As Integer = cmbPedestrian.SelectedIndex - 1

			   If optDistChosen And (indi.year = "2001" Or indi.year = "2002" Or indi.year = "2003") Then

				 condsql = condsql & "AND [PEDESTRIAN] = " & fvalue
				Else
				   condsql = condsql & "AND [PEDESTRIAN] = '" & fvalue & "' "

			   End If

				' [" & fName & "] = '" & fValue & "' "

			End If

            If cmbInter.SelectedItem.Text <> "Both" Then
                condsql = condsql & generateFieldSql("INTER", cmbInter.SelectedIndex - 1)
            End If



            If txtHwyNum.Text = "" Then
                HigNum = ""
            Else
                HigNum = GetTrueHnum(txtHwyNum.Text)
            End If



            queryString = condsql
            If Not (byPt_ImpactChosen) Or QueryType = 1 Then
                condsql = condsql & generateFieldSql("PT_IMPACT", Left(cmbPoi.SelectedItem.Text, 1))
            End If

            Dim yearstr As String = cmbYear.SelectedItem.Text

            If optstateChosen And (yearstr = "2001" Or yearstr = "2002" Or yearstr = "2003") Then
                queryString = queryString

            Else
                queryString = condsql

            End If


            condsql &= generateFieldSql("VEH1_SPEED", txt1speed.Text)

            If Not (txt1speed.Text = "" Or indi.year = 2004) Then
                'queryString &= " [VEH1_SPEED] = txt1speed.Text "
                queryString &= " AND [VEH1_SPEED] = " & txt1speed.Text & " "
            End If

            If Not (byHourChosen) Or QueryType = 2 Then
                If cmbHourFrom.SelectedIndex < cmbHourTo.SelectedIndex Then
                    If cmbHourFrom.SelectedIndex = 0 Then
                        condsql = condsql & " AND ([HOUR] >= '" & 0 & "' AND [HOUR] <= '" & cmbHourTo.SelectedItem.Text & "')"
                        If indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003 Then
                            queryString &= " AND ([HOUR] >= " & 0 & " AND [HOUR] <= " & CInt(cmbHourTo.SelectedItem.Text) & ")"
                        Else : queryString &= " AND ([HOUR] >= '" & 0 & "' AND [HOUR] <= '" & cmbHourTo.SelectedItem.Text & "')"
                        End If
                    Else
                        condsql = condsql & " AND ([HOUR] >= '" & cmbHourFrom.SelectedItem.Text & "' AND [HOUR] <= '" & cmbHourTo.SelectedItem.Text & "')"
                        If indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003 Then
                            queryString &= " AND ([HOUR] >= " & CInt(cmbHourFrom.SelectedItem.Text) & " AND [HOUR] <= " & CInt(cmbHourTo.SelectedItem.Text) & ")"
                        Else : queryString &= " AND ([HOUR] >= '" & cmbHourFrom.SelectedItem.Text & "' AND [HOUR] <= '" & cmbHourTo.SelectedItem.Text & "')"
                        End If
                    End If
                ElseIf cmbHourFrom.SelectedIndex > cmbHourTo.SelectedIndex Then
                    If cmbHourTo.SelectedIndex = 0 Then
                        condsql = condsql & " AND( [HOUR] >= '" & cmbHourFrom.SelectedItem.Text & "' AND [HOUR] <= '" & 24 & "')"
                        If indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003 Then
                            queryString &= " AND ([HOUR] >= " & CInt(cmbHourFrom.SelectedItem.Text) & " AND [HOUR] <= " & 24 & ")"
                        Else : queryString &= " AND ( [HOUR] >=  '" & cmbHourFrom.SelectedItem.Text & "' AND [HOUR] <= '" & 24 & "')"
                        End If

                    Else
                        condsql = condsql & " OR ([HOUR] >=  '" & 0 & "' AND [HOUR] <= '" & cmbHourTo.SelectedItem.Text & "')"
                        If indi.year = 2001 Or indi.year = 2002 Or indi.year = 2003 Then
                            queryString &= " AND ([HOUR] >= " & 0 & " AND [HOUR] <= " & CInt(cmbHourTo.SelectedItem.Text) & ")"
                        Else : queryString &= " AND ([HOUR] >= '" & 0 & "' AND [HOUR] <= '" & cmbHourTo.SelectedItem.Text & "')"
                        End If

                    End If
				End If

If optDistChosen And (yearstr = "2001" Or yearstr = "2002" Or yearstr = "2003") Then
   condsql = queryString
End If

			End If

        End Sub

        Private Sub draw2TimeChart(ByRef yValues() As String, ByRef strCat As String, ByRef SeriesNum As Integer)
            Dim objCSpace As Microsoft.Office.Interop.Owc11.ChartSpace = New Microsoft.Office.Interop.Owc11.ChartSpaceClass()
            Dim objChart
            Dim i As Integer

            objChart = objCSpace.Charts.Add(0)
            objChart.Type = Microsoft.Office.Interop.Owc11.ChartChartTypeEnum.chChartTypeColumnClustered
            'Specify if the chart needs to have legend.
            objChart.HasLegend = True

            'Give title to graph.
            caption = GetCaption()
            objChart.HasTitle = True
            'objChart.SeriesCollection(0).Interior.Color = "Rosybrown"'set the color of chart

            If byMonthChosen Then
                objChart.Title.Caption = "Results Displayed by Month--" & caption
            ElseIf byDayChosen Then
                objChart.Title.Caption = "Results Displayed by Day--" & caption
            ElseIf byHourChosen Then
                objChart.Title.Caption = "Results Displayed by Hour--" & caption
            End If

            'Give the caption for the X axis and Y axis of the graph
            If byMonthChosen Then
                objChart.Axes(0).HasTitle = True
                objChart.Axes(0).Title.Caption = "Month"
                objChart.Axes(1).HasTitle = True
                objChart.Axes(1).Title.Caption = "AccNum"

            ElseIf byDayChosen Then
                objChart.Axes(0).HasTitle = True
                objChart.Axes(0).Title.Caption = "Day"
                objChart.Axes(1).HasTitle = True
                objChart.Axes(1).Title.Caption = "AccNum"
            Else
                objChart.Axes(0).HasTitle = True
                objChart.Axes(0).Title.Caption = "Hour"
                objChart.Axes(1).HasTitle = True
                objChart.Axes(1).Title.Caption = "AccNum"
            End If
            'Add a series to the chart’s series collection
            For i = 0 To SeriesNum
                objChart.SeriesCollection.Add(i)
            Next i

            For i = 0 To SeriesNum
                objChart.SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimSeriesNames, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, indi.year + i)
                'Give the Categories
                objChart.SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimCategories, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, strCat)
                'Give The values
                objChart.SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimValues, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, yValues(i))
            Next i

            'Now a chart is ready to export to a GIF.
            showlargechar = False 'don't used large chart
			Dim ChartName As String = Rnd() & ".gif"	  'times & Rnd() & ".gif"
            Dim strAbsolutePath As String = Server.MapPath(".") & "\" & ChartName
            objCSpace.ExportPicture(strAbsolutePath, "GIF", 900, 450) '******************************************

            Dim strRelativePath As String = "./" & ChartName
            Session("relPath") = strRelativePath
            'Catch e As Exception
            '    Session("error") = e.Message
            '    Response.Redirect("error.aspx")
            'End Try
        End Sub

        'The following functions generate the category values for the charts
        Private Sub genStrCat4Month(ByRef strCat As String)
            strCat = "January" & vbTab & "February" & vbTab & _
                "March" & vbTab & "April" & vbTab & _
                "May" & vbTab & "June" & vbTab & _
                "July" & vbTab & "August" & vbTab & _
                "September" & vbTab & "October" & vbTab & _
                "November" & vbTab & "December" & vbTab
        End Sub
        Private Sub genStrCat4Day(ByRef strCat As String)
            strCat = "Monday" & vbTab & "Tuesday" & vbTab & _
                "Wednesday" & vbTab & "Thursday" & vbTab & _
                "Friday" & vbTab & "Saturday" & vbTab & _
                "Sunday" & vbTab
        End Sub
        Private Sub genStrCat4Hour(ByRef strCat As String)
            strCat = "1" & vbTab & "2" & vbTab & _
                "3" & vbTab & "4" & vbTab & _
                "5" & vbTab & "6" & vbTab & _
                "7" & vbTab & "8" & vbTab & _
                "9" & vbTab & "10" & vbTab & _
                "11" & vbTab & "12" & vbTab & _
                "13" & vbTab & "14" & vbTab & _
                "15" & vbTab & "16" & vbTab & _
                "17" & vbTab & "18" & vbTab & _
                "19" & vbTab & "20" & vbTab & _
                "21" & vbTab & "22" & vbTab & _
                "23" & vbTab & "24" & vbTab
        End Sub

        Private Sub sqlStmt4Month(ByRef tempsql As String, ByRef year As String)
            If optDistChosen Then 'by District
                'tempsql = "SELECT format([ACC_DATE], ""mm/yyyy"") as MONTH1, COUNT(*) as AccNum, " & _
                '              "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                '              "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                '              "FROM accidents" & year & "Dist" & distNo & condsql & " AND ACC_DATE IS NOT NULL group by format([ACC_DATE], ""mm/yyyy"")"

                tempsql = "SELECT format([ACC_DATE], ""mm/yyyy"") as MONTH1, COUNT(*) as AccNum, " & _
                            "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                            "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                           "FROM a" & year & "D" & distNo & condsql & " AND ACC_DATE IS NOT NULL group by format([ACC_DATE], ""mm/yyyy"")"


            Else 'by State
                tempsql = "SELECT format([ACC_DATE], ""mm/yyyy"") as MONTH1, COUNT(*) as AccNum," & _
              "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
              "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
              "FROM " & mdbtable & year & condsql & " AND ACC_DATE IS NOT NULL group by format([ACC_DATE], ""mm/yyyy"")"
            End If
            'condsql = tempsql
        End Sub

        Private Sub sqlStmt4Day(ByRef tempsql As String, ByRef year As String)
            If optDistChosen Then 'by District
                tempsql = "SELECT [WEEKDAY], COUNT(*) as AccNum, " & _
                            "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                            "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                            "FROM a" & year & "D" & distNo & condsql & " AND WEEKDAY IS NOT NULL group by [WEEKDAY]"
              
            Else 'by State
                tempsql = "SELECT [WEEKDAY], COUNT(*) as AccNum, " & _
                            "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious," & _
                            "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                            "FROM " & mdbtable & year & condsql & " AND WEEKDAY IS NOT NULL group by [WEEKDAY]"
            End If
            'condsql = tempsql
        End Sub

        Private Sub sqlStmt4Hour(ByRef tempsql As String, ByRef year As String)
            If optDistChosen Then 'by District
                tempsql = "SELECT [HOUR], COUNT(*) as AccNum, " & _
                        "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                        "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                        "FROM a" & year & "D" & distNo & condsql & " AND HOUR IS NOT NULL group by [HOUR]"
            Else 'by State
                tempsql = "SELECT [HOUR], COUNT(*) as AccNum, " & _
                        "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                        "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                        "FROM " & mdbtable & year & condsql & " AND HOUR IS NOT NULL group by [HOUR]"
            End If
            'condsql = tempsql
        End Sub

        Private Function GetCaption() As String
            Dim tempcaption As String
            If optstateChosen Then 'if choose the state
                tempcaption = "Analysis by whole state,"
            Else
                tempcaption = "Analysis by district:" & distNo & ","
            End If
            If Len(cmbYear.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & "Year:" & indi.year
            End If
            If Len(cmbYear2.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & " to " & cmbYear2.SelectedItem.Text
            End If
            If Len(txtHwyNum.Text) > 0 Then
                tempcaption = tempcaption & ", Highway Number:" & txtHwyNum.Text
            End If
            If Len(txtCSec.Text) > 0 Then
                tempcaption = tempcaption & ", Control Section:" & txtCSec.Text
            End If
            If Not (byHourChosen) Then
                If Len(cmbHourFrom.SelectedItem.Text) > 0 Then
                    tempcaption = tempcaption & ", Hour From:" & cmbHourFrom.SelectedItem.Text
                    If cmbHourTo.SelectedIndex = 0 Then
                        tempcaption = tempcaption & ", Hour To:" & "24"
                    End If
                End If
                If Len(cmbHourTo.SelectedItem.Text) > 0 Then
                    If cmbHourFrom.SelectedIndex = 0 Then
                        tempcaption = tempcaption & ", Hour From:" & "0"
                    End If
                    tempcaption = tempcaption & ", Hour To:" & cmbHourTo.SelectedItem.Text
                End If
            End If
            If Len(cmbHwyClass.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Highway Class:" & cmbHwyClass.SelectedItem.Text
            End If
            If Len(cmbFuncClass.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Function Class:" & cmbFuncClass.SelectedItem.Text
            End If
            If Len(cmbAccClass.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Accident Class:" & cmbAccClass.SelectedItem.Text
            End If
            If Len(cmbWeather.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Weather Condition:" & cmbWeather.SelectedItem.Text
            End If
            If Len(cmbSufaceType.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Surface Type:" & cmbSufaceType.SelectedItem.Text
            End If
            If Len(cmbRoadType.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Road Type:" & cmbRoadType.SelectedItem.Text
            End If
            If Len(cmbSufaceCond.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Surface Condition:" & cmbSufaceCond.SelectedItem.Text
            End If
            If Len(cmbTypeofColl2.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Type of Collision:" & cmbTypeofColl2.SelectedItem.Text
            End If
            If Len(cmbRoadwayAlign.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Roadway Alignment:" & cmbRoadwayAlign.SelectedItem.Text
            End If
            If Len(cmbPedestrian.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Pedestrian:" & cmbPedestrian.SelectedItem.Text
            End If
            If Len(txt1speed.Text) > 0 Then
                tempcaption = tempcaption & ", 1st vehicle's speed :" & txt1speed.Text
            End If
            If Len(cmb1driver.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", 1st Driver's condition:" & cmb1driver.SelectedItem.Text
            End If
            If Len(cmbTrafCtrl.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Traffic Control:" & cmbTrafCtrl.SelectedItem.Text
            End If
            If Len(cmbKindofLoc.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Kind of Location:" & cmbKindofLoc.SelectedItem.Text
            End If
            If Len(cmbLighting.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Lighting:" & cmbLighting.SelectedItem.Text
            End If
            If Len(cmbViolation1.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Violation of 1st Vehicle:" & cmbViolation1.SelectedItem.Text
            End If
            If Len(cmbTypeofAcc.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Type of Accident:" & cmbTypeofAcc.SelectedItem.Text
            End If
            If Len(cmbPoi.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Point of Impact:" & cmbPoi.SelectedItem.Text
            End If
            If Len(cmbRoadCond.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Road Condition:" & cmbRoadCond.SelectedItem.Text
            End If
            If Len(cmbMov1.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Movement of vehicle1:" & cmbMov1.SelectedItem.Text
            End If
            If Len(cmbMov2.SelectedItem.Text) > 0 Then
                tempcaption = tempcaption & ", Movement of vehicle2:" & cmbMov2.SelectedItem.Text
            End If
            tempcaption = tempcaption & ", Inter or segment:" & cmbInter.SelectedItem.Text
            GetCaption = tempcaption
        End Function

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

        Private Sub CancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelButton.Click
            Response.Redirect("options.aspx")
        End Sub

        Private Sub cmdTypeColl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTypeColl.Click


            If cmbYear.SelectedIndex = 0 Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('You must select one year for your query!')"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

            If cmbYear2.Enabled = True Then
                If CInt(cmbYear.SelectedIndex.ToString) >= CInt(cmbYear2.SelectedIndex.ToString) And cmbYear2.SelectedIndex.ToString <> "0" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The YearFrom can not be greater than or equal to YearTo!')"
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

            If txtHwyNum.Text = "" Then
                HigNum = ""
            Else
                HigNum = GetTrueHnum(txtHwyNum.Text)
            End If

            indi.year = cmbYear.SelectedItem.Text
            byCollChosen = True
            byAccChosen = False
            byPt_ImpactChosen = False
            byViolChosen = False

            Call genChaResults()
            NoShow = True
            QueryType = 2 '2=By Crash Characteristics
			SubQueryType = 1 'by Type of Coll

	saveintoSession()
            Response.Write("<script language='javascript'>window.open('rep2TypeColl2.aspx');</Script>")
        End Sub

        Private Sub cmdTypeAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTypeAcc.Click


            If cmbYear.SelectedIndex = 0 Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('You must select one year for your query!')"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

            If cmbYear2.Enabled = True Then
                If cmbYear2.SelectedIndex = 0 Then
                    mutipleYear = False
                    YearNum = 0
                End If
                If CInt(cmbYear.SelectedIndex.ToString) > CInt(cmbYear2.SelectedIndex.ToString) And cmbYear2.SelectedIndex.ToString <> "0" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The YearFrom can not be greater than or equal to YearTo!')"
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

            If txtHwyNum.Text = "" Then
                HigNum = ""
            Else
                HigNum = GetTrueHnum(txtHwyNum.Text)
            End If

            indi.year = cmbYear.SelectedItem.Text
            byCollChosen = False
            byAccChosen = True
            byPt_ImpactChosen = False
            byViolChosen = False

            Call genChaResults()
            NoShow = True
            QueryType = 2 '2=By Crash Characteristics
			SubQueryType = 2 'by Type of Acc

		 saveintoSession()
            Response.Write("<script language='javascript'>window.open('rep2TypeAcc.aspx')</script>;")
        End Sub

        Private Sub cmdPt_Impact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPt_Impact.Click


            If cmbYear.SelectedIndex = 0 Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('You must select one year for your query!')"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

            If cmbYear2.Enabled = True Then
                If CInt(cmbYear.SelectedIndex.ToString) > CInt(cmbYear2.SelectedIndex.ToString) And cmbYear2.SelectedIndex.ToString <> "0" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The YearFrom can not be greater than or equal to YearTo!')"
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

            If txtHwyNum.Text = "" Then
                HigNum = ""
            Else
                HigNum = GetTrueHnum(txtHwyNum.Text)
            End If

            indi.year = cmbYear.SelectedItem.Text
            byPt_ImpactChosen = True
            byCollChosen = False
            byAccChosen = False
            byViolChosen = False

            Call genChaResults()
            NoShow = True
            QueryType = 2 '2=By Crash Characteristics
			SubQueryType = 3 'by Point of Impact

				 saveintoSession()
            Response.Write("<script language ='javascript'>window.open('rep2POI.aspx');</script>")

        End Sub

        Private Sub cmdViol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdViol.Click


            If cmbYear.SelectedIndex = 0 Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('You must select one year for your query!')"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

            If cmbYear2.Enabled = True Then
                If CInt(cmbYear.SelectedIndex.ToString) > CInt(cmbYear2.SelectedIndex.ToString) And cmbYear2.SelectedIndex.ToString <> "0" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The YearFrom can not be greater than or equal to YearTo!')"
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

            If txtHwyNum.Text = "" Then
                HigNum = ""
            Else
                HigNum = GetTrueHnum(txtHwyNum.Text)
            End If

            indi.year = cmbYear.SelectedItem.Text
            byPt_ImpactChosen = False
            byCollChosen = False
            byAccChosen = False
            byViolChosen = True

            Call genChaResults()
            NoShow = True
            QueryType = 2 '2=By Crash Characteristics
			SubQueryType = 4 'by Violation of 1st Vehicle

			 saveintoSession()
            Response.Write("<script language ='javascript'>window.open('rep2Viol1.aspx');</script>")
        End Sub

        Private Sub genChaResults()
            Dim tempSql As String
            Dim strValues(YearNum) As String 'used for chart created(value)
            Dim strCategory As String
            Dim j As Integer
            Dim myReader As OleDbDataReader
            Dim cmd As OleDbCommand
            Dim myAdapter As OleDbDataAdapter
            Dim ds As DataSet
            ds = New DataSet()

            Call generateSql()
            If byCollChosen Then
                Call genStrCat4coll(strCategory)
            ElseIf byAccChosen Then
                Call genStrCat4acc(strCategory)
            ElseIf byPt_ImpactChosen Then
                Call genStrCat4Pt_Impact(strCategory)
            ElseIf byViolChosen Then
                Call genStrCat4Viol(strCategory)
            End If

            Dim cycle As Integer
            For cycle = 0 To YearNum
                If optDistChosen Then 'by District
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\Data\GisPrj\District; Extended Properties =DBASE III;"
                Else 'by State
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= C:\Data\" & statemdb & indi.year + cycle & ".mdb"
                End If

                'get the sql of check for every year
                If byCollChosen Then
                    Call sqlStmt4coll(tempSql, indi.year + cycle)
                ElseIf byAccChosen Then
                    Call sqlStmt4acc(tempSql, indi.year + cycle)
                ElseIf byPt_ImpactChosen Then
                    Call sqlStmt4Pt_Impact(tempSql, indi.year + cycle)
                ElseIf byViolChosen Then
                    Call sqlStmt4Viol(tempSql, indi.year + cycle)
                End If
                If optDistChosen Then
                    tempSql = ChangeExp(tempSql, indi.year)
                End If



                Myconnection = Nothing
                Myconnection = New OleDbConnection(connString)
                myReader = Nothing
                cmd = Nothing
                myAdapter = Nothing

                openDbConn()
                cmd = New OleDbCommand(tempSql, Myconnection)
                myAdapter = New OleDbDataAdapter(tempSql, Myconnection)
                myReader = cmd.ExecuteReader()

                Dim numbercols As Integer = myReader.FieldCount
                Dim tempcoll As Integer = 65
                Dim rowCount As Integer
                Dim checkRowCount As Integer
                rowCount = 0
                While myReader.Read
                    rowCount = rowCount + 1
                End While
                myReader.Close()
                myReader = cmd.ExecuteReader()
                Try
                    If byCollChosen Then
                        myReader.Read()
                        checkRowCount = 0
                        For j = 1 To 7
                            If checkRowCount = rowCount Then Exit For
                            If tempcoll = Asc(myReader.GetValue(0).ToString()) Then
                                strValues(cycle) += myReader.GetValue(1).ToString() + vbTab
                                checkRowCount += 1
                                myReader.Read()
                            Else
                                strValues(cycle) += "0" + vbTab
                            End If
                            tempcoll += 1
                        Next
                        If cycle = YearNum Then
                            draw2CharChart(strValues, strCategory, cycle)
                        End If
                    ElseIf byAccChosen Then
                        myReader.Read()
                        checkRowCount = 0
                        For j = 1 To 11
                            If checkRowCount = rowCount Then Exit For
                            If tempcoll = Asc(myReader.GetValue(0).ToString()) Then
                                strValues(cycle) += myReader.GetValue(1).ToString() + vbTab
                                checkRowCount += 1
                                myReader.Read()
                            Else
                                strValues(cycle) += "0" + vbTab
                            End If
                            tempcoll += 1
                        Next
                        If cycle = YearNum Then
                            draw2CharChart(strValues, strCategory, cycle)
                        End If
                    ElseIf byPt_ImpactChosen Then
                        myReader.Read()
                        checkRowCount = 0
                        For j = 1 To 9
                            If checkRowCount = rowCount Then Exit For
                            If tempcoll = Asc(myReader.GetValue(0).ToString()) Then
                                strValues(cycle) += myReader.GetValue(1).ToString() + vbTab
                                checkRowCount = checkRowCount + 1
                                myReader.Read()
                            Else
                                strValues(cycle) += "0" + vbTab
                            End If
                            tempcoll += 1
                        Next
                        If cycle = YearNum Then
                            draw2CharChart(strValues, strCategory, cycle)
                        End If
                    ElseIf byViolChosen Then
                        myReader.Read()
                        checkRowCount = 0
                        For j = 1 To 20
                            If checkRowCount = rowCount Then Exit For
                            If tempcoll = Asc(myReader.GetValue(0).ToString()) Then
                                strValues(cycle) += myReader.GetValue(1).ToString() + vbTab
                                checkRowCount += 1
                                myReader.Read()
                            Else
                                strValues(cycle) += "0" + vbTab
                            End If
                            tempcoll += 1
                        Next
                        If cycle = YearNum Then
                            draw2CharChart(strValues, strCategory, cycle)
                        End If
                    End If
                    myReader.Close()

                    Dim Datafile As String
                    If optDistChosen Then 'by District
                        Datafile = "Accidents" & indi.year + cycle & "Dist" & distNo
                    Else 'by State
                        Datafile = "Accidents" & indi.year + cycle
                    End If
                    myAdapter.Fill(ds, Datafile)

                Catch e As Exception
                    Exit Sub
                End Try
                closeDbConn()
            Next cycle
            Session("datas") = ds
        End Sub


        Private Sub draw2CharChart(ByRef yValues() As String, ByRef strCat As String, ByRef SeriesNum As Integer)
            Dim i As Integer
            Dim objCSpace As Microsoft.Office.Interop.Owc11.ChartSpace = New Microsoft.Office.Interop.Owc11.ChartSpaceClass()
            Dim objChart
            Dim temp As New Bitmap(1000, 500)

            'Try
            objChart = objCSpace.Charts.Add(0)
            objChart.Type = Microsoft.Office.Interop.Owc11.ChartChartTypeEnum.chChartTypeColumnClustered
            'Specify if the chart needs to have legend.
            objChart.HasLegend = True

            caption = GetCaption()
            'Give title to graph.
            objChart.HasTitle = True
            If byCollChosen Then
                objChart.Title.Caption = "Results Displayed by Type of Collision--" & caption

                'Give the caption for the X axis and Y axis of the graph
                objChart.Axes(0).HasTitle = True
                objChart.Axes(0).Title.Caption = "Type of Collision"
                objChart.Axes(1).HasTitle = True
                objChart.Axes(1).Title.Caption = "AccNum"

            ElseIf byAccChosen Then
                objChart.Title.Caption = "Results Displayed by Type of Accident--" & caption

                'Give the caption for the X axis and Y axis of the graph
                objChart.Axes(0).HasTitle = True
                objChart.Axes(0).Title.Caption = "Type of Accident"
                objChart.Axes(1).HasTitle = True
                objChart.Axes(1).Title.Caption = "AccNum"

            ElseIf byPt_ImpactChosen Then
                objChart.Title.Caption = "Results Displayed by Point of Impact--" & caption

                'Give the caption for the X axis and Y axis of the graph
                objChart.Axes(0).HasTitle = True
                objChart.Axes(0).Title.Caption = "Point of Impact"
                objChart.Axes(1).HasTitle = True
                objChart.Axes(1).Title.Caption = "AccNum"

            ElseIf byViolChosen Then
                objChart.Title.Caption = "Results Displayed by Violation of 1st Vehicle--" & caption

                'Give the caption for the X axis and Y axis of the graph
                objChart.Axes(0).HasTitle = True
                objChart.Axes(0).Title.Caption = "Violation of 1st Vehicle"
                objChart.Axes(1).HasTitle = True
                objChart.Axes(1).Title.Caption = "AccNum"
            End If

            'objChart.Title.Caption = objChart.Title.Caption & " in the year " & cmbYear.SelectedItem.Text
            'Add a series to the chart’s series collection
            For i = 0 To SeriesNum
                objChart.SeriesCollection.Add(i)
            Next i

            For i = 0 To SeriesNum
                objChart.SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimSeriesNames, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, indi.year + i)
                'Give the Categories
                objChart.SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimCategories, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, strCat)
                'Give The values
                objChart.SeriesCollection(i).SetData(Microsoft.Office.Interop.Owc11.ChartDimensionsEnum.chDimValues, Microsoft.Office.Interop.Owc11.ChartSpecialDataSourcesEnum.chDataLiteral, yValues(i))
            Next i

            'Now a chart is ready to export to a GIF.
            showlargechar = False 'don't used large chart
			Dim ChartName As String = Rnd() & ".gif"  'times & Rnd() & ".gif"
            Dim strAbsolutePath As String = Server.MapPath(".") & "\" & ChartName
            objCSpace.ExportPicture(strAbsolutePath, "GIF", 900, 450)

            Dim strRelativePath As String = "./" & ChartName
            Session("relPath") = strRelativePath
        End Sub

        'The following functions generate the category values for the charts
        Private Sub genStrCat4coll(ByRef strCat As String) '
            strCat = "Non-collision" & vbTab & "Rear end" & vbTab & _
                "Head-on" & vbTab & _
                "Right angle" & vbTab & "Left turn1" & vbTab & _
                "Left turn2" & vbTab & "Left turn3" & vbTab & _
                "Right turn1" & vbTab & "Right turn2" & vbTab & _
                "SideSwipe (Same dir.)" & vbTab & "SideSwipe (Opposite dir.)" & vbTab & "Other"
        End Sub
        Private Sub genStrCat4acc(ByRef strCat As String)
            strCat = "Running off Roadway" & vbTab & "Overturning on Roadway" & vbTab & _
                "Collision with Pedestrian" & vbTab & "Collision With Other motor vehicle in traffic" & vbTab & _
                "Collision with Parked Vehicle" & vbTab & "Collision with Train" & vbTab & _
                "Collision with Bicyclist" & vbTab & "Collison with Animal" & vbTab & _
                "Collison with Fixed Object" & vbTab & "Collison with Other Object" & vbTab & _
                "Other non-collison on Road" & vbTab
        End Sub
        Private Sub genStrCat4Pt_Impact(ByRef strCat As String)
            'comment contents match 1998's data
            'strCat = "Main Travel Lane" & Char.ToString(vbTab) & "Improved Shoulder Left" & Char.ToString(vbTab) & _
            '        "Improved Shoulder Right" & Char.ToString(vbTab) & "Off Roadway Left" & Char.ToString(vbTab) & _
            '        "Off Roadway Right" & Char.ToString(vbTab) & "Off Roadway Straight Ahead" & Char.ToString(vbTab) & _
            '        "Off Roadway Direction Unknown" & Char.ToString(vbTab) & "Marked Pedestrian Crosswalk" & Char.ToString(vbTab) & _
            '        "Left turn lane/Non-freeway" & Char.ToString(vbTab) & "Right turn lane/Non--freeway" & Char.ToString(vbTab) & _
            '        "Median Opening" & Char.ToString(vbTab) & "Ramp Nose" & Char.ToString(vbTab) & _
            '        "Curb Return" & Char.ToString(vbTab) & "Traffic Island" & Char.ToString(vbTab) & _
            '        "Off Ramp taper or decel lane" & Char.ToString(vbTab) & "Off Ramp Roadway" & Char.ToString(vbTab) & _
            '        "Off Ramp Terminal" & Char.ToString(vbTab) & "On Ramp taper or decel lane" & Char.ToString(vbTab) & _
            '        "On Ramp Roadway" & Char.ToString(vbTab) & "Auxiliary Lane or Coll. Road" & Char.ToString(vbTab) & _
            '        "Freeway to freeway Connect" & Char.ToString(vbTab) & "Service Road" & Char.ToString(vbTab) & _
            '        "Within Construction Zone" & Char.ToString(vbTab) & "Other" & Char.ToString(vbTab) & _
            '        "Impact Attenuator" & Char.ToString(vbTab) & "Private property/park Lot" & Char.ToString(vbTab)
            strCat = "On roadway" & vbTab & "Shoulder" & vbTab & _
                "Median" & vbTab & "Beyond shoulder-Left" & vbTab & _
                "Beyond shoulder-Right" & vbTab & "Off roadway" & vbTab & _
                "Gore" & vbTab & "Unknown" & vbTab & _
                "Other"
            '& _Median Opening" & Char.ToString(vbTab) & "Right turn lane, Non--freeway" & Char.ToString(vbTab)
        End Sub
        Private Sub genStrCat4Viol(ByRef strCat As String)
            strCat = "Exceeding stated speed limit" & vbTab & "Exceeding safe speed limit" & vbTab & _
                "Failure to yield" & vbTab & "Following too close" & vbTab & _
                "Driving left of center" & vbTab & "Cutting in, improper passing" & vbTab & _
                "Failure to signal" & vbTab & "Made wide right turn" & vbTab & _
                "Cut corner on left turn" & vbTab & "Turned from wrong lane" & vbTab & _
                "Other improper turning" & vbTab & "Disregarded traffic control" & vbTab & _
                "Improper starting" & vbTab & "Improper parking" & vbTab & _
                "Failure to set out flags" & vbTab & "Failed to dim headlights" & vbTab & _
                "Vehicle condition" & vbTab & "Driver condition" & vbTab & _
                "Careless operation" & vbTab & "Unknown violations" & vbTab & "No violation" & vbTab & _
                "Other"
        End Sub

        Private Sub sqlStmt4coll(ByRef tempsql As String, ByRef year As String)
            If optDistChosen Then 'by District
                tempsql = "SELECT [TYPE_COLL], COUNT(*) as AccNum, " & _
                       "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                       "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                       "FROM a" & year & "D" & distNo & condsql & " AND TYPE_COLL IS NOT NULL group by [TYPE_COLL] " & _
                       "ORDER BY [TYPE_COLL]"
            Else 'by State
                tempsql = "SELECT [TYPE_COLL], COUNT(*) as AccNum, " & _
                       "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                       "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                       "FROM " & mdbtable & year & condsql & " AND TYPE_COLL IS NOT NULL group by [TYPE_COLL] " & _
                       "ORDER BY [TYPE_COLL]"
            End If
        End Sub
        Private Sub sqlStmt4acc(ByRef tempsql As String, ByRef year As String)
            If optDistChosen Then 'by District
                tempsql = "SELECT [TYPE_ACC], COUNT(*) as AccNum, " & _
                        "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                        "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                         "FROM a" & year & "D" & distNo & condsql & " AND TYPE_ACC IS NOT NULL group by [TYPE_ACC] " & _
                        "ORDER BY [TYPE_ACC]"
            Else 'by State
                tempsql = "SELECT [TYPE_ACC], COUNT(*) as AccNum, " & _
                        "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                        "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                        "FROM " & mdbtable & year & condsql & " AND TYPE_ACC IS NOT NULL group by [TYPE_ACC] " & _
                        "ORDER BY [TYPE_ACC]"
            End If
        End Sub
        Private Sub sqlStmt4Pt_Impact(ByRef tempsql As String, ByRef year As String)
            If optDistChosen Then 'by District
                tempsql = "SELECT [PT_IMPACT], COUNT(*) as AccNum, " & _
                        "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                        "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                         "FROM a" & year & "D" & distNo & condsql & " AND PT_IMPACT IS NOT NULL group by [PT_IMPACT] " & _
                        "ORDER BY [PT_IMPACT]"
            Else 'by State
                tempsql = "SELECT [PT_IMPACT], COUNT(*) as AccNum, " & _
                        "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                        "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                        "FROM " & mdbtable & year & condsql & " AND PT_IMPACT IS NOT NULL group by [PT_IMPACT] " & _
                        "ORDER BY [PT_IMPACT]"
            End If
        End Sub
        Private Sub sqlStmt4Viol(ByRef tempsql As String, ByRef year As String)
            If optDistChosen Then 'by District
                tempsql = "SELECT [VIOLATION1], COUNT(*) as AccNum, " & _
                        "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                        "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                         "FROM a" & year & "D" & distNo & condsql & " AND VIOLATION1 IS NOT NULL group by [VIOLATION1] " & _
                        "ORDER BY [VIOLATION1]"
            Else 'by State
                tempsql = "SELECT [VIOLATION1], COUNT(*) as AccNum, " & _
                        "SUM([NUM_KILLED]) as Fatal, SUM([NUM_INJ2]) as Critical, SUM([NUM_INJ3]) as Serious, " & _
                        "SUM([NUM_INJ4]) as Severe, SUM([NUM_INJ5]) as Moderate, SUM([NUM_INJ6]) as Minor " & _
                        "FROM " & mdbtable & year & condsql & " AND VIOLATION1 IS NOT NULL group by [VIOLATION1] " & _
                        "ORDER BY [VIOLATION1]"
            End If
        End Sub

        Protected Sub cmbHourFrom_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbHourFrom.SelectedIndexChanged
            If Len(cmbHourFrom.SelectedItem.Text) > 0 Then
                cmdHour.Enabled = False
            ElseIf Len(cmbHourFrom.SelectedItem.Text) = 0 Then
                cmdHour.Enabled = True
            End If
        End Sub

        Protected Sub cmbHourTo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbHourTo.SelectedIndexChanged
            If Len(cmbHourTo.SelectedItem.Text) > 0 Then
                cmdHour.Enabled = False
                cmbPoi.Enabled = False
            ElseIf Len(cmbHourTo.SelectedItem.Text) = 0 Then
                cmdHour.Enabled = True
                cmbPoi.Enabled = True
            End If
        End Sub


       
        Protected Sub cmbYear_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbYear.SelectedIndexChanged

            Dim yearstr As String = cmbYear.SelectedItem.Text
            If optDistChosen Then
                If yearstr = "2001" Or yearstr = "2002" Or yearstr = "2003" Then
                    cmdPt_Impact.Enabled = False
                    cmbPoi.Enabled = False
                Else
                    cmdPt_Impact.Enabled = True
                    cmbPoi.Enabled = True
                End If
            End If
        End Sub

        Protected Sub cmbYear2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbYear2.SelectedIndexChanged

            Dim yearstr As String = cmbYear.SelectedItem.Text
            If optDistChosen Then
                If yearstr = "2001" Or yearstr = "2002" Or yearstr = "2003" Then
                    cmdPt_Impact.Enabled = False
                Else
                    cmdPt_Impact.Enabled = True
                End If
            End If
		End Sub

	Public Sub openDbConn()
			On Error GoTo ce
			Myconnection.Open()
ce:
			strscript = "<script language='javascript'>"
			strscript = strscript & "alert('Error Number and Description'& Err.Num & Err.Description)"
			strscript = strscript & "</script>"
			'RegisterClientScriptBlock("Msg", strscript.ToString)
			Exit Sub

		End Sub

		Public Sub closeDbConn()
			Myconnection.Close()
		End Sub

Sub saveintoSession()

				'Dim ShowAcc As Boolean

			Session("QueryType") = QueryType
			Session("SubQueryType") = SubQueryType
			Session("indi") = indi
			Session("caption") = caption
			Session("caption1") = caption1
			Session("caption2") = caption2
			Session("queryString") = queryString
			Session("mutipleYear ") = mutipleYear
			Session("YearNum") = YearNum

			Session("connString") = connString
			Session("showlargechar") = showlargechar
			Session("NoShow") = NoShow

	 Session("byMonthChosen") = byMonthChosen
		Session("byDayChosen") = byDayChosen
		Session("byHourChosen") = byHourChosen
			Session("byCollChosen") = byCollChosen
	 Session("byViolChosen") = byViolChosen
		 Session("byAccChosen") = byAccChosen
		Session("byPt_ImpactChosen") = byPt_ImpactChosen




		End Sub

Sub getSession()
		optstateChosen = Session("optstateChosen")
	optDistChosen = Session("optDistChosen")
	distNo = Session("distNo")
		QueryType = Session("QueryType")
	 ShowAcc = Session("ShowAcc")
End Sub
    End Class

End Namespace
