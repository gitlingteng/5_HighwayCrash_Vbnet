Imports System.Data.OleDb
Imports System.Data


Namespace Crashsafe


Partial Class rep1Day
		Inherits System.Web.UI.Page
		'global variable
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


		'end of gobal variable

        Dim cstype As Type = Me.GetType()
        Dim myTotal() As Integer = {0, 0, 0, 0, 0, 0, 0, 0}

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label


    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

		  getSession()

		If NoShow Then
			cmbYear.Enabled = False
			cmbYear.Visible = False
			BtnView.Enabled = False
			BtnView.Visible = False
			'
			Dim i As Integer
			Dim rep2DsTOA As DataSet

			Label1.Text = caption

			rep2DsTOA = New DataSet()
			rep2DsTOA = Session("datas")

			Dim Datafile As String
			If optDistChosen Then 'by District
				Datafile = "Accidents" & indi.year & "Dist" & distNo
			Else 'by State
				Datafile = "Accidents" & indi.year
			End If
			DataGrid1.DataSource = rep2DsTOA.Tables(Datafile).DefaultView
			DataGrid1.DataBind()

				For i = 0 To DataGrid1.Items.Count - 1
					Dim wekday As String = DataGrid1.Items(i).Cells(0).Text

					If optDistChosen Then
						If indi.year = "2001" Or indi.year = "2002" Or indi.year = "2003" Then
						wekday = changWeekdayback(wekday)
						End If
					End If

					DataGrid1.Items(i).Cells(0).Text = GetDay(wekday)
				Next i

			If mutipleYear = True Then
				cmbYear.Enabled = True
				cmbYear.Visible = True
				BtnView.Enabled = True
				BtnView.Visible = True

				Dim j As Integer
				For j = 0 To YearNum
					cmbYear.Items.Add(indi.year + j)
				Next j
			End If
			NoShow = False
		End If
    End Sub
        Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, _
                                      ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound

            Dim j, n As Integer

            n = e.Item.Cells.Count

            Select Case e.Item.ItemType

                Case ListItemType.AlternatingItem, ListItemType.Item
                    For j = 1 To n - 1
                        'Calculate total for the field of each row and alternating row.
                        myTotal(j) += CInt(e.Item.Cells(j).Text)
                        'Format the data, and then align the text of each cell to the right.
                        e.Item.Cells(j).Attributes.Add("align", "central")
                    Next
                Case ListItemType.Footer
                    'Use the footer to display the summary row.
                    e.Item.Cells(0).Text = "Total"
                    For j = 1 To n - 1
                        e.Item.Cells(j).Attributes.Add("align", "central")
                        e.Item.Cells(j).Text = myTotal(j)
                    Next
            End Select

        End Sub
    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick
        'Response.Redirect("Map.aspx")
        If Not (mutipleYear) And indi.year = 1999 Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('We are sorry about that we do not have the crash map of 1999!')"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            Exit Sub
        End If
        Response.Write("<script language ='javascript'>window.open('Map.aspx?');</script>")
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
            Response.Write("<script language ='javascript'> window.opener = window; window.close();</script>")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Response.Redirect("rep2resinChart.aspx")
    End Sub

    Private Sub BtnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnView.Click
        Dim rep2DsTOA As DataSet
        Dim ViewYear, data As String
        ViewYear = cmbYear.SelectedItem.Text - indi.year
        data = "datas"
        rep2DsTOA = New DataSet()
        rep2DsTOA = Session(data)
        '
        Dim Datafile As String
        If optDistChosen Then 'by District
            Datafile = "Accidents" & cmbYear.SelectedItem.Text & "Dist" & distNo
        Else 'by State
            Datafile = "Accidents" & cmbYear.SelectedItem.Text
        End If
        '
        DataGrid1.SelectedIndex = -1
        DataGrid1.EditItemIndex = -1
        DataGrid1.DataSource = rep2DsTOA.Tables(Datafile).DefaultView
        DataGrid1.DataBind()

        Dim i As Integer
        For i = 0 To DataGrid1.Items.Count - 1
            DataGrid1.Items(i).Cells(0).Text = GetDay(DataGrid1.Items(i).Cells(0).Text)
        Next i

        NoShow = False
	End Sub

Sub getSession()
			indi = Session("indi")
			caption = Session("caption")
			caption1 = Session("caption1")
			caption2 = Session("caption2")
		 optDistChosen = Session("optDistChosen")
			mutipleYear = Session("mutipleYear")
			YearNum = Session("YearNum")
			NoShow = Session("NoShow")
			connString = Session("connString")
			 distNo = Session("distNo")

		End Sub

End Class

End Namespace
