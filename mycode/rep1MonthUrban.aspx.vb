Imports System.Data.OleDb


Namespace Crashsafe

Partial Class rep1MonthUrban
    Inherits System.Web.UI.Page
        Public myDs As New DataSet()
		Dim cstype As Type = Me.GetType()

		 ''global variable

		Dim strscript As String
		Dim indi As StrIndividual
		Dim caption As String
		Dim mutipleYear As Boolean
		Dim YearNum As Integer
		Dim NoShow As Boolean
		Dim connString As String
		
		'local variables
		' Dim optsUrbanChosen As Boolean
		 Dim Area As String

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

	  getSession()

		If NoShow Then
			Dim OleDbConnection As OleDbConnection
			cmbYear.Enabled = False
			cmbYear.Visible = False
			BtnView.Enabled = False
			BtnView.Visible = False
			'

			lblResDis.Text = lblResDis.Text & "---" & Area
			Label1.Text = caption

			Dim fileplace As String
			fileplace = Session(CStr(indi.year) & "Urbanname")
			'connString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\data\tempBlackspot" & indi.year & ".xls" & ";" & "Extended Properties=Excel 8.0;"
			connString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & fileplace & ";" & "Extended Properties=Excel 8.0;"

			OleDbConnection = New OleDbConnection(connString)
			OleDbConnection.Open()
			Dim objCmdSelect As New OleDbCommand("SELECT * FROM  [Sheet1$] ", OleDbConnection)
			Dim timeAdapter As New OleDbDataAdapter()
			timeAdapter.SelectCommand = objCmdSelect
			timeAdapter.Fill(myDs, "[Sheet1$]")

			DataGrid1.DataSource = myDs.Tables("[Sheet1$]").DefaultView
			DataGrid1.DataBind()
			OleDbConnection = Nothing


			NoShow = False
			'
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
		End If
    End Sub
        Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, _
                ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
            If e.Item.ItemIndex = 12 Then

                e.Item.Cells(1).Text = "Total"
            End If
        End Sub

        Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
            Response.Redirect("rep2resinChart.aspx")
        End Sub

        Private Sub BtnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnView.Click
            Dim OleDbConnection As OleDbConnection

            Dim fileplace As String
            fileplace = Session(cmbYear.SelectedItem.Text & "Urbanname")
            'connString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\data\tempBlackspot" & indi.year & ".xls" & ";" & "Extended Properties=Excel 8.0;"
            connString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & fileplace & ";" & "Extended Properties=Excel 8.0;"

            OleDbConnection = New OleDbConnection(connString)
            OleDbConnection.Open()
            Dim objCmdSelect As New OleDbCommand("SELECT * FROM  [Sheet1$] ", OleDbConnection)
            Dim timeAdapter As New OleDbDataAdapter()
            timeAdapter.SelectCommand = objCmdSelect
            timeAdapter.Fill(myDs, "[Sheet1$]")

            DataGrid1.SelectedIndex = -1
            DataGrid1.EditItemIndex = -1
            DataGrid1.DataSource = myDs.Tables("[Sheet1$]").DefaultView
            DataGrid1.DataBind()
            '
            OleDbConnection = Nothing

            NoShow = False
        End Sub


        Protected Sub Button2_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.ServerClick
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('We are sorry,but we do not have the urban crash map for all years!')"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            Exit Sub



            'If Not (mutipleYear) And Not (indi.year = 2006) Then
            '    strscript = "<script language='javascript'>"
            '    strscript = strscript & "alert('We are sorry about that we do not have the crash map of 1999!')"
            '    strscript = strscript & "</script>"
            '    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            '    Exit Sub
            'End If
            'Response.Write("<script language ='javascript'>window.open('Map.aspx?');</script>")
        End Sub

        Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
            Response.Write("<script language ='javascript'> window.opener = window; window.close();</script>")

		End Sub

Sub getSession()
			indi = Session("indi")
			caption = Session("caption")			
			mutipleYear = Session("mutipleYear")
			YearNum = Session("YearNum")
			NoShow = Session("NoShow")
			connString = Session("connString")

		End Sub
    End Class

End Namespace
