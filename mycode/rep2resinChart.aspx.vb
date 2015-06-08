Namespace Crashsafe

Partial Class rep2resinChart
    Inherits System.Web.UI.Page
    Protected WithEvents chartHolder1 As System.Web.UI.WebControls.PlaceHolder
        Protected WithEvents btnResinExcel As System.Web.UI.WebControls.Button
        Dim cstype As Type = Me.GetType()

        Dim strscript As String

        'original global variables
        Dim byMonthChosen As Boolean
        Dim byDayChosen As Boolean
        Dim QueryType As Integer
        Dim crashrate As Boolean

        Dim optDistChosen As Boolean
        Dim optstateChosen As Boolean
        Dim optRateChosen As Boolean

        Dim byCollChosen As Boolean
        Dim byAccChosen As Boolean
        Dim byPt_ImpactChosen As Boolean

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
        'Put user code to initialize the page here

            getSession()
            Dim relPath As String = Session("relPath")
        Dim strImageTag As String = "<IMG SRC='" + relPath + "'/>"
        chartHolder.Controls.Add(New LiteralControl(strImageTag))
        If QueryType = 7 Or QueryType = 4 Or (QueryType = 3 And Not (optRateChosen)) Then 'if show the blackspot(top number and crash rate),or particularhighway, or partition window
            BtnPic.Visible = True
                BtnPic.Enabled = True

        End If
        If showlargechar Then
            BtnCancel.Disabled = False 'run at sever control
        End If
    End Sub

    Private Sub btnResinXL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Response.Redirect("rep2Result.aspx")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Response.Redirect("options.aspx")
    End Sub

    Private Sub Button3_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.ServerClick
      
        If Not (mutipleYear) And indi.year = 1999 Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('We are sorry about that we do not have the crash map of 1999!')"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            Exit Sub
		End If

If QueryType = 0 Then
		strscript = "<script language='javascript'>"
			strscript = strscript & "alert('We are sorry  that we do not have the crash map of urban area!')"
			strscript = strscript & "</script>"
				ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
			Exit Sub
End If

            Response.Write("<script language ='javascript'>window.open('Map.aspx?');</script>")
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        NoShow = True
        '1=By time,2=By Crash Characteristics,3=blackspot, 4=by partial highway,5=compound query
        If (QueryType = 1) Then '1=By time
            If byMonthChosen Then
                Response.Redirect("rep1Month.aspx")
            ElseIf byDayChosen Then
                Response.Redirect("rep1Day.aspx")
            Else
                Response.Redirect("rep1Hour.aspx")
            End If
        ElseIf (QueryType = 2) Then 'by blackspot
            If byCollChosen Then
                Response.Redirect("rep2TypeColl2.aspx")
            ElseIf byAccChosen Then
                Response.Redirect("rep2TypeAcc.aspx")
            ElseIf byPt_ImpactChosen Then
                Response.Redirect("rep2POI.aspx")
            Else
                Response.Redirect("rep2Viol1.aspx")
            End If
        ElseIf (QueryType = 3) Then 'by blackspot
            If optRateChosen Then
                Response.Redirect("repBlackspot2.aspx")
            Else
                Response.Redirect("repBlackspot.aspx")
            End If
        ElseIf (QueryType = 4) Then 'by partial highway
            Response.Redirect("repSpcHway.aspx")
        ElseIf (QueryType = 5) Then 'by Compound  highway
            Response.Redirect("repComQry.aspx")
        ElseIf (QueryType = 7) Then 'by partition window
            Response.Redirect("repPartiWin.aspx")
        ElseIf (QueryType = 8) Then 'by partition window
            Response.Redirect("repIntel.aspx")
        ElseIf (QueryType = 0) Then ' by urban and time
            If byMonthChosen Then
                Response.Redirect("rep1MonthUrban.aspx")
            ElseIf byDayChosen Then
                Response.Redirect("rep1DayUrban.aspx")
            Else
                Response.Redirect("rep1HourUrban.aspx")
            End If
        End If
    End Sub

        Private Sub BtnPic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPic.Click

            crashrate = Session("crashrate")
            If crashrate Then 'if want to show number of crashes
                Dim relPath As String = Session("relPath2")
                Dim strImageTag As String = "<IMG SRC='" + relPath + "'/>"
                chartHolder.Controls.Clear()
                chartHolder.Controls.Add(New LiteralControl(strImageTag))
                If QueryType = 7 Then
                    BtnSum.Visible = True
                    BtnSum.Enabled = True
                End If
                crashrate = False
                Session("crashrate") = crashrate
                BtnPic.Text = "Crash rate"

            Else
                Dim relPath As String = Session("relPath")
                Dim strImageTag As String = "<IMG SRC='" + relPath + "'/>"
                chartHolder.Controls.Clear()
                chartHolder.Controls.Add(New LiteralControl(strImageTag))
                crashrate = True
                BtnPic.Text = "Number of crashes"
                BtnSum.Visible = False
                BtnSum.Enabled = False
                Session("crashrate") = crashrate

            End If
        End Sub

        Protected Sub BtnSum_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnSum.Click
            Dim relPath As String = Session("relPath3")
            Dim strImageTag As String = "<IMG SRC='" + relPath + "'/>"
            chartHolder.Controls.Clear()
            chartHolder.Controls.Add(New LiteralControl(strImageTag))

        End Sub

        Sub getSession()
            byMonthChosen = Session("byMonthChosen")
            byDayChosen = Session("byDayChosen")
            QueryType = Session("QueryType")
            'crashrate = Session("crashrate")
            optRateChosen = Session("optRateChosen")
            byCollChosen = Session("byCollChosen")
            
            byAccChosen = Session("byAccChosen")
            byPt_ImpactChosen = Session("byPt_ImpactChosen")

          
        End Sub

			Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
			   Response.Write("<script language ='javascript'> window.opener = window; window.close();</script>")

		End Sub
	End Class

End Namespace
