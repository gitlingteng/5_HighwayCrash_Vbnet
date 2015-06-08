Imports System.data
Imports System.Data.OleDb
Namespace Crashsafe


    Partial Class LoginIn
        Inherits System.Web.UI.Page

        Dim cstype As Type = Me.GetType()
        Dim strscript As String
        Dim RightLogin As Integer


#Region " Web Form Designer Generated Code "

        'This call is required by the Web Form Designer.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents Label3 As System.Web.UI.WebControls.Label
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label


        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: This method call is required by the Web Form Designer
            'Do not modify it using the code editor.
            InitializeComponent()
        End Sub

#End Region

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            'Put user code to initialize the page here
            RightLogin = 0
        End Sub

        Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click
            Dim AppPath As String = "C:\data\user.mdb"
            Dim connString As String = "Provider = Microsoft.Jet.Oledb.4.0; User Id=;Password =;Data Source = " & AppPath
            Dim Myconn As OleDbConnection = New OleDbConnection(connString)
            Dim tempSql As String = ""
            Dim dataReader As OleDbDataReader = Nothing
            Dim myDataset As New DataSet
            Dim rank As Integer
            Dim visitnum As Integer

            Dim user, password As String

            If TxtUsername.Text = "" Or TxtPassword.Text = "" Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please enter the Username and PassWord!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            Else
                user = TxtUsername.Text
                password = TxtPassword.Text
            End If

            On Error GoTo GetTo_Err

            Dim whereClause As String
            whereClause = " WHERE [UserName] = '" & user & "' AND [Pwd]='" & password & "'"
            tempSql = "SELECT  * FROM userinfo" & whereClause

            Myconn.Open()
            Dim Mycommand As OleDbCommand = New OleDbCommand(tempSql, Myconn)
            Dim timeAdapter As OleDbDataAdapter = New OleDbDataAdapter(Mycommand)
            'timeadapter must put before datareader
            timeAdapter.Fill(myDataset, "userinfo")
            dataReader = Mycommand.ExecuteReader


            If dataReader.HasRows Then
                dataReader.Read()
                rank = dataReader.GetValue(10)
                visitnum = dataReader.GetValue(11)

                If rank = 2 Then 'administrator
                    RightLogin = 2
                    Session("succ") = "adm"

                Else 'general user
                    RightLogin = 1

                    Session("succ") = "user"

                End If
                'update the value of field 4 times 

            Else
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Username or PassWord is Invalid!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                TxtUsername.Text = ""
                TxtPassword.Text = ""
                RightLogin = 0

            End If

            dataReader.Close()

            Myconn.Close()

            'calculate total general users visiting number 
            Dim totalvisit As Integer
            totalvisit = gettotalvisitor(connString)

            If rank = 1 Then  'update the number of visitors
                visitnum += 1
                totalvisit += 1
                Dim updatesql As String = "UPDATE userinfo SET Num = @visitnum" & _
                                " WHERE UserName = @user AND Pwd= @password"
                Dim updatecomm As OleDbCommand = New OleDbCommand(updatesql, Myconn)

                updatecomm.Parameters.AddWithValue("@visitnum", visitnum)
                updatecomm.Parameters.AddWithValue("@user", user)
                updatecomm.Parameters.AddWithValue("@password", password)
                Myconn.Open()
                updatecomm.ExecuteNonQuery()

                Myconn.Close()
            End If



            Session("visitnum") = totalvisit
            Session("RightLogin") = RightLogin

            Response.Redirect("options.aspx")




GetTo_Exit:
            'If Not (rst Is Nothing) Then
            '    If (rst.State And ConnectionState.Open) = ConnectionState.Open Then
            '        rst.Close()
            '    End If
            '    rst = Nothing
            'End If
            'If Not (Conn Is Nothing) Then
            '    If (Conn.State And ConnectionState.Open) = ConnectionState.Open Then
            '        Conn.Close()
            '    End If
            '    Conn = Nothing
            'End If
            Exit Sub
GetTo_Err:
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Error" & Err.Description & "')"
            strscript = strscript & "</script>"
            ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            Resume GetTo_Exit
        End Sub
       


        Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
            Response.Write("<script language ='javascript'>window.close('options.aspx?');</script>")
        End Sub

        Private Sub TxtUsername_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtUsername.TextChanged

        End Sub

        Function gettotalvisitor(ByVal connString)
            Dim total As Integer = 0
            Dim whereClause, tempSql As String
            Dim dataReader As OleDbDataReader

            Dim Myconn As OleDbConnection = New OleDbConnection(connString)
            whereClause = " WHERE [Rank] = 1"
            tempSql = "SELECT  * FROM userinfo" & whereClause

            Myconn.Open()
            Dim Mycommand As OleDbCommand = New OleDbCommand(tempSql, Myconn)
            Dim timeAdapter As OleDbDataAdapter = New OleDbDataAdapter(Mycommand)
            'timeadapter.fill  must put before executereader

            dataReader = Mycommand.ExecuteReader()
            If dataReader.HasRows Then
                While dataReader.Read()
                    total = total + dataReader.GetValue(11)
                End While

            End If

            Myconn.close()
            Return total


        End Function
    End Class

End Namespace

