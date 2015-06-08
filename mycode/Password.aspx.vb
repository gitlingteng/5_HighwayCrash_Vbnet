Imports System.Data.OleDb


Namespace Crashsafe

Partial Class Password
        Inherits System.Web.UI.Page
        Dim cstype As Type = Me.GetType()
        Dim strscript As String

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
    End Sub

    Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click
        If TxtUsername.Text = "" Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Please enter your Username!' )"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
            If TxtPas.Text = "" Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please enter old PassWord!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
        If TxtNewPas.Text = "" Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Please enter new PassWord!' )"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
            If TxtConfirm.Text = "" Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please confirm new PassWord!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
        If TxtNewPas.Text <> TxtConfirm.Text Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('New Password does not match Confirm Password!' )"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

            'check for correct password and username
            If CheckUserName(TxtUsername.Text, TxtPas.Text) Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Username or PassWord is Invalid!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

        'reset the password
        ResetPass(TxtUsername.Text, TxtNewPas.Text)
        Response.Redirect("ResPasSuc.aspx")
    End Sub

    Private Sub ResetPass(ByRef username As String, ByRef password As String)
        Dim Conn As OleDbConnection
        Dim tempSql As String
        Dim Command As OleDbCommand

        Conn = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=C:\Data\user.mdb")

        tempSql = "UPDATE  userinfo SET Pwd=@Pwd WHERE UserName=@UserName"
        Command = New OleDbCommand(tempSql, Conn)

            Command.Parameters.AddWithValue("@Pwd", password)
            Command.Parameters.AddWithValue("@UserName", username)

        Conn.Open()
        Command.ExecuteNonQuery()
        '********************
        Conn.Close()
        Conn = Nothing
        Command = Nothing
    End Sub

    Private Function CheckUserName(ByRef username As String, ByRef password As String) As Boolean
        Dim tempSql As String
        Dim Conn As ADODB.Connection
        Dim rst As ADODB.Recordset
        Dim AppPath, DB As String

        AppPath = "C:\data"
        DB = "\user.mdb"
        Conn = New ADODB.Connection()

        With Conn
            'Telling ADO to use JOLT Here
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .Open(AppPath & DB)
        End With

        tempSql = " WHERE [UserName] = '" & username & "' AND [Pwd]='" & password & "'"
        tempSql = "SELECT  * FROM userinfo" & tempSql

        rst = New ADODB.Recordset()
        rst.Open(tempSql, Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            With rst
                If (.RecordCount < 1) Or IsDBNull(.Fields(0).Value) Then
                    Conn.Close()
                    Conn = Nothing
                    rst = Nothing
                    Return True ' then username is not exist
                Else
                    Conn.Close()
                    Conn = Nothing
                    rst = Nothing
                    Return False ' the username is fine
                End If
            End With
    End Function
End Class

End Namespace
