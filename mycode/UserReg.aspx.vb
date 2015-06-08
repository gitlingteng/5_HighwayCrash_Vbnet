'Imports System.Data.SqlClient
Imports System.Data.OleDb


Namespace Crashsafe

Partial Class UserReg
        Inherits System.Web.UI.Page
		Dim cstype As Type = Me.GetType()
		'global variable
		Dim strscript As String
		Dim RightLogin As Integer

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
        Dim rank As Integer

        'Check the Authorization ID first
        If TxtAuth.Text = "" Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Please enter your Authorization ID!' )"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            Exit Sub
        Else
            Dim id As String
            Dim tempSql As String
            Dim Conn As ADODB.Connection
            Dim rst As ADODB.Recordset
            Dim AppPath, DB As String

            AppPath = "C:\data"
            DB = "\userauthid.mdb"
            id = TxtAuth.Text

            Conn = New ADODB.Connection()
            With Conn
                'Telling ADO to use JOLT Here
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Open(AppPath & DB)
            End With

            tempSql = "WHERE [AuthId] = '" & id & "'"
            tempSql = "SELECT  * FROM Authid " & tempSql
            rst = New ADODB.Recordset()
            rst.Open(tempSql, Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            With rst
                If (.RecordCount < 1) Or IsDBNull(.Fields(0).Value) Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('Please enter valid Authorization ID!' )"
                    strscript = strscript & "</script>"
                        ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                        TxtUsername.Text = ""
                        TxtPassword.Text = ""
                        RightLogin = 0
                        Conn.Close()
                        Conn = Nothing
                        rst = Nothing
                        Exit Sub
                    Else
                        rank = .Fields(4).Value ' get the rank of user
                    End If
                End With
            End If

            If Len(TxtUsername.Text) < 4 Or Len(TxtUsername.Text) > 24 Then
                If TxtUsername.Text = "" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('Please enter your Username!' )"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                Else
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The Username should be 4 to 24 characters in length!' )"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    TxtUsername.Text = ""
                    Exit Sub
                End If
            End If

            If TxtEmail.Text = "" Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please enter your Email!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
        If InStr(1, TxtEmail.Text, "@") = 0 Or InStr(1, TxtEmail.Text, ".") = 0 Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Please correct the Email address since it does not appear to be valid.Thanks.' )"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                TxtEmail.Text = ""
                Exit Sub
            End If

            If Len(TxtPassword.Text) < 4 Or Len(TxtPassword.Text) > 14 Then
                If TxtPassword.Text = "" Then
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('Please enter PassWord!' )"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                Else
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('The Password should be 4 to 14 characters in length!' )"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    TxtUsername.Text = ""
                    Exit Sub
                End If
            End If

            If TxtConfirm.Text = "" Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please Confrim your PassWord!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
        If TxtFirstname.Text = "" Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Please enter your First Name!' )"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
            If TxtLastname.Text = "" Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please enter your Last Name!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
        If Len(TxtAdd.Text) < 4 Then
            If TxtAdd.Text = "" Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please enter your Address!' )"
                strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
                Else
                    strscript = "<script language='javascript'>"
                    strscript = strscript & "alert('Your Address is too simple!' )"
                    strscript = strscript & "</script>"
                    ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                    Exit Sub
            End If
        End If
        If TxtCity.Text = "" Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Please enter your City!' )"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
            If CmpState.SelectedItem.Text = "-select state-" Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('Please select your State!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If
        If TxtPassword.Text <> TxtConfirm.Text Then
            strscript = "<script language='javascript'>"
            strscript = strscript & "alert('Password does not match Confirm Password!' )"
            strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

            'save the information to the database
            'Check the user name, it can not be repeat
            If Not (CheckUserName(TxtUsername.Text)) Then
                strscript = "<script language='javascript'>"
                strscript = strscript & "alert('The username has been registered!' )"
                strscript = strscript & "</script>"
                ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
                Exit Sub
            End If

        'Get the time of system
        Dim time As Date
        time = Format(Now(), "Short Date")

        'add information
        Dim Conn2 As OleDbConnection
        Dim tempSql2 As String
        Dim Command As OleDbCommand

        Conn2 = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=C:\Data\user.mdb")

        'tempSql2 = "INSERT  INTO userinfo(AuthId,UserName,Pwd,Email,FirstName,LastName,Address,RegTime,Rank) VALUES" & _
        '             "(@AuthId,@UserName,@Pwd,@Email,@FirstName,@LastName,@Address,@RegTime,@Rank))"
        tempSql2 = "INSERT  INTO userinfo(AuthId,UserName,Pwd,Email,FirstName,LastName,Address,City,State,RegTime,Rank) VALUES" & _
                     "(@AuthId,@UserName,@Pwd,@Email,@FirstName,@LastName,@Address,@City,@State,@RegTime,@Rank)"
        Command = New OleDbCommand(tempSql2, Conn2)

            Command.Parameters.AddWithValue("@AuthId", TxtAuth.Text)
            Command.Parameters.AddWithValue("@UserName", TxtUsername.Text)
            Command.Parameters.AddWithValue("@Pwd", TxtPassword.Text)
            Command.Parameters.AddWithValue("@Email", TxtEmail.Text)
            Command.Parameters.AddWithValue("@FirstName", TxtFirstname.Text)
            Command.Parameters.AddWithValue("@LastName", TxtLastname.Text)
            Command.Parameters.AddWithValue("@Address", TxtAdd.Text)
            Command.Parameters.AddWithValue("@City", TxtCity.Text)
            Command.Parameters.AddWithValue("@State", CmpState.SelectedItem.Text)
            Command.Parameters.AddWithValue("@RegTime", time)
            Command.Parameters.AddWithValue("@Rank", rank)

        Conn2.Open()
        Command.ExecuteNonQuery()
        '********************
        Conn2.Close()
        Conn2 = Nothing
        Command = Nothing
        '*******************
        'finish, show the sucessful window
        Session("UserName") = TxtUsername.Text
        Session("Email") = TxtEmail.Text
        Session("FirstName") = TxtFirstname.Text
        Session("LastName") = TxtLastname.Text
        Session("Address") = TxtAdd.Text
        Session("City") = TxtCity.Text
        Session("State") = CmpState.SelectedItem.Text

        Response.Redirect("UserInfor.aspx")
    End Sub

    Private Function CheckUserName(ByRef username As String) As Boolean
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

        tempSql = " WHERE [UserName] = '" & username & "'"
        tempSql = "SELECT  * FROM userinfo" & tempSql

        rst = New ADODB.Recordset()
        rst.Open(tempSql, Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        With rst
            If (.RecordCount < 1) Or IsDBNull(.Fields(0).Value) Then
                Conn.Close()
                Conn = Nothing
                rst = Nothing
                Return True ' then username is fine
            Else
                Conn.Close()
                Conn = Nothing
                rst = Nothing
                Return False ' the username has been registered
            End If
        End With
    End Function

        Protected Sub TxtAuth_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAuth.TextChanged

        End Sub
    End Class

End Namespace
