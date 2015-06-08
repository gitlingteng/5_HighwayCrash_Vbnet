Imports System.IO


Namespace Crashsafe


Partial Class Video
        Inherits System.Web.UI.Page
        Dim cstype As Type = Me.GetType()

        ''used in show video
        Public File(1000) As String
        Public filenum As Integer
        Public filename As String
        Public currentPictureIndex As Integer
        Public videoDirPath As String

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
        'Put user code to initialize the page here\
        'get the information for the map data
        If Label1.Text = "VIDEO DEMO" Then
            Exit Sub
        End If
            'Dim xValue As String
            'Dim yValue As String
            'xValue = Request("X")
            'yValue = Request("Y")

            'xValue = Replace(xValue, "X:", "")
            'yValue = Replace(yValue, "Y:", "")

            'On Error GoTo Proc_Err
            'Dim s As Double
            'Dim minX, minY, maxX, maxY As Double
            's = 0.003
            'minX = xValue - s
            'minY = yValue - s
            'maxX = xValue + s
            'maxY = yValue + s

            '    Dim sServer As String = System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPSERVER")
            '    Dim iPort As Integer = CInt(System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPPORT"))
            '    Dim sService As String = System.Configuration.ConfigurationManager.AppSettings("DEFAULT_MAPSERVICE")
            'Dim conArcIMS As New ESRI.ArcIMS.Server.ServerConnection(sServer, iPort)
            'conArcIMS.ServiceName = sService

            ''get the feature of accidents
            'Dim sAXLText As String
            'sAXLText = "<?xml version=""1.0"" encoding=""UTF-8""?><ARCXML version=""1.1"">"
            'sAXLText = sAXLText & "<REQUEST>"
            'sAXLText = sAXLText & "<GET_FEATURES outputmode=""newxml"" geometry=""false"">"
            'sAXLText = sAXLText & "<LAYER id=""2"" /> "
            'sAXLText = sAXLText & "<SPATIALQUERY subfields=""#ALL#"">"
            'sAXLText = sAXLText & "<SPATIALFILTER relation=""area_intersection"">"
            'sAXLText = sAXLText & "<ENVELOPE minx=""" & minX & """ miny=""" & minY & """ maxx=""" & maxX & """ maxy=""" & maxY & """ />"
            'sAXLText = sAXLText & "</SPATIALFILTER>"
            'sAXLText = sAXLText & "</SPATIALQUERY>"
            'sAXLText = sAXLText & "</GET_FEATURES></REQUEST></ARCXML>"

            'Dim response As String
            'response = conArcIMS.Send(sAXLText, "Query")
            ''
            'Dim axlResponse As New System.Xml.XmlDocument()
            'axlResponse.LoadXml(response)
            'Dim CsectNodelist As System.Xml.XmlNodeList
            'Dim CSECT As String
            'CsectNodelist = axlResponse.GetElementsByTagName("FIELD")
            'If CsectNodelist.Count = 0 Then
            '    strscript = "<script language='javascript'>"
            '    strscript = strscript & "alert('Nothing has been found in the point!!')"
            '    strscript = strscript & "</script>"
            '        ClientScript.RegisterClientScriptBlock(cstype, "Msg", strscript.ToString)
            '    Exit Sub
            '    'Else
            '    Dim a As String
            '    CSECT = CsectNodelist(2).Attributes("value").Value
            '    CSECT = Replace(CSECT, "-", "")
            'End If


        '**********currently is the demo of video
            Label1.Text = "VIDEO DEMO"
            videoDirPath = "c:\Data\000000000001\"
            'For cycle = 1 To 999
            '    File(cycle) = ""
            'Next
            File = Directory.GetFiles(videoDirPath)
            currentPictureIndex = 0
            filename = File(currentPictureIndex)
            filenum = File.Length
Proc_Exit:
        Exit Sub
Proc_Err:
        Resume Proc_Exit
    End Sub

    Private Sub BtOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtOk.Click
        'Dim mstrfileName As String = "c:data\you.wmv"

        'With AxWindowsMediaPlayer1
        '    .Stop()

        '    .FileName = mstrfileName

        '    .Play()
        'End With
    End Sub
End Class

End Namespace
