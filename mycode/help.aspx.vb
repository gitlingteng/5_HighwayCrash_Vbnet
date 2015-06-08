Imports System.IO


Partial Class help
	Inherits System.Web.UI.Page





Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
' ''//transmitfile self buffers	
''Response.Buffer = False
''Response.Clear()
''Response.ClearContent()
''Response.ClearHeaders()
''Response.ContentType = "application/pdf"
''Response.AddHeader("Content-Disposition", "attachment; filename=Highway_Crash_Analysis_User_Guide.pdf")
' ''//transmitfile keeps entire file from loading into memory
''Response.TransmitFile("Highway_Crash_Analysis_User_Guide.pdf")
''Response.Flush()

''Response.Close()
''Response.End()


'Response.Buffer = True
'Response.Clear()
'Response.ClearContent()
'Response.ClearHeaders()
'Response.ContentType = "application/pdf"
'Response.AddHeader("Content-Disposition", "CIC Report")

'Dim fs As New FileStream("C:\Inetpub\wwwroot\Common_SelectBufferTool_VBNet\Highway_Crash_Analysis_User_Guide.pdf", FileMode.Open)
'Dim br As New BinaryReader(fs)

' Dim dataBytes As Byte() = br.ReadBytes((fs.Length - 1))
'Response.BinaryWrite(dataBytes)
'br.Close()
'fs.Close()

'Response.Flush()

'Response.Close()

'Response.End()




End Sub
End Class


