 Imports System.IO
Partial Class helpdoc
    Inherits System.Web.UI.Page
	

Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
Response.Clear()
Response.ClearContent()
Response.ClearHeaders()
Response.ContentType = "application/ms-word"
Response.AddHeader("Content-Disposition", "inline; filename=Highway_Crash_Analysis_Rough_Draft.doc")

        Dim fs As New FileStream("C:\Inetpub\wwwroot\6Crashbyurban\Highway_Crash_Analysis_Rough_Draft.doc", FileMode.Open)
Dim br As New BinaryReader(fs)

 Dim dataBytes As Byte() = br.ReadBytes((fs.Length - 1))
Response.BinaryWrite(dataBytes)
br.Close()
fs.Close()

Response.Flush()

Response.Close()

Response.End()
End Sub
End Class
