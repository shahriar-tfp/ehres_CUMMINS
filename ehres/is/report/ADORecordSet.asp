<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<%

Set oConn = Server.CreateObject("ADODB.Connection")
    oConn.Open Session("ConnectStr")
    
Set session("oRs") = Server.CreateObject("ADODB.Recordset")
Set webdbCommand = Server.CreateObject("ADODB.Command")

ssql = "Select 'a' 'empid', 'b' 'empname' from is_emppersonal "
Set webdbCommand.ActiveConnection = oConn
	webdbCommand.CommandText = ssql
    session("oRs").Open webdbCommand, , 1  , 3 

'session("oRs").ActiveConnection = oConn
'session("oRs").Open "Select empid, empname from is_emppersonal order by empname"


reportname = "Report3.rpt"

If Not IsObject (session("oApp")) Then                              
   Set session("oApp") = Server.CreateObject("CrystalRuntime.Application")
End If                                                                

Path = Request.ServerVariables("PATH_TRANSLATED")                     

While (Right(Path, 1) <> "\" And Len(Path) <> 0)                      
  iLen = Len(Path) - 1                                                  
  Path = Left(Path, iLen)                                               
Wend                                                                  
                                                                      
If IsObject(session("oRpt")) then
	Set session("oRpt") = nothing
End if

Set session("oRpt") = session("oApp").OpenReport(path & reportname, 1)

session("oRpt").MorePrintEngineErrorMessages = False
session("oRpt").EnableParameterPrompting = False


session("oRpt").DiscardSavedData
set Database = session("oRpt").Database
set Tables = Database.Tables
set Table1 = Tables.Item(1)
Table1.SetPrivateData 3, session("oRs") 


On Error Resume Next                                                  
session("oRpt").ReadRecords                                           
If Err.Number <> 0 Then                                               
   Response.Write "An Error has occured on the server in attempting to access the data source"
Else

   If IsObject(session("oPageEngine")) Then                              
      set session("oPageEngine") = nothing
   End If

   set session("oPageEngine") = session("oRpt").PageEngine
End If                                                                


viewer = "ActiveX11"

'This line collects the value passed for the viewer to be used, and stores
'it in the "viewer" variable.

If cstr(viewer) = "ActiveX" then
%>
<!-- #include file="SmartViewerActiveX.asp" -->
<%
ElseIf cstr(viewer) = "Netscape Plug-in" then
%>
<!-- #include file="ActiveXPluginViewer.asp" -->
<%
ElseIf cstr(viewer) = "Java using Browser JVM" then
%>
<!-- #include file="SmartViewerJava.asp" -->
<%
ElseIf cstr(viewer) = "Java using Java Plug-in" then
%>
<!-- #include file="JavaPluginViewer.asp" -->
<%
ElseIf cstr(viewer) = "HTML Frame" then
	Response.Redirect("htmstart.asp")
Else
	Response.Redirect("rptserver.asp")
End If
%>