

<!-- #include virtual ="/ehres/global/loginfunc.asp" -->

<%

Response.Buffer = true

Dim ssql
'Dim vUserID
'Dim vPassword
    
'vUserID = Request.Form("txtUserID")
'vPassword = Request.Form("txtPassword")


'Call ValidateUser (vUserID,vPassword) 
Call Sub ValidateUser (Request.Form("txtUserID"),Request.Form("txtPassword")) 

If Session("RanNo") <> "" Then
   Response.Redirect "../main.asp"
Else
   Response.Redirect "../loginfail.htm"
   Response.Write "RanNo = BLANK"
end if

%>