<!-- #INCLUDE VIRTUAL = "/ehres/global/loginfunc.asp"-->

<%

Response.Buffer = true

Dim ssql

Call ValidateUser (Request.Form("txtUserID"),Request.Form("txtPassword")) 


If Session("RanNo") <> "" Then
   Response.Redirect "../main.asp"              'main.asp"
Else
   Response.Redirect "../loginfail.htm"
   Response.Write "RanNo = BLANK"
End if

%>