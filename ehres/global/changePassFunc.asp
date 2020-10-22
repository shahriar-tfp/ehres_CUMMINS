<!-- #INCLUDE FILE = "ConnectStr.asp"-->
<% dim temppassword %>

<%
function msgBox(strMessage)
   dim strHTML
   strHTML = "<script language=""JavaScript"">"
   strHTML = strHTML & "alert('"& strMessage &"');"
   strHTML = strHTML & "history.go(-1);"
   strHTML = strHTML & "</script>"
   Response.Write strHTML
end function
      
sub NewPassword(vExisPassword,vNewPassword,vConfPassword)
    Set webdb = Server.CreateObject("ADODB.Connection")
	    webdb.Open Session("ConnectStr")
	Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	Set webdbCommand = Server.CreateObject("ADODB.Command")
	
	 'ssql = "Exec sp_sa_changepassword1 '"+ Session("EmpID")+"','"+ vExisPassword +"','"+ vNewPassword +"','"+ vConfPassword +"','CHANGE'"
	ssql = "Exec sp_sa_changepassword1 '"+ Session("EmpID")+"','"+ vExisPassword +"','"+ vNewPassword +"','"+ vConfPassword +"','CHANGE'"
	
	Response.Write ssql
	
	Set webdbCommand.ActiveConnection = webdb
	    webdbCommand.CommandText = ssql
	    'Response.Write ssql
	    webdbRecordset.Open webdbCommand,,1 , 3
     
	'Session("RanNo") = webdbRecordset.Fields("RanNo")
	'temppassword = webdbRecordset.Fields("password")
	'Response.Write (temppassword)
	
	'if temppassword = vNewPassword then
	  ' msgBox("You Have Changed Your Password !")
	'else
	   'Response.Redirect ("http://10.10.10.4/ehres/changepass.asp") 
	'end if      
	
	'Response.Write session("empid")
	'Response.Write(temp)
	'Response.Write "hello"
	'Do Until webdbRecordset.EOF
	
	'If webdbRecordset.Fields("RanNo")<> "ReEnterPass" Then
			'msgBox("Invalid Existing Password !")
    'end if 
    
    'If webdbRecordset.Fields("RanNo") <> "ErrorP" Then
		'Response.Redirect ("http://10.10.10.4/ehres/main.asp")
	'else
	    'Response.Redirect ("http://10.10.10.4/ehres/changepass.asp")   
    'end if
    
    if webdbRecordset.Fields ("RanNo") = "Pass" then
    'If webdbRecordset.Fields("RanNo")= "" Then
          msgBox("You Have Changed Your Password !")
    'Response.Redirect ("http://SSDCSWPDB416/ehres/main.asp")
			'msgBox("You Have successfully changed YOUR PASSWORD!")
    else
		msgBox("You Have successfully changed YOUR PASSWORD!")		
    end if

	   'If webdbRecordset.Fields("RanNo")<> "ErrorP" Then
			'msgBox("You Have successful changed your PASSWORD !")
	   'ELSE
			'msgBox("Please Try Again !")
	   'end if 
	   'webdbRecordset.MoveNext  
    'Loop 
end sub	    
%>


