<!-- #INCLUDE FILE = "Initialize.asp"-->


<%

Sub ValidateUser(vUserID, vPassword)
    Dim ssql 
    
	Set webdb = Server.CreateObject("ADODB.Connection")
	webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=WEBHR;PWD=password;Initial catalog=HRDB_CSEM;Data Source=DESKTOP-5I6A2IP\MSSQLSERVER12;Connect Timeout=900000"
	Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	Set webdbCommand = Server.CreateObject("ADODB.Command")
	'W if vUserID <> "" and vPassword  <> "" THEN
	 'ssql = "Exec sp_sa_selvalidateuser '" +  vUserID + "','" + vPassword + "'"
    'ssql = "Exec sp_sa_selvalidateuser '"+vUserID+"','"+vPassword+"'"
     'ssql = "Exec sp_sa_changepassword1 '"+vUserID+"','"+vPassword+"' 
     'Response.Write ssql	
	ssql = "Exec sp_sa_changepassword1 '"+vUserID+"','"+vPassword+"','','','LOGIN'" 
	
	Set webdbCommand.ActiveConnection = webdb
	webdbCommand.CommandText = ssql
	Response.Write ssql
	webdbRecordset.Open webdbCommand,,1 , 3
    
    Do Until webdbRecordset.EOF
       If webdbRecordset.Fields("RanNo") <> "ERROR" Then
          Session("RanNo") = webdbRecordset.Fields("RanNo")
          Session("Responsibility") = webdbRecordset.Fields("responsibility")
	      Session("EmpID") = webdbRecordset.Fields("empid")
	      Session("EmpName") = webdbRecordset.Fields("empname")
	      Session("Regisno") = webdbRecordset.Fields("regisno")
	      Session("CurrentDate") = webdbRecordset.Fields("date")
	      Session("organname") = webdbRecordset.Fields("organisation")

       'else
        'Response.Redirect "../loginfail.htm"
        'Response.Write  "RanNo = BLANK"
        end if
       webdbRecordset.MoveNext  
    Loop
 end sub   
'W    END IF
'	If Session("RanNo") <> "" Then
'	   Response.Write " Success "
'    Else
'	   Response.Write " Fail "
'    End If

'End Sub

   

%>

