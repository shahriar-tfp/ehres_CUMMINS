<!-- #INCLUDE FILE = "ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/AdoVbs.asp"-->
<% dim temppassword
   dim connect_string
   dim tempempid
   dim tempuserid
   dim temppass1
   dim temppass2
   dim temppass3
    
   connect_string =Session("ConnectStr")
%>

<% 
	sub updateDb(vtemppass1, vtemppass2,vtemppass3)
	 'If Success = 0 and sucess1 = 0 then 
	 'if strcomp(TRIM(CSTR(vExisPassword)),TRIM(CSTR(temppass)),vbTextCompare)= 0 then 'and (strcomp(TRIM(CSTR(vNewPassword)),TRIM(CSTR(vConfPassword)),vbtextCompare)= 0 ) then 
	   'if ((temppass = vExisPassword) and (vNewPassword = vConfPassword)) then
	   set myconn = server.CreateObject("ADODB.Connection")
       set rs = server.CreateObject("ADODB.Recordset")
           myconn.open connect_string
	   
	       ssql ="Exec sp_sa_changepassword1 '"+ Session("EmpID")+"','"+ vtemppass1 +"','"+ vtemppass2 +"','"+ vtemppass3 +"','CHANGE'"
	       rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText
	       Response.Write ssql
	       do while not rs.eof
	          tempRanNo = rs.Fields("RanNo")
	        rs.MoveNext  
		   count = abs(count - 1)        
		Response.Write(tempRanNo)
		loop
		
		rs.close
		set rs = nothing
		myconn.close
	    set myconn = nothing 
	   
	end sub    
	 %>

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
    
    temppass1 = trim(vExisPassword)
    temppass2 = trim(vNewPassword)
	temppass3 = trim(vConfPassword) 
     
	if not found then
	set myconn = server.CreateObject("ADODB.Connection")
    set rs1 = server.CreateObject("ADODB.Recordset")
        myconn.open connect_string
	
	sql = "Exec sp_wls_selwebpass '"+ Session("EmpID")+"'"
	'Response.Write sql
	rs1.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText
	
	 do while not rs1.EOF 
	 'do until not rs1.EOF 
	        tempuserid = rs1.Fields("userid")
	        temppass =rs1.Fields("password")
	        'tempRanNo = strcomp(trim(cstr(vExisPassword)),trim(cstr(temppass)))   
	        
	       rs1.MoveNext  
		   count = abs(count - 1) 
		   Response.Write(temppass)
		         
		loop
		rs1.close
		set rs1 = nothing
		myconn.close
	    set myconn = nothing 
	    FOUND = true   
	end if
	
	IF FOUND THEN 
	 '  Response.Write(success)
	 '  Response.write(success1)
	  ' if strcomp(vnewPassword ,vConfPassword)=0 and (success =0)then
	     'if tempRanNo=0 then
	     Response.Write (vExisPassword) %> <br> <%
	     Response.Write (temppass)
	     if cstr(vExisPassword) = cstr(temppass) then
	        call updateDb(temppass1,temppass2,temppass3)
            msgBox("You Have Changed Your Password !")
       
         else    
            msgBox(" Your Existing Password OR NEW OR CONFIRM PASSWORD is wrong !")
         end if
     
	     Response.Write (tempRanNo)
	end if 
	'END IF
	end sub
	%>
	
	
	