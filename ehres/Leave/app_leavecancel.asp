 <!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<%
'	if Request("txtAction")="DEL" then
	
	   Set webdb = Server.CreateObject("ADODB.Connection")
	   webdb.Open"Provider=SQLOLEDB.1;Persist Security Info=False;UID=HRISMGR;PWD=TIGER;Initial catalog=HRDB_CSEM;Data Source=DESKTOP-3D92T51\MSSQLSERVER2017;Connect Timeout=900000"
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	   Set webdbCommand = Server.CreateObject("ADODB.Command")
	   Set webdbCommand.ActiveConnection = webdb

	   dim maxrow
  	   dim rowcount
	
	   maxrow = Request.Form("txtRowNo")
	   del ="false"
    
       if isnumeric(maxrow) then	
          do until rowcount = cint(maxrow)
	         rowcount = rowcount + 1
   	         if Request.Form("D" + cstr(rowcount)) <> "0" then
			    del = "true"
	      	    ssql = "Exec sp_Wls_delLeaveTransaction '" & Session("Regisno") & "', '" & Session("EmpID") & "' , '" _
	      	            & Request.Form("D"+ cstr(rowcount)) & "', '" & Request.Form("L"+ cstr(rowcount)) & "', '" _
	      	            & Request.Form("P"+ cstr(rowcount)) & "', 'DEL'"
			    webdbCommand.CommandText = ssql
			    webdb.Execute webdbCommand.CommandText
	         end if	
         loop
      end if
      
      if del = "true" then
         response.redirect "app_leavecancel.asp"
      end if

'	end if	
%>	                      
<html><title>eHRES</title>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">

<body bgColor="#ffffff" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<script language="vbscript">
<!--
function ValidateDelData() 
	
	dim rowcount 
	dim ssql
	dim maxrow
	dim action
	dim del
	maxrow = document.frmDelLeave.txtRowNo.value	
	action = ""
    if isnumeric(maxrow) then	
	   do until rowcount = cint(maxrow)
	      rowcount = rowcount + 1
	      ssql="if " + "document.frmDelLeave.C" + cstr(rowcount) + ".checked = false then" + chr(10) 
	      ssql= ssql + " document.frmDelLeave.D" + cstr(rowcount) + ".value=0" + chr(10) 
	      ssql=ssql + "end if"
	
  	      execute ssql
	   loop
	   document.frmDelLeave.submit()
	end if
'	document.frmDelLeave.txtAction.value="DEL"
End function	

function checkAll()
	frmUserAccess.txtAction.value="CHECK"
	document.frmUserAccess.submit()
end function
function uncheckAll()
	frmUserAccess.txtAction.value="UNCHECK"
	document.frmUserAccess.submit()
end function

// -->
</script>


<div align="center">
  <center>
<table cellSpacing="0" cellPadding="0" border="0" width="100%" height="100">
  <tbody>
  <tr>
    <td vAlign="top" align="center" colspan="2" width="936" bgcolor="#0099CC" height="29">
      <div align="center">
        <center>
      <table border="0" width="100%">
        <tr>
          <td width="3%">
          </td>
          <td width="23%">
    <font class="marinewhite">
 Employee ID :
       <%    
          response.write session("EmpID")
       %>  
       
     </font></td>
                  <td width="37%"><font class="marinewhite">Name : <%   '   changePass.asp
          response.write session("EmpName")
                    %>
                    </font></td>
		<td width="37%"><font class="marinewhite">Organisation Name : <%   '   changePass.asp
          response.write session("Organname")
                    %>  
 
      
    </font>
          </td>
        </tr>
      </table>
        </center>
      </div>
    </td></tr>
  <tr>
    <td vAlign="top" colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></td></tr>
  <tr>
    <td vAlign="top" align="center" width="27" height="109"></td>
    <td vAlign="top" align="center" width="907" height="109"><br><img alt="Main Menu" src="../Image/englscancel.gif" border="0" width="696" height="89"><br>
      &nbsp;</td></tr>
  <tr>
    <td vAlign="top" align="center" colspan="2" width="936" height="100%">
      <div align="center">
        <center>
         <form method="POST" action="app_leavecancel.asp" name="frmDelLeave">
          <table border="0" width="100%" bordercolor="#808080" cellSpacing="0" cellPadding="0">
            <tr>
              <td height="20" align="center" width="7%"><font class="marineblack"> </font></td>            
              <td height="20" align="center" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Delete</font></td>
              <td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Date Apply For</font></td>
              <td height="20" width="18%" bgcolor="#F3F3F3"><font class="marineblack">Leave Type</font></td>
              <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Period</font></td>
              <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Status</font></td>              
              <td height="20" width="45%" bgcolor="#F3F3F3"><font class="marineblack">Reason</font></td>
            </tr>

            <%   
            	   Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		           Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           Set webdbCommand = Server.CreateObject("ADODB.Command")

		           ssql = "Exec sp_Wls_selLeaveTransaction """ + Session("Regisno") + """, """ + Session("EmpID") + """, 0, 'ENG', 'C'"
					
				   'Response.Write ssql
			        Set webdbCommand.ActiveConnection = webdb
			            webdbCommand.CommandText = ssql
			            webdbRecordset.Open webdbCommand,,1 , 3

					 colour = 0

			        Do Until webdbRecordset.EOF
					    rowno = rowno + 1			        
					 
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			        
				        response.write "<tr>"
						response.write "<td> </td>" 				        
						response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' name=C" + cstr(rowno) + " value='ON' " + strcheck + "></font></td>" 
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("dateapplyfor") + "<input type='hidden' name=D" + cstr(rowno) + " value= " + webdbRecordset.Fields("dateapplyfor") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavetype") + "<input type='hidden' name=L" + cstr(rowno) + " value= " + webdbRecordset.Fields("leaveid") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("period") + "<input type='hidden' name=P" + cstr(rowno) + " value= " + webdbRecordset.Fields("periodid") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("status") + "</td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("reason") + "</td></tr>"
				        webdbRecordset.MoveNext  
				        count = abs(count - 1)   
			        loop
	  				 response.write "<input type=hidden name=txtRowNo value=" + cstr(rowno) + ">"

			        webdbRecordset.close
			        webdb.close      
			 %>

          </table>
          </td>
        </tr>
        <tr>
          <td height="19"></td>
        </tr>
        <tr>
          <td width="6%" height="19"></td>
          <td width="94%" height="19"><input type="submit" value="Delete" name="cmdDelete" onclick="ValidateDelData()" class="small"></td>
        </tr>
      </table>
    </form>
    <p>&nbsp;</p>

    </td>
  </tr>
</table>
<p>&nbsp;</p>
<table border="0" width="96%">
  <tr>
    <td width="100%" align="center"><img border="0" src="/ehres/Image/dottedlinenav.gif" WIDTH="408" HEIGHT="4"></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<br>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>

<p>&nbsp;</p>
</html>