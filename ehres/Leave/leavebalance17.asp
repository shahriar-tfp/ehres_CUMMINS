<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->

<%
Response.Buffer = true
Dim ssql
Dim i
Dim colour
Dim count
%>

<html>

<head>
<link rel="stylesheet" type="text/css" HREF="../css/login.css">

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Enquiry - Leave Balance</title>
</head>

<body bgcolor="#ffffff" topmargin="0" leftmargin="0">
<table border="0" width="96%">
  <tr>
    <td width="50%">
      &nbsp;
      <table border="0" width="100%" background="../Image/LeaveBal.gif" height="12%">
        <tr>
          <td width="50%" height="22">&nbsp;<BR></td>
        </tr>
      </table>
    </td>
    <td width="50%"></td>
  </tr>
</table>
  <table  border="0" style="FONT-SIZE: larger">
  <tr><form method="POST" action="leavebalance17.asp" name="frmLeaveApproval">
    <td width="100%" height="12%" colspan="2" align="left" style="FONT-FAMILY: ">&nbsp;&nbsp;
    <FONT class="marineblack">&nbsp;Select Option</FONT>
    <!--<FONT 
      class="marineblack style=" style="FONT-SIZE: x-small" large? 
      FONT-SIZE:>Select Option : </FONT>--><select size="1" onchange="if(options[selectedIndex].value) top.location.href=('leavebalance16.asp?cboSelect=' + options[selectedIndex].value) " style="font-size: 8pt" name="cboSelect">
        
         <%  
					dim vSelect
			          vSelect = request("cboSelect")
					  if vSelect ="" then
					     vSelect="I"
					   end if    
					   if vSelect  = "A" then
			             response.write "<Option Selected value = 'A'> All Subordinates </Option>"
			             response.write "<Option value = 'I'> Individual </Option>"
			          else
			             response.write "<Option value = 'A'> All Subordinates </Option>"
			             response.write "<Option Selected value = 'I'> Individual </Option>"
			          end if
			           				    
				  %>
      
      </select>
     <!--<br>--><br>
      <FONT class="marineblack">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Employee ID&nbsp;&nbsp;</FONT>
     <select name=cboempid style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 300px"  >  
       			<%  dim tmpEmpID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_Wls_selApprovalAuthority '" + Session("Regisno") + "','HRLS','','','" + trim(Session("EmpID")) + "','" + trim(Session("EmpID")) + "','BY_AUTHORITY'"
         	   
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			tmpEmpID = request("employeeid") '" "
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("empid")) = Request.form("cboempid") ) or ( trim(webdbRecordset.Fields("empid")) = tmpEmpID )  then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
 				    end if
 				   
 				    if tmpEmpID = "" then
					      tmpEmpID = trim(webdbRecordset.Fields("empid"))				
					end if   
			  	 
				   webdbRecordset.MoveNext  
		        loop       
				
			%></select>&nbsp;&nbsp;&nbsp;<INPUT class=small id=button1 name=button1 style="LEFT: 200px; TOP: 107px" type=submit value=Search>
			
            
            <%IF tmpEmpID <> "" OR request("employeeid")<> "" THEN
             If Request.form("cboempid") <> "" then
			     Response.Write "<a href=""leaveapp6.asp?employeeid=" + Request.Form("cboempid")+ """>"%><font class="marineblue"><u>Leave Application</u></font></td>
            <%                                                                         'cboempid 
            else
                 
                 Response.Write "<a href=""leaveapp6.asp?employeeid=" + tmpEmpID + """>"%><font class="marineblue"><u>Leave Application</u></font>  
			     
			<%
			END IF
			END IF
			%>     
     
      </table>
         
<table border="0" width="96%">  
<tr>
<td width="12%" height="1" colspan="2" align="right">
 </td>
 </tr>
  <TR>
    <td width="2%">&nbsp;</td>
    <td width="95%">
      <table cellSpacing="0" cellPadding="0" border="0" width="100%" bordercolor="#808080">
            <tr>
			    <td height="20" width="30%" bgcolor="#f3f3f3"><font class="marineblack"><b>Leave Type</b></font></td>
			    <td height="20" width="5%" bgcolor="#f3f3f3"><font class="marineblack"><b>Year</b></font></td>
			    <td height="20" width="18%" bgcolor="#f3f3f3"><font class="marineblack"><b>Total Entitlement</b></font></td>
			    <td height="20" width="16%" bgcolor="#f3f3f3"><font class="marineblack"><b>Entitle To Date</b></font></td>
			    <td height="20" width="10%" bgcolor="#f3f3f3"><font class="marineblack"><b>Leave B/F</b></font></td>
			    <td height="20" width="7%" bgcolor="#f3f3f3"><font class="marineblack"><b>Day(s)</b></font></td>
			    <td height="20" width="10%" bgcolor="#f3f3f3"><font class="marineblack"><b>Burn Leave</b></font></td>
			    <td height="20" width="12%" bgcolor="#f3f3f3"><font class="marineblack"><b>Balance</b></font></td>    
            </tr>
        <tr>
                
			
			    <%  
			    tempcbo= request("cboempid")
			   
			    IF  Request("cboempid")<> "" or request("employeeid") <> ""  THEN
			    Set webdb = Server.CreateObject("ADODB.Connection")
					   webdb.Open Session("ConnectStr")
			      Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			      Set webdbCommand = Server.CreateObject("ADODB.Command")

                 if request("cboempid") <> "" then
                 ssql = "Exec sp_ls_selAllLeaveBal """ + Session("Regisno") + """, """ + Request("cboempid") + """, """ + Session("CurrentDate") + """, '', '0', 'RETRIEVE'"
                 'Response.Write ssql 
                 else 
                 ssql = "Exec sp_ls_selAllLeaveBal """ + Session("Regisno") + """, """ + tmpEmpID + """, """ + Session("CurrentDate") + """, '', '0', 'RETRIEVE'"
                   'Response.Write ssql     
                 end if
   
                 Set webdbCommand.ActiveConnection = webdb
                     webdbCommand.CommandText = ssql                                 
                     webdbRecordset.Open webdbCommand,,1 , 3
                     
                     colour = 0
		  	   
                                   Do Until webdbRecordset.EOF
                                      if count = 1 then
                                         colour = " bgcolor='#eeeeee'"
                                      else
                                         colour = ""
                                      end if
                               
                                        response.write "<tr>"                                   
                                        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("description") + "</td>"
                                        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + cstr(webdbRecordset.Fields("entitlementyear")) + "</td>"
                                        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("actualentitlement") + "</td>"
                                        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("tolentitlement") + "</td>"
                                        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("carryforwarddays") + "</td>"
                                        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("daystakenexcp") + "</td>"                                     
                                        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("burnleave") + "</td>"
                                        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavebalexcp") + "</td></tr>"
                                        'response.write "<td align='left'" + colour + "><font class='small'>" +             "<a href=""leavebalance17.asp?employeeid=" + webdbRecordset.Fields("empid")+ ">" + webdbRecordset.Fields("empid") + "</a>"  
                                         
                                        webdbRecordset.MoveNext  
                                        count = abs(count - 1)        
                         loop     
                         end if  
			%>
        </tr>
      </table>
    </td>
    <td width="3%">&nbsp;</td></TR>
<tr>
<td width="12%" height="28" colspan="2" align="right">
 </td>
 </tr>
</table>  

</form>
    </td>
    <td width="3%">&nbsp;</td></TR>
<tr>
<td width="12%" height="28" colspan="2" align="right">
 </td>
 </tr>
</table>  


<p>&nbsp;</p>
<table border="0" width="96%">
  <tr>
    <td width="100%" align="middle"><IMG border=0 height=4 src="/ehres/Image/dottedlinenav.gif" width=408></td>
  </tr>
  <tr>

    <td align="middle" colspan="2" width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<br>&nbsp;<p>Copyright � 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>
</FORM>
</body>

</html>