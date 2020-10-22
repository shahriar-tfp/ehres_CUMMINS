<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/inputSession.asp" -->

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
          <td width="50%" height="18">&nbsp;<BR></td>
        </tr>
      </table>
    </td>
    <td width="50%"></td>
  </tr>
</table>
  <table  border="0" style="FONT-SIZE: larger">
  <tr><form method="POST" action="leavebalance16.asp" name="frmLeaveApproval">
    <td width="100%" height="12%" colspan="2" align="left" style="FONT-FAMILY: ">&nbsp;&nbsp;&nbsp;<FONT class="marineblack">Select Option : </FONT><select size="1" onchange="if(options[selectedIndex].value) top.location.href=('leavebalance17.asp?cboSelect=' + options[selectedIndex].value) " style="font-size: 8pt" name="cboSelect">
        
         <%  
					dim vSelect
			          vSelect = request("cboSelect")
					    
					   if vSelect  = "I" then
			             response.write "<Option  value = 'A'> All Subordinates </Option>"
			             response.write "<Option Selected value = 'I'> Individual </Option>"
			          else
			             response.write "<Option Selected value = 'A'> All Subordinates </Option>"
			             response.write "<Option  value = 'I'> Individual </Option>"
			          end if				    
				  %>
      
      </select>
      
      <%sub sessioncall()
	     templeavetype = request("leavetypeid")
	     tempdate1 = ""
	     tempdate2 =""
	     tempstatus =""
	     call inputsession(templeavetype,tempdate1,tempdate2,tempstatus)
	   end sub%>  
      
      <% If vSelect = "A" Or vSelect ="" Then %> 
      <p><FONT class="marineblack">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Select Leave Type :&nbsp;</FONT>
	   <select name=cboleavetype style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 300px"  >     
      <%  dim tmpEmpID1
          

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_Wls_selleavetype '" + Session("Regisno") + "','BY_Leavetype' , ''"
         		'Response.Write SSQL
         		Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			tmpEmpID1 = Request("leavetypeid")'""

			  	Do Until webdbRecordset.EOF

 					if ( trim(webdbRecordset.Fields("leaveid")) = Request.form("cboleavetype") ) or ( trim(webdbRecordset.Fields("leaveid")) = tmpEmpID1 )  then 'then 'cboleavetype
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("leaveid")) + ">" + trim(webdbRecordset.Fields("description")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("leaveid")) + ">"  + trim(webdbRecordset.Fields("description")) + "</option>"
 				    end if
 				    
 				    if tmpEmpID1 = "" then
					      tmpEmpID1 = trim(webdbRecordset.Fields("leaveid"))
					end if
	              
				   webdbRecordset.MoveNext  
		        loop 
		           	
      %>&nbsp;</select>&nbsp;&nbsp;&nbsp;<input type="submit" value="Search" onclick =<% call sessioncall()%>  name="btnSearch" class="small">
        
       <%IF tmpEmpID1 <> "" OR request("leavetypeid")<> "" THEN
             If Request.form("cboleavetype") <> "" then
			     Response.Write "<a href=""leaveappPG.asp?leavetypeid=" + Request.Form("cboleavetype") + """>"%><font class="marineblue"><u>Leave Application</u></font> </td>
            <%
            else
                 
                 Response.Write "<a href=""leaveappPG.asp?leavetypeid=" + tmpEmpID1 + """>"%><font class="marineblue"><u>Leave Application</u></font>  
			     
			<%
			END IF
			END IF
			%>
		  <% End If%>
		 </table>

<% If vSelect = "A" OR VSELECT = "" Then%>

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
			    <td height="20" width="15%" bgcolor="#f3f3f3"><font class="marineblack"><b>Employee Id</b></font></td>
			    <td height="20" width="20%" bgcolor="#f3f3f3"><font class="marineblack"><b>Employee Name</b></font></td>
			    <td height="20" width="10%" bgcolor="#f3f3f3"><font class="marineblack"><b>Year</b></font></td>
			    <td height="20" width="12%" bgcolor="#f3f3f3"><font class="marineblack"><b>Total Entitlement</b></font></td>
			    <td height="20" width="10%" bgcolor="#f3f3f3"><font class="marineblack"><b>Entitle To Date</b></font></td>
			    <td height="20" width="10%" bgcolor="#f3f3f3"><font class="marineblack"><b>Leave B/F</b></font></td>
			    <td height="20" width="8%" bgcolor="#f3f3f3"><font class="marineblack"><b>Day(s)</b></font></td>
			    <td height="20" width="25%" bgcolor="#f3f3f3"><font class="marineblack"><b>Burn Leave</b></font></td>
			    <td height="20" width="20%" bgcolor="#f3f3f3"><font class="marineblack"><b>Balance</b></font></td>    
            </tr>
        <tr>
           
		<%  IF  Request("cboleavetype")<> "" or request("leavetypeid") <> "" then 'request("leavetypeid")<> ""    THEN
			     'or request("leavetypeid") <> ""
			      Set webdb = Server.CreateObject("ADODB.Connection")
					  webdb.Open Session("ConnectStr")
			      Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			      Set webdbCommand = Server.CreateObject("ADODB.Command")
				 'ssql = "Exec sp_ls_selAllLeaveBal1 """ + Session("Regisno") + """, """ + Session("empid") + """, """ + Session("CurrentDate") + """,""" + Request("cboleavetype") + """,'0','RETRIEVE'"
				 'if request("leavetypeid") <> "" then
				 'or request("cboleavetype") <> ""
				 if Request("leavetypeid") <> ""   then
                 ssql = "Exec sp_ls_selAllLeaveBal1 """ + Session("Regisno") + """, """ + Session("empid") + """, """ + Session("CurrentDate") + """,""" + Request("leavetypeid") + """,'0','RETRIEVE'"
                  'Response.Write ssql    
                 else 
                 ssql = "Exec sp_ls_selAllLeaveBal1 """ + Session("Regisno") + """, """ + Session("empid") + """, """ + Session("CurrentDate") + """,""" + Request("cboleavetype") + """,'0','RETRIEVE'"
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
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("empid1") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("empname") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + cstr(webdbRecordset.Fields("entitlementyear")) + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("actualentitlement") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("tolentitlement") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("carryforwarddays") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("daystakenexcp") + "</td>"				        
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("burnleave") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavebalexcp") + "</td></tr>"
						'response.write "<td align='left'" + colour + "><font class='small'>" +             "<a href=""leavebalance13.asp?employeeid=" + webdbRecordset.Fields("empid")+ ">" + webdbRecordset.Fields("empid") + "</a>"  
						
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
<% End if %>

<p>&nbsp;</p>
<table border="0" width="96%">
  <tr>
    <td width="100%" align="middle"><IMG border=0 height=4 src="/ehres/Image/dottedlinenav.gif" width=408></td>
  </tr>
  <tr>

    <td align="middle" colspan="2" width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<br>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>
</FORM>
</body>

</html>