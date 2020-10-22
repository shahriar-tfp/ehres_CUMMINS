<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 5.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div align="center">
  <center>
<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%" height="392">
  <TBODY>
  <TR>
    <TD vAlign=top align=center colspan="2" width="936" bgcolor="#0099CC" height="29">
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
    </TD></TR>
  <TR>
    <TD vAlign=top colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></TD></TR>
  <TR>
    <TD vAlign=top align=center width="27" height="109"></TD>
    <TD vAlign=top align=center width="907" height="109"><IMG alt='Main Menu' 
      src="../Image/englsbal.gif" 
    border=0 width="712" height="88"><br>
      &nbsp;</TD></TR>  
  <TR>
    <TD vAlign=top align=center colspan=2 width="936" height="193">
    
    <TABLE cellSpacing=0 width="100%" border=0>
        <TR>
           <TD WIDTH=4%></TD>
           <TD>
           <FONT class=small>Employee ID&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT>
           <select size="1" onchange="if(options[selectedIndex].value) top.location.href=('enq_balance.asp?employeeid=' + options[selectedIndex].value)" name="cboempid" style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px">
            <!--<select name=cboempid style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px" > style="font-size: 8pt" --> 
       			<%  dim tmpEmpID
                'IF tmpEmpID= "" then
					tmpEmpID = Request.Form("cboempid")
                   
       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
         	   ssql ="Exec sp_Wls_selApprovalAuthority '" + Session("Regisno") + "','HRLS','','','" + trim(Session("EmpID")) + "','" + trim(Session("EmpID")) + "','BY_AUTHORITY'"
			  	Set webdbCommand.ActiveConnection = webdb
				
			  	webdbCommand.CommandText = ssql
			  	'Response.Write ssql
			  	'Response.Write "hello"
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			tmpEmpID = request("employeeid") '" "
			 	
			 	'if Request.form("cboempid") = "" then 
	              'response.write "<option selected value='1'>All Subordinate</option>"
	    
	           ' else   
	            '  response.write "<option value='1'>All Subordinate</option>"
	            'end if 
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("empid")) = Request.form("cboempid") ) or ( trim(webdbRecordset.Fields("empid")) = tmpEmpID ) then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
 				    end if
 				   
 				    if tmpEmpID = "" then
					      tmpEmpID = trim(webdbRecordset.Fields("empid"))				
					end if   
			  	 
				   webdbRecordset.MoveNext  
		        loop    
   	        webdbRecordset.close
	        webdb.close      
			%></select>&nbsp;&nbsp;&nbsp; 
           </TD>
        </TR>
        </TABLE>
      <div align="center">
        <center>
      <TABLE cellSpacing=0 cellPadding=0 width=98% border=0 height="1">
        <TBODY>
        <TR><TD HEIGHT=20></TD></TR>
    <tr>
      <td bgcolor="#FFFFFF" width="20">&nbsp;</td>
      <td bgcolor="#F3F3F3" width="160"><font class="marineblack">Leave Type</font></td>
      <td bgcolor="#F3F3F3" width="31"><font class="marineblack">Year</font></td>
      <td align="right" bgcolor="#F3F3F3" width="101"><font class="marineblack">Total Entitlement</font></td>
      <td align="right" bgcolor="#F3F3F3" width="93"><font class="marineblack">Entitle To Date</font></td>
      <td align="right" bgcolor="#F3F3F3" width="75"><font class="marineblack">Leave B/F</font></td>
      <td align="right" bgcolor="#F3F3F3" width="85"><font class="marineblack">Day(s) Taken</font></td>
      <td align="right" bgcolor="#F3F3F3" width="78"><font class="marineblack">Burn Leave</font></td>
      <td align="right" bgcolor="#F3F3F3" width="68"><font class="marineblack">Balance</font></td>
    </tr>

<%    

dim colour
dim count
dim webdbRecordset
dim currentdate
'dim currentdate1

     Set webdb = Server.CreateObject("ADODB.Connection")
     webdb.Open Session("ConnectStr")
	 Set webdbRecordset = Server.CreateObject("ADODB.Recordset") ''" + Session("empid") + "'
	 Set webdbCommand = Server.CreateObject("ADODB.Command")
     '01/08/2005
 		'response.write formatdatetime(now(),vbshortdate)
 		'response.end
 	'add by tey sing yeh on Aug 02 2005
 	datearray = split(session("currentDate"),"/")
	currentdate = datearray(1) & "/" & datearray(0) & "/" & datearray(2)
	currentdate = MonthName(Month(currentdate)) & " " & Day(currentdate) & " " & Year(currentdate)
	'end add by tey sing yeh
	'currentdate1 = formatdatetime(now(),vbshortdate)
	
	'response.write currentdate
	'response.end
	 'ssql = "Exec sp_ls_selAllLeaveBal '" + Session("Regisno") + "','" + tmpEmpID + "', '01/08/2005', '', '0', 'RETRIEVE'"
	 'ssql = "Exec sp_ls_selAllLeaveBal """ + Session("Regisno") + """, """ + tmpEmpID + """, """ + session("currentDate") + """, '', '0', 'RETRIEVE'"   'mark by tey sing yeh
   	 ssql = "Exec sp_ls_selAllLeaveBal """ + Session("Regisno") + """, """ + tmpEmpID + """, """ + session("currentdate") + """, '', '0', 'RETRIEVE'"

	'ssql = "Exec sp_ls_selAllLeaveBal '" + Session("Regisno") + "','" + tmpEmpID + "', '', '', '0', 'RETRIEVE'"   
	 Set webdbCommand.ActiveConnection = webdb
	     webdbCommand.CommandText = ssql
	     webdbRecordset.Open webdbCommand,,1 , 3

     colour = 0
     
      Do Until webdbRecordset.EOF
       
        if count = 1 then
           colour = " bgcolor='#F3F3F3'"
        else
           colour = ""
        end if
                
         response.write "<tr>"
	     response.write "<td height='20'><font class='small'></font></td>"
	     response.write "<td height='20'" & colour & "><font class='small'>" & webdbRecordset.Fields("description") & "</font></td>"
		 response.write "<td height='20'" & colour & "><font class='small'>" & webdbRecordset.Fields("entitlementyear") & "</font></td>"
		 response.write "<td height='20'align='right'" + colour + "><font class='small'>" + webdbRecordset.Fields("actualentitlement") + "</font></td>"
		 response.write "<td height='20'align='right'" + colour + "><font class='small'>" + webdbRecordset.Fields("tolentitlement") + "</font></td>"
		 response.write "<td height='20'align='right'" + colour + "><font class='small'>" + webdbRecordset.Fields("carryforwarddays") + "</font></td>"
		 response.write "<td height='20'align='right'" + colour + "><font class='small'>" + webdbRecordset.Fields("daystakenexcp") + "</font></td>"
		 response.write "<td height='20'align='right'" + colour + "><font class='small'>" + webdbRecordset.Fields("burnleave") + "</font></td>"
		 response.write "<td height='20'align='right'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavebalexcp") + "</font></td></tr>"
		 
        webdbRecordset.MoveNext  
        count = abs(count - 1)
     loop    
	        webdbRecordset.close
			        webdb.close     
 		            
%>

        </TABLE></center>
      </div>
    </TD></TR>
  <TR>
    <TD align=middle colspan=2 width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<BR>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </TD></TR></TBODY></TABLE></center>
</div>
</BODY>