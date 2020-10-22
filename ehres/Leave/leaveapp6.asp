<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->

<%
Response.Buffer = true
Dim ssql
Dim i
Dim colour
Dim count
Dim rowno
Dim ApproveRow
Dim maxrow
Dim rowcount
%>

<html>

<head>
<link rel="stylesheet" type="text/css" HREF="../css/login.css">

<script language="javascript" type="text/javascript">
<!--
var win=null;

function NewWindow(mypage,myname,w,h,scroll,pos)
{

if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=no,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);
if(win.focus){win.focus();}}

// -->
</script>

<title>Leave Approval System</title>

<!--Place this script anywhere in a page.-->
<!--NOTE: You do not need to modify this script.-->

	<script LANGUAGE="JavaScript">

	function Verify()
	{
		msg = "";
		m = true;
		n = true;

m = CheckDate('txtDate1');
		if (!m)
		{
			window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
			document.frmLeaveApproval.txtDate1.value = ""
			document.frmLeaveApproval.txtDate1.focus() 
		}
		
		if (m)
		{
		   m = CheckDate('txtDate2');
		   if (!m)
		      {
			    window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
   			    document.frmLeaveApproval.txtDate2.value = ""
			    document.frmLeaveApproval.txtDate2.focus()

		      }
	    }	
	    
	    if (m)
	    {
		   document.forms[0].submit();
	    }
	}

	function CheckDate(x)
	{
		o = true;

		if ( eval("document.forms[0]." + x + ".value.length == 8") )
		{
			day = eval("document.forms[0]." + x + ".value.substring(0,2)");
			month = eval("document.forms[0]." + x + ".value.substring(2,4)");
			year = eval("document.forms[0]." + x + ".value.substring(4,8)");
			o = o && CheckDay(day, month, year);
			o = o && (month < 13);
		}
		else
		{
		  if ( eval("document.forms[0]." + x + ".value.length == 10") )
		  {
			day = eval("document.forms[0]." + x + ".value.substring(0,2)");
			month = eval("document.forms[0]." + x + ".value.substring(3,5)");
			year = eval("document.forms[0]." + x + ".value.substring(6,10)");
			o = o && CheckDay(day, month, year);
			o = o && (month < 13);
		  }
                  else 
                  { 
 		    if ( eval("document.forms[0]." + x + ".value.length == 0") )
                    {
                        o = true;
                    }
                    else 
                    {
			o = false;
                    }
                  }
		}
	
		if (o) return true;
		else return false;
	}

	function CheckDay(dd, mm, yy)
	{
		MaxDay = new Array (31,28,31,30,31,30,31,31,30,31,30,31);
	
		if (yy%4 == 0) MaxDay[1]++;

		if (dd <= MaxDay[mm-1]) return true;
	}
	
	</script>

</head>
<title>Enquiry - Leave Application</title>
<body bgColor="#ffffff" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<table border="0" width="96%">
  <tr>
    <td width="50%">
      &nbsp;
      <table border="0" width="100%" background="../Image/LeaveApp.gif" height="12%">
        <tr>
          <td width="50%" height="30">&nbsp;<BR></td>
        </tr>
      </table>
    </td>
    <td width="50%"></td>
  </tr>
</table>

<p></P>
<table border="0" width="96%" cellspacing="0" height="121">
  <tr>
    <td width="164%" height="44" colspan="2">
      <table border="0" width="100%" height="1">
        <tr>
          <td width="100%" height="1"><!--'leaveapp3.asp?employeeid=" + Request("employeeid") + "'-->
          <%response.write "<form method='POST' action='leaveapp6.asp?employeeid=" + Request("employeeid") + "' name='frmLeaveApproval'>" %> 
           <!--<%Response.Write "<a href=""leaveapp6.asp?employeeid=" + Request("employeeid") + """>"%>
           <!--<%response.write "<form method='POST' action='leaveapp5.asp?employeeid= write(globalempid)' name='frmLeaveApproval'>" %> -->
           <!--<form method='POST' action='leaveapp5.asp?employeeid= request(employeeid)' name='frmLeaveApproval'>-->  
		   <form method='POST' action='leaveapp6.asp' name='frmLeaveApproval'>
              <td></td>
           
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="small">Employee ID </font>&nbsp;&nbsp;<font class="marineblack"> <%Response.Write Request("employeeid")%></font>
               
              <p>&nbsp;<font class="small">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Status</font>&nbsp;&nbsp;&nbsp;<!--cboStatus=' + options[selectedIndex].value)" action='leaveapp5.asp?-->
              <select size="1" name="cboStatus" style="font-size: 8pt"  > 
              <!--<select size="1" onchange= " if(options[selectedIndex].value) top.location.href=('leaveapp6.asp?cboStatus=' + options[selectedIndex].value )" name="cboStatus" style="font-size: 8pt" >
              <%response.write "<form method='POST' action='leaveapp5.asp?employeeid=" + Request("employeeid") + "' name='frmLeaveApproval'>" %> 
              <!--<select size="1" onchange= " if(options[selectedIndex].value) top.location.href=('leaveapp5.asp?cboStatus=' + options[selectedIndex].value)" name="cboStatus" style="font-size: 8pt" >-->
                               
                  <%  
					dim vStatus
					
					    If request("cboStatus") <> "" Then
					      vStatus = request("cboStatus")
					    ElseIf request.form("cboStatus") = "" Then
					      vStatus = "P"
					    Else
					      vStatus = request.form("cbostatus")
					    End If
					   
					   if vStatus  = "R" then
			             response.write "<Option value = 'A'> Approved </Option>"
			             response.write "<Option value = 'P'> Pending </Option>"
			             response.write "<Option Selected value = 'R'> Rejected </Option>"
			           elseif vStatus  = "A" then
			             response.write "<Option Selected value = 'A'> Approved </Option>"
			             response.write "<Option value = 'P'> Pending </Option>"
			             response.write "<Option value = 'R'> Rejected </Option>"
			           else
			             response.write "<Option value = 'A'> Approved </Option>"
			             response.write "<Option Selected value = 'P'> Pending </Option>"
			             response.write "<Option value = 'R'> Rejected </Option>"
			          end if
				%>
				
			</select>
			
			<% If vStatus = "P" OR vStatus = "A" OR vStatus = "R" Then%>
         
            </select>&nbsp;&nbsp;&nbsp;<font class="small">
              Date Apply For&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="txtDate1" size="9" class="small" <% 
	             if request.form("txtDate1") <> "" then
	                response.write " value='" & request.form("txtDate1") & "'"              
	             end if   
	          %>>&nbsp;&nbsp 
              to&nbsp;&nbsp; <input type="text" name="txtDate2" size="9" class="small" <% 
	             if request.form("txtDate2") <> "" then
	                response.write " value='" & request.form("txtDate2") & "'"              
	             end if   
	          %>>
	          <!--<input type="submit" value="Search" name="cmdRefresh" cmdSearch style="font-size: 8pt">-->
              
              <!--<td></td>-->
             </select>
             &nbsp;&nbsp;<input type="button" value="Search" name="cmdSearch" onClick="Verify()" class="small">
              <%Response.Write "<a href=""leavebalance17.asp?employeeid=" + Request("employeeid") + """>"%>
							
				<font class="marineblue"><u>Back </u></a><BR></font></td>
              <% end if %>
  				
			  <% if vStatus ="P" then %>
				<center>
                <table cellSpacing="0" cellPadding="1" border="0" width="90%" bordercolor="#808080">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
				<tr>
					<!--<td height="20" width="4%" bgcolor="#F3F3F3"></td>--> 
					<td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Date Apply For</font></td>
					<td height="20" width="35%" bgcolor="#F3F3F3"><font class="marineblack">Leave Type</font></td>
					<td height="20" width="8%" bgcolor="#F3F3F3"><font class="marineblack">Period</font></td>
					<!--<td height="20" width="55%" bgcolor="#F3F3F3"><font class="marineblack">Reason</font></td>-->
				</tr>
          
                 <%  
                     dim tempemployeeid          
                     tempemployeeid = ""
 
            		 Set webdb = Server.CreateObject("ADODB.Connection")
   		                 webdb.Open Session("ConnectStr")
  		             Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		             Set webdbCommand = Server.CreateObject("ADODB.Command")
					
					 if tempemployeeid = "" Then
						tempemployeeid = Request("employeeid")
				     end if
					
					if (request("txtDate1") = "" And request("txtDate2") = "") or (request("txtDate1") = "" and request("txtDate2")<> "") or (request("txtDate1") <> "" and request("txtDate2") = "") then
					 ssql = "Exec sp_Wls_LeaveApproval '" + Session("Regisno") + "', '" + Session("EmpID") + "', 'ENG', 'P', '" + request("employeeid") + "','', '01/01/1900', '01/01/1900', 'EMPID'" 
		               'Response.Write SSQL
	    
				    else
					 ssql = "Exec sp_Wls_LeaveApproval '" + Session("Regisno") + "', '" + Session("EmpID") + "', 'ENG', '" _
					         + vStatus + "', '" + request("employeeid") + "','', '" + request("txtDate1") + "', '" _
					         + request("txtDate2") + "', 'EMPID'" 
				      'Response.Write SSQL
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
                        'response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("ApplyFor") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + cstr(webdbRecordset.Fields("leavetype")) + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + cstr(webdbRecordset.Fields("period")) + "</td>"
				        'response.write "<td height='20' align='left'" + colour + "><font class='small'>" + cstr(webdbRecordset.Fields("reason")) + "</td>
				        Response.Write "</tr>"
				       
				        webdbRecordset.MoveNext  
				        count = abs(count - 1)        
			        loop
			        webdbRecordset.close
			        webdb.close  
			 %>
          
          </table>
         </center>  
       <% end if %>  
          
              
              <!--<td></td>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="marineblue">Employee ID :</font>&nbsp;&nbsp;<font class="marineblack"> <%Response.Write request("employeeid")%></font>-->
	         <% IF vStatus = "R" THEN %>
	         <center>
	         <table cellSpacing="0" cellPadding="0" border="0" width="90%" bordercolor="#808080">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
				
				 <tr>
		                <td height="20" width="4%" bgcolor="#F3F3F3"></td>
		                <td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Date Apply For</font></td>
					    <td height="20" width="35%" bgcolor="#F3F3F3"><font class="marineblack">Leave Type</font></td>
					    <td height="20" width="8%" bgcolor="#F3F3F3"><font class="marineblack">Period</font></td>
					    <!--<td height="20" width="55%" bgcolor="#F3F3F3"><font class="marineblack">Reason</font></td>--> 
                    </tr>
                <tr>
                
                <% IF vStatus = "R" then  'or vStatus ="A" 
               
            		 Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		             Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		             Set webdbCommand = Server.CreateObject("ADODB.Command")
                    end if
                  
				    if request("txtDate1") = "" And request("txtDate2") = "" or (request("txtDate1") = "" and request("txtDate2")<> "") or (request("txtDate1") <> "" and request("txtDate2") = "") then
					 ssql = "Exec sp_Wls_LeaveApproval '" + Session("Regisno") + "', '" + Session("EmpID") + "', 'ENG', '" _
					         + vStatus + "','" + request("employeeid") + "','', '01/01/1900', '01/01/1900', 'EMPID'" 
				    else
					 ssql = "Exec sp_Wls_LeaveApproval '" + Session("Regisno") + "', '" + Session("EmpID") + "', 'ENG', '" _
					         + vStatus + "', '" + request("employeeid") + "','', '" + request("txtDate1") + "', '" _
					         + request("txtDate2") + "', 'EMPID'" 
				    end if
				 	'Response.Write ssql			 
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
                        response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("ApplyFor") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavetype") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("period") + "</td>"
				        'response.write "<td height='20' align='left'" + colour + "><font class='small'>" + cstr(webdbRecordset.Fields("reason")) + "</td>
				        Response.Write "</tr>"
				        webdbRecordset.MoveNext  
				        count = abs(count - 1)        
			        loop
			        webdbRecordset.close
			        webdb.close      
			        end if 
			  %>
            </tr>
 	        </table>
 	        </center>
			<% IF vStatus ="A" then %>  <!--or vStatus ="A"-->
			<center>
            <table cellSpacing="0" cellPadding="0" border="0" width="90%" bordercolor="#808080">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
            <!--<table cellSpacing="0" cellPadding="0" border="0" width="100%">-->
			
		             <tr>
		                <td height="20" width="4%" bgcolor="#F3F3F3"></td>
		                <td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Date Apply For</font></td>
					    <td height="20" width="35%" bgcolor="#F3F3F3"><font class="marineblack">Leave Type</font></td>
					    <td height="20" width="8%" bgcolor="#F3F3F3"><font class="marineblack">Period</font></td>
					    <!--<td height="20" width="55%" bgcolor="#F3F3F3"><font class="marineblack">Reason</font></td>--> 
                    </tr>
                <tr>
                
                <% dim tempemployeeid1 
                   tempemployeeid1 = ""
                  IF  vStatus = "A" then  'or vStatus ="A" 
               
            		 Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		             Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		             Set webdbCommand = Server.CreateObject("ADODB.Command")
                    end if
                 
				    if (request("txtDate1") = "" and request("txtDate2") = "") or (request("txtDate1") = "" and request("txtDate2")<> "") or (request("txtDate1") <> "" and request("txtDate2") = "") then
					 ssql = "Exec sp_Wls_LeaveApproval '" + Session("Regisno") + "', '" + Session("EmpID") + "', 'ENG', '" _
					         + vStatus + "', '" + request("employeeid") + "','', '01/01/1900', '01/01/1900', 'EMPID'" 
					 'Response.Write ssql        
				    else
					 ssql = "Exec sp_Wls_LeaveApproval '" + Session("Regisno") + "', '" + Session("EmpID") + "', 'ENG', '" _
					         + vStatus + "', '" + request("employeeid") + "', '', '" + request("txtDate1") + "', '" _
					         + request("txtDate2") + "', 'EMPID'" 
					 'Response.Write ssql        
				    end if
				 	'Response.Write ssql			 
			        
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
                        response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("applyfor") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavetype") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("period") + "</td>"
				        'response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("reason") + "</td>
				        Response.Write "</tr>"
				        webdbRecordset.MoveNext  
				        count = abs(count - 1)        
			        loop
			        webdbRecordset.close
			        webdb.close      
			        end if 
			  %>
            </tr>
 	        </table>
 	        </center>
              <p>&nbsp;</p>
              
            </form>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%" height="21" colspan="2">

    </td>
  </tr>
</table>
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