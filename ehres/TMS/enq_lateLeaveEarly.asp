<!-- #INCLUDE FILE = "../global/ConnectStr.asp"-->
<!-- #INCLUDE FILE = "../global/AdoVbs.asp"-->
<!-- #INCLUDE FILE = "../global/inputSession.asp"-->


<% dim connect_string
     connect_string =Session("ConnectStr")
%>
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 5.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

	<SCRIPT LANGUAGE="JavaScript">

	function Verify()
	{
		msg = "";
		m = true;
		n = true;
		document.frmLeaveApproval.txtsearch.value ="Search"
		m = CheckDate('txtDate1');
		if (!m)
		{
			window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
			document.forms[0].reset()
		}
		else
		{
			document.forms[0].submit();
		}
		
		m = CheckDate('txtDate2');
		if (!m)
		{
			window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
			document.forms[0].reset()
		}
		else
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
	
	</SCRIPT>


<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div align="center">
  <center>
<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%" height="392">
  <TBODY>
  <TR>
    <TD vAlign=top align=middle colspan="2" width="936" bgcolor="#0099cc" height="29">
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
       
    </font>
          </td>
          <td width="74%"><font class="marinewhite"> Name :
       <%    
          response.write session("EmpName")
       %>  
      
    </font>
          </td>
        </tr>
      </table>
        </center>
      </div>
    </TD></TR>
  <TR>
    <TD vAlign=top colspan="2" width="100%" height="21" class="small" align="middle">
      <p align="right"><A href="../main.asp"><font color="#000000">Home</font></A>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <A href="../signout.asp"><font color="#000000">Logout</font></A></p></TD></TR>
  <TR>
    <TD vAlign=top align=middle width="27" height="109"></TD>
    <TD vAlign=top align=middle width="907" height="109">
      <P><IMG height=87 src="../Image/engtmslateleave.gif" width=684 border=0 ><BR><BR>
      
      <table  border="0" width ="100%" style="FONT-SIZE: larger">
  <tr><form method="POST" action="enq_lateleaveEarly.asp" name="frmLeaveApproval">
      <p>
      <FONT class="small">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Employee ID&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT>
     <select name=cboempid style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px"  >  
       			<%  dim tmpEmpID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_Wls_selApprovalAuthority '" + Session("Regisno") + "','HRLS','','','" + trim(Session("EmpID")) + "','" + trim(Session("EmpID")) + "','BY_AUTHORITY'"
         	   
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			'tmpEmpID = request("employeeid") '" "
			 	
			 	if Request.form("cboempid") = "" then 
	              response.write "<option selected value='1'>All Subordinate</option>"
	            else   
	              response.write "<option value='1'>All Subordinate</option>"
	            end if 
			 	 
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
				
			%></select><input name =txtsearch type=hidden>&nbsp;&nbsp;&nbsp;
			<p>
			</select>&nbsp;&nbsp;&nbsp;&nbsp;<FONT class="small" >
              Date Apply For&nbsp;&nbsp;&nbsp;<input type="text" name="txtDate1" size="9" class="small"
              <% if request.form("txtdate1")="" then
						response.write " value='" & session("ssDate1lv1") & "'"
					 else
					    response.write " value='" & request.form("txtDate1") & "'"
	                 end if               
                  %>>&nbsp;&nbsp 
              to&nbsp;&nbsp; <input type="text" name="txtDate2" size="9" class="small"
              <% if request.form("txtdate2")="" then
						response.write " value='" & session("ssDate2lv1") & "'"
					 else
					    response.write " value='" & request.form("txtDate2") & "'"
	                 end if               
                  %> >&nbsp;&nbsp;<input type="button" value="Search" name="cmdSearch" onClick="Verify()" onmouseover="this.style.cursor='hand';" class="small" >
      </table> 
      <br></br>
      <TABLE cellSpacing=0 cellPadding=0 width="95%" border=0 height="1">
    <tr>
          <!--<td bgcolor="#ffffff" width="30">&nbsp;</td>-->
	      <td width="8%" bgcolor="#f3f3f3"><font class="marineblack"><b>Date</b></font></td>
	      <td width="10%" bgcolor="#f3f3f3"><font class="marineblack"><b>Employee ID</b></font></td>
	      <td width="10%" bgcolor="#f3f3f3"><font class="marineblack"><b>Name</b></font></td>
	      <td width="8%" bgcolor="#f3f3f3"><font class="marineblack"><b>Department</b></font></td>
	      <td width="8%" bgcolor="#f3f3f3"><font class="marineblack"><b>In</b></font></td>
	      <td width="8%" bgcolor="#f3f3f3"><font class="marineblack"><b>Out</b></font></td>
	      <td width="5%" bgcolor="#f3f3f3"><font class="marineblack"><b>Late</b></font></td>
	      <td width="5%" bgcolor="#f3f3f3"><font class="marineblack"><b>Leave Early</b></font></td>
	      <td width="5%" bgcolor="#f3f3f3"><font class="marineblack"><b>Leave</b></font></td>
	      <td width="5%" bgcolor="#f3f3f3"><font class="marineblack"><b>Period</b></font></td>
	      <!--<td width="5%" bgcolor="#f3f3f3"><font class="marineblack"><b>Leave</b></font></td>
	      <td width="5%" bgcolor="#f3f3f3"><font class="marineblack"><b>Period</b></font></td>-->
	 </tr>
<%  
                     dim sumleave
                     dim sumlateEarly
                               
                     templeaveid = ""
                     
                     'IF  vStatus = "P" then
                     page_size = 10
			
					 if request("page") = "" then
						current_page = 1
					 else
						current_page = CInt(request("page"))
					 end if
            		 
            		 set myconn = server.CreateObject("ADODB.Connection")
			         set rs = server.CreateObject("ADODB.Recordset")
		                 myconn.open connect_string
		                 
		             rs.cursorlocation = adUseClient
		             rs.pagesize = page_size         
            		
					 'if templeaveid = "" Then
					 'templeaveid = Request("leavetypeid")
				     'end if  vtempdate1,vtempdate2,vtempcboempid
				     if Request.Form("txtSearch") = "Search" then
						vtempdate1 = Request.Form("txtdate1") 
						vtempdate2 = Request.Form("txtdate2")
						vtempcboempid = Request.Form("cboempid")
					    call inputleavebal1(vtempdate1,vtempdate2,vtempcboempid)
					 end if
					    
				     vtempcboempid = session("sscboempidlv1")
					 'Response.Write(vtempcboempid)
					 
					 if Request.form("cboempid") = "1" or vtempcboempid ="1" then
						sql = "Exec sp_wtms_rptattleavesummary '" + Session("Regisno") + "','" + Session("empid") + "','" + session("ssdate1lv1") + "','" + session("ssdate2lv1") + "','GRP'"
					 'if Request.form("cboempid") = "1" then
			            'sql = "Exec sp_wtms_rptattleavesummary '" + Session("Regisno") + "','" + Session("empid") + "','" + Request("txtdate1") + "','" + Request("txtdate2") + "','GRP'"
					   Response.Write sql                                                                                           
			         else
						sql = "Exec sp_wtms_rptattleavesummary '" + Session("Regisno") + "','" + session("sscboempid") + "','" + session("ssdate1lv1") + "','" + session("ssdate2lv1") + "','EMP'"
			            'sql = "Exec sp_wtms_rptattleavesummary '" + Session("Regisno") + "','" + request("cboempid") + "','" + Request("txtdate1") + "','" + Request("txtdate2") + "','EMP'"
			            response.Write sql
			         end if
				    
				    rs.Open sql, myconn ,adopenstatic, adLockReadOnly, adCmdText  		      
				    'rs.Open sql, myconn,adopenstatic, adLockReadOnly, adCmdText 
					
					 'date1  = session("ssdate1")
					 'Response.Write(date1)
					 'date2 =session("ssdate2")
					 'Response.Write(date2)
					if rs.PageCount >0 then	 
						page_count = rs.pagecount
					else	
			            page_count =1
			            
			        end if 					
			        'page_count = rs.pagecount
			        
			        if 1 > current_page then current_page = 1
			        if current_page > page_count then current_page = page_count
			        
			        if rs.RecordCount > 0 then
			        tempcount = rs.RecordCount
			        rs.AbsolutePage = current_page
				    sumleave =0
				    sumlateEarly =0
					colour =0
					
					'if rs.RecordCount >0 then  Do Until rs.EOF   '
			        do while rs.AbsolutePage = current_page and not rs.EOF   'Do Until webdbRecordset.EOF
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if 
				      
						'do until rs.EOF
						    sumleave = sumleave + cdbl(rs("late"))
							sumlateEarly = sumlateEarly + cdbl(rs("leavearly"))
						  'rs.MoveNext
					    
						'Response.Write(rs.RecordCount)
						'Response.Write "<br>"
						'Response.Write(i)
						   	
				        response.write "<tr>"
                        'response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + CSTR(rs("date")) + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empid") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empname") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("Department") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("TimeIn") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("TimeOut") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("late") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("leavearly") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("Leave") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("Period") + "</td></tr>"
				        'response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("reason") + "</td>"
				        'response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("reason") + "</td>
				        
				        rs.MoveNext  
				        count = abs(count - 1)        
			        loop
			        
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
			        'if current_page = page_count then
			        Response.Write "<tr>"  
			        response.write "<td height='20'><font class='marineblack'>Total</td>"
			        response.write "<td height='20'></td>"
			        response.write "<td height='20'></td>"
			        response.write "<td height='20'></td>"
			        response.write "<td height='20'></td>"
			        response.write "<td height='20'></td>"
			        response.write "<td height='20'><font class='marineblack'>" + cstr(sumleave) + "</td>"   'late
			        response.write "<td height='20'><font class='marineblack'>" + cstr(sumlateEarly) + "</td>"   'leave Early
			        Response.Write "</tr>"
			        'end if
			    end if    
   
			   %>
     
   </table>
   
  <table cellSpacing="0" cellPadding="1" border="0" width="90%" bordercolor="#808080">
			
		    <p align=center>
			
			<%Response.Write "<br>" 
			Response.Write "<td colspan=""4"" align=""center"">"
  ''''''''''''''''''''''''''''''''''''''''''''''paging records start'''''''''''''''''''''''''''''''''''''''''''''''''
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "First</font><font=""2""> | </font>"
			end if
  
			iF current_page >= 2 then
				Response.Write "<a href=""enq_lateleaveEarly.asp?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""enq_lateleaveEarly.asp?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""enq_lateleaveEarly.asp?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""enq_lateleaveEarly.asp?page="
				Response.Write page_count
				Response.Write """><font face=""Verdana"" & size=""1"">Last >></font></a>" & vbCrlf
			end if 
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Last</font>" & "</font>"
			end if
      ''''''''''''''''''''''''''''''''''''''''paging records end''''''''''''''''''''''''''''''''''''''''''''''''''              
			Response.Write "</center>"%>
		
			<font face=Verdana size=1><center>Page <%=current_page%> of <%=page_count%></center>
			
			</table>
          <p>&nbsp;</p>
      </div>
    </TD></TR>
  <TR>
    <TD align=middle colspan=2 width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<BR>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </TD></TR></TBODY></TABLE></center>
</div>
</BODY>