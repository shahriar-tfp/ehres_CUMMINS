<!-- #INCLUDE FILE = "../global/ConnectStr.asp"-->
<!-- #INCLUDE FILE = "../global/AdoVbs.asp"-->
<!-- #INCLUDE FILE = "../global/inputSession.asp"-->
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 5.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>
<% dim connect_string
       connect_string =Session("ConnectStr")
%>
	<SCRIPT LANGUAGE="JavaScript">

	function Verify()
	{
		msg = "";
		m = true;
		n = true;
		document.frmAttendance.txtsearch.value ="Search"
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
      <P><IMG height=87 src="../Image/engtmserr.gif" width=684 border=0><BR><BR>
      <FORM name=frmAttendance action=enq_atterror.asp method=post>
      </form>
      <TABLE cellSpacing=0 width="100%" border=0>
        <tr>
            <td width ="5%"></td>
			<td>
			<FONT class=small>Employee ID&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT>
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
	  			tmpEmpID = request("employeeid") '" "
			 	
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
				
			%></select>&nbsp;&nbsp;&nbsp; 
			</td>
        </tr>
        <tr><td height="20"></td></tr>
        <TR>
          <td width ="5%"></td>
          <td>
			<font class="small">Type Of Error&nbsp;&nbsp;&nbsp;&nbsp;</font><select size="1" name="cboErrCode" class="small">
                 
                &nbsp;<%    
                    dim vErrCode
					dim selected
					
			          vErrCode = request.form("cboErrCode")
			          selected = false
  
				      Set webdb = Server.CreateObject("ADODB.Connection")
				          webdb.Open Session("ConnectStr")
				      Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
				      Set webdbCommand = Server.CreateObject("ADODB.Command")

				      ssql = "Exec sp_Wtms_selTMSError '', '', 0, '', '', 'DESC'"

				      Set webdbCommand.ActiveConnection = webdb
				          webdbCommand.CommandText = ssql
				          webdbRecordset.Open webdbCommand,,1 , 3
          
					   i = 1
	                    
				       Do Until webdbRecordset.EOF
				          If vErrCode = "" Then
				 	          response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("errcode")) + "'>" + cstr(webdbRecordset.Fields("description")) + "</OPTION> "
							Elseif vErrCode = webdbRecordset.Fields("ErrCode") then
					          response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("errcode")) + "'>" + cstr(webdbRecordset.Fields("description")) + "</OPTION> "
					          selected = true
						 	else   
						       response.write "<OPTION value='" + cstr(webdbRecordset.Fields("errcode")) + "'>" + cstr(webdbRecordset.Fields("description")) + "</OPTION> "
							End If
							i = i + 1
							webdbRecordset.MoveNext
					   loop
				       webdbRecordset.close
				       webdb.close
				  %>

				  <% if selected then %>
                   <option value= "ALL">All Error</option>
				  <% else %>
                   <option selected value= "7">All Error</option>
                  <%end if %> 
                    
                <% response.write (selected)%>
              </select>&nbsp;&nbsp;&nbsp;                   
                  <FONT class=small>Date (ddmmyyyy)</FONT> <INPUT 
                  style="FONT-SIZE: 8pt" size=16 name=txtDate1 
                  <% if Request.Form("txtdate1") = "" then
                       response.write " value='" & session("ssDate1") & "'"
                     else  
                       response.write " value='" & request.form("txtDate1") & "'"
                     end if    
                  %>> <FONT 
                  class=small>to</FONT><INPUT style="FONT-SIZE: 8pt" size=15 
                  name=txtDate2 
                  <% if Request.Form("txtdate2") = "" then
						response.write " value='" & session("ssDate2") & "'"
					 else	
                        response.write " value='" & request.form("txtDate2") & "'"
                     end if    

                  %>><B>&nbsp;&nbsp;<INPUT onmouseover="this.style.cursor='hand';" style="FONT-SIZE: 8pt" onclick=Verify() type=button value=Search name=cmdSearch></B></FORM></TD></TR></TABLE>&nbsp;</P>
      </TD></TR>
  <TR>
    <TD vAlign=top align=middle colspan=2 width="936" height="193">
      <div align="center">
        <center>
      <TABLE cellSpacing=0 cellPadding=0 width="98%" border=0 height="1">
        <TBODY>
    
 <tr>  
    <% If (request("cboErrCode") <> "" And request("cboErrCode") <> "ALL") Then %>
    <!--<% If (error <> "" And error <> "ALL") Then %>-->
    
          <td bgcolor="#ffffff" width="30">&nbsp;</td>
          <td width="100" bgcolor="#f3f3f3"><font class="marineblack"><b>EmpID</b></font></td>
          <td width="100" bgcolor="#f3f3f3"><font class="marineblack"><b>Name</b></font></td>
	      <td width="100" bgcolor="#f3f3f3"><font class="marineblack"><b>Date</b></font></td>
	      <td width="130" bgcolor="#f3f3f3"><font class="marineblack"><b>Date In</b></font></td>
	      <td width="130" bgcolor="#f3f3f3"><font class="marineblack"><b>Date Out</b></font></td>
	      <td width="50" bgcolor="#f3f3f3"><font class="marineblack"><b>Shift In</b></font></td>
	      <td width="50" bgcolor="#f3f3f3"><font class="marineblack"><b>Shift Out</b></font></td>
	      <td width="30" bgcolor="#f3f3f3"><font class="marineblack"><b>Late</b></font></td>
 
	<%Elseif request("cboErrCode") = "ALL" then %>
	
	      <td bgcolor="#ffffff" width="30">&nbsp;</td>
	      <td width="150" bgcolor="#f3f3f3"><font class="marineblack"><b>EmpID</b></font></td>
	      <td width="150" bgcolor="#f3f3f3"><font class="marineblack"><b>Name</b></font></td>
	      <td width="150" bgcolor="#f3f3f3"><font class="marineblack"><b>Type Of Error</b></font></td>
	      <td width="100" bgcolor="#f3f3f3"><font class="marineblack"><b>Date</b></font></td>
	      <td width="130" bgcolor="#f3f3f3"><font class="marineblack"><b>Date In</b></font></td>
	      <td width="130" bgcolor="#f3f3f3"><font class="marineblack"><b>Date Out</b></font></td>
	      <td width="50" bgcolor="#f3f3f3"><font class="marineblack"><b>Shift In</b></font></td>
	      <td width="50" bgcolor="#f3f3f3"><font class="marineblack"><b>Shift Out</b></font></td>
	      <td width="30" bgcolor="#f3f3f3"><font class="marineblack"><b>Late</b></font></td>
    
	<% End If%>
	</tr>
		<% Response.Write Request.Form("cboErrCode")%>	

            <%	   'Set webdb = Server.CreateObject("ADODB.Connection")
   		               'webdb.Open Session("ConnectStr")
  		           'Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           'Set webdbCommand = Server.CreateObject("ADODB.Command")
  		           page_size =10
			
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
  		           
  		           if Request.Form("txtSearch") = "Search" then
  		              vtempempid = Request.Form("cboempid")
  		              vtemptypeerror = request("cboErrCode")
  		              vtempdate1 = Request.Form("txtdate1")
  		              vtempdate2 = Request.Form("txtdate2")
  		              call inputtmserror(vtempempid,vtemptypeerror,vtempdate1,vtempdate2)
  		            end if  'request("cboErrCode")
  		           
  		           Response.Write Request.Form("cboErrCode")
  		           empid = session("ssempid")
  		           error = session("sserror")
  		           date1 = session("ssdate1")
  		           date2 = session("ssdate2")
  		              
  		           'Response.Write (Request.Form("cboErrCode"))
  		           If date1 <> "" And date2 <> "" And error <> "" And error <> "ALL" then
  		           'IF Request("txtDate1") <> "" And Request("txtDate2") <> "" And Request("cboErrCode") <> ""  And Request("cboErrCode") <> "ALL" Then
  		              if empid ="1" OR Request.Form("cboempid") ="1" then
  		              'if Request.Form("cboempid") ="1" then
  						 'ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Session("EmpID") + "', " + request("cboErrCode") + ", '" + request("txtDate1") + "', '" + request("txtDate2") + "', '" + request("cboErrCode") + "',''"
  						 ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Session("EmpID") + "','" + error + "','" + Date1 + "','" + Date2 + "','" + error + "',''"
  						 'Response.Write ssql
  					  else 	 
			             'ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Request.Form("cboempid") + "', " + request("cboErrCode") + ", '" + request("txtDate1") + "', '" + request("txtDate2") + "', '" + request("cboErrCode") + "','EMP'"
			             ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + empid + "', '" + error + "', '" + Date1 + "', '" + Date2 + "', '" + error + "','EMP'"
			             'Response.Write ssql
			          end if   
		           'Elseif Request("txtDate1") <> "" And Request("txtDate2") <> "" And Request("cboErrCode") = "ALL" Then
		           elseif Date1 <> "" And Date2 <> "" And error = "ALL" Then
		              'if Request.Form("cboempid") ="1" then
		              if empid ="1" or Request.Form("cboempid") ="1" then
		                 'ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Session("EmpID") + "', 0 , '" + request("txtDate1") + "', '" + request("txtDate2") + "', '" + request("cboErrCode") + "',''"
		                 ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Session("EmpID") + "', 0 , '" + Date1 + "', '" + Date2 + "', '" + error + "',''"
		                 'Response.Write ssql
		              else   
			             'ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Request.Form("cboempid") + "', 0 , '" + request("txtDate1") + "', '" + request("txtDate2") + "', '" + request("cboErrCode") + "','EMP'"
			             ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + empid + "', 0 , '" + Date1 + "', '" + Date2 + "', '" + error + "','EMP'"
			             'Response.Write ssql
			          end if   
		           'Elseif Request("txtDate1") = "" And Request("txtDate2") = "" And (Request("cboErrCode") = "" or Request("cboErrCode") = "ALL") Then
		           Elseif Date1 = "" And Date2 = "" And error = "" or error = "ALL" Then
		              'if Request.Form("cboempid") ="1" then
		              if empid ="1" or Request.Form("cboempid") ="1" then
		                 'ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Session("EmpID") + "', 0, '', '', 'ALL',''"
		                 ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Session("EmpID") + "', 0, '', '', 'ALL',''"
		                 'Response.Write ssql
		              else   
			             ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + empid + "', 0, '', '', 'ALL','EMP'"
			             'Response.Write ssql
			          end if   
		           'Elseif Request("txtDate1") = "" And Request("txtDate2") = "" And Request("cboErrCode") <> "" And Request("cboErrCode") <> "ALL" Then
		           Elseif Date1 = "" And Date2 = "" And error <> "" And error <> "ALL" Then
		              'if Request.Form("cboempid") ="1" then
		              if empid ="1" or Request.Form("cboempid") ="1" then
			             ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Session("EmpID") + "', 0, '', '', '',''"
			             'ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Session("EmpID") + "', 0, '', '', '',''"			             
			             'Response.Write ssql			           
			          else 
			             'ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + Request.Form("cboempid") + "', 0, '', '', '','EMP'"
			             ssql = "Exec sp_Wtms_selTMSError1 '" + Session("Regisno") + "', '" + empid + "', 0, '', '', '','EMP'"
			             'Response.Write ssql			           
			          end if    
		           End If
  		           
			       'Set webdbCommand.ActiveConnection = webdb
			           'webdbCommand.CommandText = ssql
			           'webdbRecordset.Open webdbCommand,,1 , 3
			           
			       rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
						 
			       page_count = rs.pagecount
			        
			       if 1 > current_page then current_page = 1
			       if current_page > page_count then current_page = page_count
				   
				   'Response.Write (rs.RecordCount) 
				   if rs.RecordCount > 0 then	
			       rs.AbsolutePage = current_page

				   colour = 0
										
			       'Do Until webdbRecordset.EOF
			        do while rs.AbsolutePage = current_page and not rs.EOF
				      if count = 1 then
				         colour = " bgcolor='#eeeeee'"
				      else
				         colour = ""
				      end if
                                          
			          If request("cboErrCode") <> "ALL" and request("cboErrCode") <> "" Then
			          'If (error <> "ALL" and error <> "") or request("cboErrCode") <> "ALL" and request("cboErrCode") <> "" Then
				         response.write "<tr>"
				         response.write "<td></td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empid") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empname") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("date") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("datein") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("dateout") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("shiftin") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("shiftout") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("late") + "</td></tr>"
				      ElseIf request("cboErrCode") = "ALL" or request("cboErrCode") = "" Then
				      'ElseIf (error = "ALL" or error = "") OR request("cboErrCode") = "ALL" or request("cboErrCode") = "" Then
				         response.write "<tr>"
				         response.write "<td></td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empid") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empname") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("description") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("date") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("datein") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("dateout") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("shiftin") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("shiftout") + "</td>"				        
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("late") + "</td></tr>"
				      End If
				  
				      rs.MoveNext  
				      count = abs(count - 1)        
			        loop
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
			        
			        'webdbRecordset.close
			        'webdb.close      
				  end if
				end if  
			 %>
            </TBODY>
            </TABLE></center>
            
            <table cellSpacing="0" cellPadding="1" border="0" width="90%" bordercolor="#808080">
			
		    <p align=center>
			
			<%Response.Write "<br>" 
			Response.Write "<td colspan=""4"" align=""center"">"
  ''''''''''''''''''''''''''''''''''''''''''''''paging records start'''''''''''''''''''''''''''''''''''''''''''''''''
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "First</font><font=""2""> | </font>"
			end if
  
			if current_page >= 2 then
				Response.Write "<a href=""enq_atterror.asp?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""enq_atterror.asp?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""enq_atterror.asp?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""enq_atterror.asp?page="
				Response.Write page_count
				Response.Write """><font face=""Verdana"" & size=""1"">Last >></font></a>" & vbCrlf
			end if 
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Last</font>" & "</font>"
			end if
      ''''''''''''''''''''''''''''''''''''''''paging records end''''''''''''''''''''''''''''''''''''''''''''''''''              
			Response.Write "</center>"%>
		
			<font face=Verdana size=1><center>Page <%=current_page%> of <%=page_count%></font></center>
			
			</table>     
      </div>
    </TD></TR>
  <TR>
    <TD align=middle colspan=2 width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<BR>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </TD></TR></TBODY></TABLE></center>
</div>
</BODY>