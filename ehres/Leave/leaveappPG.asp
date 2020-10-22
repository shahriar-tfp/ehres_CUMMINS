<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/AdoVbs.asp"-->
<!-- #include virtual ="/ehres/global/inputSession.asp"-->

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
Dim page_size
Dim current_page
Dim myconn
Dim rs
Dim sql
Dim page_count
Dim connect_string

connect_string =Session("ConnectStr")
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
     document.frmLeaveApproval1.txttempleave.value ="Search"
     m = CheckDate('txtDate1');
		if (!m)
		{
			window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
			document.frmLeaveApproval1.txtDate1.value = ""
			document.frmLeaveApproval1.txtDate1.focus() 
		}
		
		if (m)
		{
		   m = CheckDate('txtDate2');
		   if (!m)
		      {
			    window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
   			    document.frmLeaveApproval1.txtDate2.value = ""
			    document.frmLeaveApproval1.txtDate2.focus()

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
          <%if session("ssleavetypeid") = "" then
				response.write "<form method='POST' action='leaveappPG.asp?leavetypeid=" + Request("leavetypeid") + "' name='frmLeaveApproval1'>"
            else 
				response.write "<form method='POST' action='leaveappPG.asp?leavetypeid=" + Session("ssleavetypeid") + "' name='frmLeaveApproval1'>"
            end if%> 
           <!--<%Response.Write "<a href=""leaveapp6.asp?employeeid=" + Request("employeeid") + """>"%>
           <!--<%response.write "<form method='POST' action='leaveapp5.asp?employeeid= write(globalempid)' name='frmLeaveApproval'>" %> -->
           <!--<form method='POST' action='leaveapp5.asp?employeeid= request(employeeid)' name='frmLeaveApproval'>-->  
		   <form method='POST' action='leaveappPG.asp' name='frmLeaveApproval1'>
              <td></td>
           
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="marineblue">Leave ID :</font>&nbsp;&nbsp;<font class="marineblack"> <!--<%Response.Write Request("leavetypeid")%></font>-->
              <%
                  set myconn = server.CreateObject("ADODB.Connection")
			      set rs1 = server.CreateObject("ADODB.Recordset")
		              myconn.open connect_string
				  
                   if session("ssleavetypeid") = "" then          
						sql="Exec sp_Wls_selleavetype '" + Session("Regisno") + "','BY_Leavedesc',""" + Request("leavetypeid") + """"
			       else 
			            sql="Exec sp_Wls_selleavetype '" + Session("Regisno") + "','BY_Leavedesc',""" + Session("ssleavetypeid") + """" 	
			       end if      		

                   rs1.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 	
                   
					response.write rs1("description")
                  %></font>
               
              <p>&nbsp;<font class="small">&nbsp;&nbsp;&nbsp;&nbsp; Status</font>&nbsp;&nbsp;&nbsp;&nbsp;<!--cboStatus=' + options[selectedIndex].value)" action='leaveapp5.asp?-->
              <select size="1" name="cboStatus" style="font-size: 8pt"  > 
              <!--<select size="1" onchange= " if(options[selectedIndex].value) top.location.href=('leaveapp6.asp?cboStatus=' + options[selectedIndex].value )" name="cboStatus" style="font-size: 8pt" >
              <%response.write "<form method='POST' action='leaveapp5.asp?employeeid=" + Request("employeeid") + "' name='frmLeaveApproval'>" %> 
              <!--<select size="1" onchange= " if(options[selectedIndex].value) top.location.href=('leaveapp5.asp?cboStatus=' + options[selectedIndex].value)" name="cboStatus" style="font-size: 8pt" >-->
                               
                  <%  
					dim vStatus
					    tempstatus= session("ssStatus")
					    If request("cboStatus") <> "" Then
					       vStatus = request("cboStatus")
					    elseIf request.form("cboStatus") ="" and tempstatus="p" Then 
						   vStatus ="P"   					    
					    Else
					      vStatus=tempstatus
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
			
			<% If vStatus ="P" OR vStatus ="A" OR vStatus ="R" Then%>
         
            </select>&nbsp;&nbsp;&nbsp;&nbsp;<font class="small">
              Date Apply For&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="txtDate1" size="9" class="small" <% 
	              tempssdate1=session("ssdate1")
	             if request.form("txtDate1") <> "" or tempssdate1 ="01/01/1900" then
	                response.write " value='" & request.form("txtDate1") & "'"
	             else 
	                response.write " value='" & session("ssDate1") & "'"                
	             end if   
	          %>>&nbsp;&nbsp 
              to&nbsp;&nbsp; <input type="text" name="txtDate2" size="9" class="small" <% 
	             tempssdate2=session("ssdate2")
                
	             if request.form("txtDate2") <> "" or tempssdate2 ="01/01/1900" then
	                response.write " value='" & request.form("txtDate2") & "'" 
	              else 
	                response.write " value='" & session("ssDate2") & "'"                   
	             end if   
	             	                 
	          %>>
	          
             </select><input type="hidden" name="txttempleave" size="9" class="small">
             &nbsp;&nbsp;<input type="button" value="Search" name="cmdSearch" onClick="Verify()" class="small">
              <%Response.Write "<a href=""leavebalance16.asp?leavetypeid=" + Request("leavetypeid") + """>"%>					
				<font class="marineblue"><u>Back </u></a><BR></font></td>
              <% end if %>
   			  
			  <% if request("txttempleave") ="Search" or session("ssleavetypeid") ="" then
			     templeavetypeid = request("leavetypeid")
				 tempdate1 = request("txtDate1")
				 tempdate2 = request("txtDate2")
				 tempstatus = request("cboStatus")
			     call inputsession(templeavetypeid,tempdate1,tempdate2,tempstatus)
			    end if
			   
			  %>   
			  
			             				
			  <% if vStatus ="P"  then %>
				
                <table cellSpacing="0" cellPadding="0" border="0" width="120%" bordercolor="#808080">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
				<tr>
					<td height="20" width="4%" bgcolor="#F3F3F3"></td> 
					<td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Emp ID</font></td>
					<td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Name</font></td>
					<td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Date Apply For</font></td>
					<td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Period</font></td>
					<td height="20" width="55%" bgcolor="#F3F3F3"><font class="marineblack">Reason</font></td>
				</tr>
          
                 <%  
                     dim templeaveid          
                     templeaveid = ""
                     'IF  vStatus = "P" then
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
            		
					 if templeaveid = "" Then
						templeaveid = Request("leavetypeid")
				     end if
					
					 if session("ssleavetypeid")="" and request("txttempleave") ="" then 
			            sql = "Exec sp_Wls_selLeaveTransaction2 """ + Session("Regisno") + """,""" + Session("empid") + """,""" + Request("leavetypeid") + """,'01/01/1900','01/01/1900','ENG','P'"
					    'esponse.Write sql                                                                                           
			         else
			            
			            sql = "Exec sp_Wls_selLeaveTransaction2 """ + Session("Regisno") + """,""" + Session("empid") + """,""" + Session("ssleavetypeid") + """,""" + Session("ssDate1") + """,""" + Session("ssDate2") + """,'ENG',""" + Session("ssStatus") + """"  'p'"  '""" + Session("ssStatus") + """"   ''p'"      ''01/01/1900','01/01/1900','ENG','p'"    '""" + Session("ssStatus") + """"     'p'"
			            'esponse.Write sql
			          end if
				     		      
				    rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
				
			        page_count = rs.pagecount
			        
			        if 1 > current_page then current_page = 1
			        if current_page > page_count then current_page = page_count
			        
			        rs.AbsolutePage = current_page
				    
					 colour = 0
					
			        do while rs.AbsolutePage = current_page and not rs.EOF
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if

				        response.write "<tr>"
                        response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empid") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empname") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("dateapplyfor") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("period") + "</td></tr>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("reason") + "</td></tr>"
				        rs.MoveNext  
				        count = abs(count - 1)        
			        loop
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing 
   
			   %>
          
          </table>
          
       <% end if %>  
     
	   <% IF vStatus = "R" THEN %>  <!--'or vStatus = "A" -->
	         <table cellSpacing="0" cellPadding="0" border="0" width="100%" bordercolor="#808080">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
				
				 <tr>
		                <td height="20" width="4%" bgcolor="#F3F3F3"></td>
		                <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Emp ID</font></td>
					    <td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Name</font></td>
					    <td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Date Apply For</font></td>
					    <td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Period</font></td> 
					    <td height="20" width="55%" bgcolor="#F3F3F3"><font class="marineblack">Reason</font></td>
                    </tr>
                <tr>
                
                <% 
                  IF vStatus = "R"  then   
               
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
               
            		
                    end if
                 

				    if session("ssleavetypeid")="" then  
			           sql = "Exec sp_Wls_selLeaveTransaction2 """ + Session("Regisno") + """,""" + Session("empid") + """,""" + Request("leavetypeid") + """,'01/01/1900','01/01/1900','ENG',""" + request("cboStatus") + """"                                                                                            '" + request("txtDate1") + "'
			        else
			           sql = "Exec sp_Wls_selLeaveTransaction2 """ + Session("Regisno") + """,""" + Session("empid") + """,""" + Session("ssleavetypeid") + """,""" + Session("ssdate1") + """,""" + Session("ssdate2") + """,'ENG',""" + Session("ssStatus") + """"
			        end if
			        					
					rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
						 
			        page_count = rs.pagecount
			        
			        if 1 > current_page then current_page = 1
			        if current_page > page_count then current_page = page_count
			        
			        rs.AbsolutePage = current_page
			       
					 colour = 0

                    do while rs.AbsolutePage = current_page and not rs.EOF
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			        
				        response.write "<tr>"
                        response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empid") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empname") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("dateapplyfor") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("period") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("reason") + "</td></tr>"
				        
				        rs.MoveNext  
				        count = abs(count - 1)        
			        loop
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing      
			        end if 
			  %>
            </tr>
 	        </table>
 	        
 	        <% IF vStatus = "A" THEN %>
	         <table cellSpacing="0" cellPadding="0" border="0" width="100%" bordercolor="#808080">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
				
				 <tr>
		                <td height="20" width="4%" bgcolor="#F3F3F3"></td>
		                <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Emp ID</font></td>
					    <td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Name</font></td>
					    <td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Date Apply For</font></td>
					    <td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Period</font></td> 
					    <td height="20" width="55%" bgcolor="#F3F3F3"><font class="marineblack">Reason</font></td>
                    </tr>
                <tr>
                
                <% 
                  IF vStatus = "A" then  
               
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
               
            		
                    end if
                  

				    if session("ssleavetypeid")="" then
			           sql = "Exec sp_Wls_selLeaveTransaction2 """ + Session("Regisno") + """,""" + Session("empid") + """,""" + Request("leavetypeid") + """,'01/01/1900','01/01/1900','ENG',""" + request("cboStatus") + """"
					                                                                                          
			        else
			           sql = "Exec sp_Wls_selLeaveTransaction2 """ + Session("Regisno") + """,""" + Session("empid") + """,""" + Session("ssleavetypeid") + """,""" + Session("ssdate1") + """,""" + Session("ssdate2") + """,'ENG',""" + Session("ssStatus") + """"
			          
			        end if
			        					
					rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
						 
			        page_count = rs.pagecount
			        
			        if 1 > current_page then current_page = 1
			        if current_page > page_count then current_page = page_count
			        
			        rs.AbsolutePage = current_page
			        
					 colour = 0

                    do while rs.AbsolutePage = current_page and not rs.EOF
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			        
				        response.write "<tr>"
                        response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empid") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("empname") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("dateapplyfor") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("period") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("reason") + "</td></tr>"
				        
				        rs.MoveNext  
				        count = abs(count - 1)        
			        loop
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing      
			        end if 
			  %>
            </tr>
 	        </table>
 	        
			<table cellSpacing="0" cellPadding="1" border="0" width="90%" bordercolor="#808080">
			<% if vStatus ="P" or vStatus="A" or vStatus="R" then%>
		    <p align=center>
			
			<%Response.Write "<br>" 
			Response.Write "<td colspan=""4"" align=""center"">"
  ''''''''''''''''''''''''''''''''''''''''''''''paging records start'''''''''''''''''''''''''''''''''''''''''''''''''
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "First</font><font=""2""> | </font>"
			end if
  
			iF current_page >= 2 then
				Response.Write "<a href=""leaveappPG.asp?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""leaveappPG.asp?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""leaveappPG.asp?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""leaveappPG.asp?page="
				Response.Write page_count
				Response.Write """><font face=""Verdana"" & size=""1"">Last >></font></a>" & vbCrlf
			end if 
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Last</font>" & "</font>"
			end if
      ''''''''''''''''''''''''''''''''''''''''paging records end''''''''''''''''''''''''''''''''''''''''''''''''''              
			Response.Write "</center>"%>
		
			<font face=Verdana size=1><center>Page <%=current_page%> of <%=page_count%></center>
			<% end if %>
			</table>
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