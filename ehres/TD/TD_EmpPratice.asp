<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/AdoVbs.asp"-->
<!-- #include virtual ="/ehres/global/inputSession.asp"-->
<%
	dim Row
	
	row = request.Form("txtaction")

%>
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

<% dim connect_string
       connect_string =Session("ConnectStr")
%>
       
	<SCRIPT LANGUAGE="vbScript">
	
	function Back()
	document.frmEmpPratice.action = "TD_TrainCalender.asp"
	document.frmEmpPratice.submit()
	end function

	function Change()
		document.frmEmpPratice.submit()
	End function
	
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
      <P><IMG height=84 src="../Image/engtraincal.gif" width=683 border=0 ><BR><BR>
        <FORM name=frmEmpPratice action=TD_EmpPratice.asp method=post>
      <TABLE cellSpacing=0 width="100%" border=0>
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <TD>
           <FONT class=small><b>Course Name &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
            <% if request.Form("D" + cstr(row)) = "" then%>
            <input name=txtCourseName type=text readonly value="<%=request.Form("txtCourseName")%>" size=100><input type=hidden name=txtCourseID value="<%=request.Form("txtCourseID")%>">
            <%else%>
            <input name=txtCourseName type=text readonly value="<%=request.Form("D" + cstr(row))%>" size=100 ID="Text1"><input type=hidden name=txtCourseID value="<%=request.Form("C" + cstr(row))%>" ID="Hidden1">
            <%end if%>
           </TD>
        </TR>
        <tr><TD WIDTH=10%>&nbsp;</TD>
        <TD>
       	<FONT class=small><b>Period &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
	    <%if request.Form("P" + cstr(row)) = "" then %>
	    <input type=text name=txtPeriod readonly size=100 value="<%=request.Form("txtPeriod")%>">&nbsp;&nbsp;&nbsp;
	    <%else%>
	    <input type=text name=txtPeriod readonly size=100 value="<%=request.Form("P" + cstr(row))%>" ID="Text2">&nbsp;&nbsp;&nbsp;
	    <%end if%>
           </TD>
        </tr>
        <tr><td width=10%>&nbsp;</td>
        <td><FONT class=small><b>Batch &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
		<select name=cboBatch style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 500px" onchange=change()>
		<%
			dim batchID
			dim courseID
			dim period
			
			courseid=""
			period = ""
			courseid = request.Form("txtCourseID")
			period = left(request.Form("txtPeriod"),10)
			batchID = 1
			batchID = request.Form("cboBatch")
			 
			if batchID = "" then
				batchID = 1
			end if
			
			if courseid = "" then
				courseid = request.Form("C" + cstr(row))
			end if
			
			if period = "" then
				period = left(request.Form("P" + cstr(row)),10)
			end if
			
			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WTD_selDetailSchedule '" + Session("Regisno") + "','" + cstr(courseID) + "','" + cstr(Period) + "'," + cstr(batchID) + ",'tolBatch'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
			  	
	  			if not webdbRecordset.EOF then
	  				dim count
	  				
                    count = 1
                    do until count > webdbRecordset.Fields("totalbatch")
 					if ( cint(count) = cint(Request.form("cbobatch")) ) or ( cint(count) = cint(batchid) )then
   			  	        response.write "<option selected value=" + cstr(count) + ">"  + " " + cstr(count) + "</option>"
					else 
   			  	        response.write "<option value=" + cstr(count) + ">"  + " " + cstr(count) + "</option>"
 				    end if
					count = count + 1  
		        loop       
		        
		        end if
		%></select>&nbsp;&nbsp;&nbsp;
        </td>
        </tr>
        <tr><td width=10%></td>
        <td><FONT class=small><b>Time &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
        <% 
				Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WTD_SelScheduleDetail '" + Session("Regisno") + "','" + cstr(courseID) + "','" + cstr(Period) + "'," + cstr(batchID) 
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
			  	
			  	if not webdbRecordset.EOF then
					response.Write "<input type=text name=txtTime readonly size=30 value='" + webdbRecordset.fields("time") + "' >"
				else
					response.Write "<input type=text name=txtTime readonly size=30 value='' >"
				end if
		%>
		<FONT class=small><b>Venue &nbsp;&nbsp;&nbsp;</b> </FONT>
		<% 
				Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WTD_SelScheduleDetail '" + Session("Regisno") + "','" + cstr(courseID) + "','" + cstr(Period) + "'," + cstr(batchID) 
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
			  	
			  	if not webdbRecordset.EOF then
					response.Write "<input type=text name=txtVenue readonly size=30 value='" + webdbRecordset.fields("venue") +  "' >"
				else
					response.Write "<input type=text name=txtVenue readonly size=30 value='' >"
				end if
		%>
        </td>
        </tr>
        <tr>
        <td>&nbsp;</td>
        </tr>
        
  <TR>
    <TD vAlign=top align=middle colspan=2 width="936" height="193">
      <div align="center">
        <center>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="1">
        <TBODY>
    <tr>
      <td bgcolor="#ffffff" width="30">&nbsp;</td>
      <td bgcolor="#f3f3f3" width="50"><font class="marineblack">EmpID</font></td>
      <td bgcolor="#f3f3f3" width="250"><font class="marineblack">Employee Name</font></td>
      <td bgcolor="#f3f3f3" width="50"><font class="marineblack">Cost Center</font></td>
      <td bgcolor="#f3f3f3" width="50"><font class="marineblack">Job Grade</font></td>
    </tr>

            <%     page_size =10
			
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
		            	
					ssql = "Exec sp_WTD_selTrainingCalendarDetail '" + Session("Regisno") + "','" + courseID + "','" + period + "'," + cstr(batchID) + ",'" + session("EmpID") + "'"
					
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
					 
			        page_count = rs.pagecount
			        
			        if 1 > current_page then current_page = 1
			        if current_page > page_count then current_page = page_count
					
					if rs.RecordCount >0 then
			        rs.AbsolutePage = current_page
			        		        
			        do while rs.AbsolutePage = current_page and not rs.EOF
			        
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("EMPID") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("empname") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("finid") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("jobgradeid") + "</td>"
				        Response.Write "</tr>"
				       
				        rs.Movenext        
			        loop

			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
			      end if  
			       
			%>        
			<tr>
        <td>&nbsp;</td>        
			</tr>
        <tr><td width=10%>&nbsp;</td>
        <td><input name=cmdBack onclick=Back() value=Back type=button ID="Button1"></td>
        </tr>
			 <table cellSpacing="0" cellPadding="1" border="0" width="90%" bordercolor="#808080">
			
		    <p align=center>
			
			<%Response.Write "<br>" 
			Response.Write "<td colspan=""4"" align=""center"">"
  ''''''''''''''''''''''''''''''''''''''''''''''paging records start'''''''''''''''''''''''''''''''''''''''''''''''''
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "First</font><font=""2""> | </font>"
			end if
  
			iF current_page >= 2 then
				Response.Write "<a href=""TD_EmpPratice.asp?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""TD_EmpPratice.asp?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""TD_EmpPratice.asp?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""TD_EmpPratice.asp?page="
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
		
            </TBODY>
            </TABLE></center>
      </div>
    </TD></TR>
    <center>  
    <TD align=center colspan=2 width="936" height="40" class="small"><br>
      <!--&nbsp;<br>
      &nbsp;<BR>--><font class ="small" >Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>.</font></TD></TR></TBODY></TABLE></center>
      </div>
</BODY>




