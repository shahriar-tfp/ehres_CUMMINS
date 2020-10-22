<!-- #INCLUDE FILE = "../global/ConnectStr.asp"-->
<!-- #INCLUDE FILE = "../global/AdoVbs.asp"-->
<!-- #INCLUDE FILE = "../global/inputSession.asp"-->
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

<% dim connect_string
       connect_string =Session("ConnectStr")
%>
       
	<SCRIPT LANGUAGE="VBScript">

	function Change()
		document.frmTranCalendar.submit()
	end function
	
	function validate(X)
	document.frmEmpPratice.txtaction.value = X
	document.frmEmpPratice.submit()
	end function
	
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
      <TABLE cellSpacing=0 width="100%" border=0>
      <FORM name=frmTranCalendar action=TD_TrainCalender.asp method=post>
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
        </TR>
        <tr>
        <td width=10%>&nbsp;</td>
        <td><FONT class=small><b>Month&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
        <select name=cboMonth style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 100px"  onchange=Change()>
        <% 
			dim month1 
			
			month1 = ""
			month1 = request.Form("CboMonth")
			
			if month1 = "" then
				 Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="select datepart(month,dateadd(month,1,getdate())) as 'currentmonth'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
			  	
			  	if not webdbRecordset.EOF then
			  		month1 = webdbRecordset.Fields("currentmonth")
			  	end if
			end if
			
			if request.Form("CboMonth")= 1 or month1 = 1 then
				response.Write "<option selected value=1>January</option>"
			else
				response.Write "<option value=1>January</option>"
			end if
			
			if request.Form("CboMonth") = 2 or month1 = 2 then
				response.Write "<option selected value=2>February</option>"
			else
				response.Write "<option value=2>February</option>"
			end if
			
			if request.Form("CboMonth") = 3 or month1 = 3 then
				response.Write "<option selected value=3>March</option>"
			else
				response.Write "<option value=3>March</option>"
			end if
			
			if request.Form("CboMonth") = 4 or month1 = 4 then
				response.Write "<option selected value=4>April</option>"
			else
				response.Write "<option value=4>April</option>"
			end if
			
			if request.Form("CboMonth") = 5 or month1 = 5 then
				response.Write "<option selected value=5>May</option>"
			else
				response.Write "<option value=5>May</option>"
			end if
			
			if request.Form("CboMonth") = 6 or month1 = 6 then
				response.Write "<option selected value=6>Jun</option>"
			else
				response.Write "<option value=6>Jun</option>"
			end if
			
			if request.Form("CboMonth") = 7 or month1 = 7 then
				response.Write "<option selected value=7>July</option>"
			else
				response.Write "<option value=7>July</option>"
			end if
			
			if request.Form("CboMonth") = 8 or month1 = 8 then
				response.Write "<option selected value=8>August</option>"
			else
				response.Write "<option value=8>August</option>"
			end if
			
			if request.Form("CboMonth") = 9 or month1 = 9 then
				response.Write "<option selected value=9>September</option>"
			else
				response.Write "<option value=9>September</option>"
			end if
			
			if request.Form("CboMonth") = 10 or month1 = 10 then
				response.Write "<option selected value=10>October</option>"
			else
				response.Write "<option value=10>October</option>"
			end if
			
			if request.Form("CboMonth") = 11 or month1 = 11 then
				response.Write "<option selected value=11>November</option>"
			else
				response.Write "<option value=11>November</option>"
			end if
			
			if request.Form("CboMonth") = 12 or month1 = 12 then
				response.Write "<option selected value=12>December</option></select>"
			else
				response.Write "<option value=12>December</option></select>"
			end if
			                                                 
        %>
        <FONT class=small><b>Year &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
        <select name=cboYear style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 100px"  onchange=Change()>
        <%
			dim year1
			
			year1 = 0
			year1 = request.Form("CboYear")
			
			if year1 = "" then
				Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="select datepart(year,getdate()) as 'currentyear'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
			  	
			  	if not webdbRecordset.EOF then
			  		year1 = webdbRecordset.Fields("currentyear")
			  	end if
			end if
			
			if request.Form("CboYear") = 2005 or Year1 = 2005 then
				response.Write "<option selected value=2005>2005</option>"
			else
				response.Write "<option value=2005>2005</option>"
			end if
			
			if request.Form("cboYear") = 2006 or Year1 = 2006 then
				response.Write "<option selected value=2006>2006</option>"
			else
				response.Write "<option value=2006>2006</option>"
			end if
			
			if request.Form("cboYear") = 2007 or Year1 = 2007 then
				response.Write "<option selected value=2007>2007</option></select>"
			else
				response.Write "<option value=2007>2007</option></select>"
			end if
			
       %>
        </td>
        </tr>
        <tr>
        <td width=10%>&nbsp;</td>
        </tr>
        </form> 
  <TR>
    <TD vAlign=top align=middle colspan=2 width="936" height="193">
      <div align="center">
        <center>
        <form name=frmEmpPratice action=TD_EmpPratice.asp method=post>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="1">
        <TBODY>
    <tr>
      <td bgcolor="#ffffff" width="40">&nbsp;</td>
      <td bgcolor="#f3f3f3" width="100"><font class="marineblack">Date</font></td>
      <td bgcolor="#f3f3f3" width="400"><font class="marineblack">Training Course</font></td>
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
		            
					ssql = "Exec sp_WTD_SelCourseScheduleDetail '" + Session("Regisno") + "'," + cstr(Month1) + "," + cstr(Year1)
					
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
						 
			        page_count = rs.pagecount
			        
			        if 1 > current_page then current_page = 1
			        if current_page > page_count then current_page = page_count
					
					if rs.RecordCount >0 then
			        rs.AbsolutePage = current_page
			        		       
					colour = 0
					rowno = 0
			        do while rs.AbsolutePage = current_page and not rs.EOF
						rowno = rowno + 1
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			        
				        response.write "<tr>"
				        response.write "<td width=40>&nbsp;</td>"
				        response.write "<td width=100 height='20'  " + colour + "><input type=button name=check onclick=validate(" + cstr(rowno) + ")><font class='small'>" + rs("date") + "<input type=hidden name=C" + cstr(rowno) + " value=" + rs("courseID") + "><input type=hidden name=P" + cstr(rowno) + " value='" + rs("Period") + "'></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("Des") + "<input type=hidden name=D" + cstr(rowno) + " value='" + rs("description") + "'></td>"
				        Response.Write "</tr>"
				       
				        rs.Movenext  
				        count = abs(count - 1)        
			        loop
			        
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
			      end if  
			       
			%>
			 <input name=txtaction type=hidden value=>       
			 <table cellSpacing="0" cellPadding="1" border="0" width="90%" bordercolor="#808080">
			
		    <p align=center>
			
			<%Response.Write "<br>" 
			Response.Write "<td colspan=""4"" align=""center"">"
  ''''''''''''''''''''''''''''''''''''''''''''''paging records start'''''''''''''''''''''''''''''''''''''''''''''''''
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "First</font><font=""2""> | </font>"
			end if
  
			iF current_page >= 2 then
				Response.Write "<a href=""TD_TrainCalender?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""TD_TrainCalender?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""TD_TrainCalender?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""TD_TrainCalender?page="
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
            </TABLE></form></center>
      </div>
    </TD></TR>
    <center>  
    <TD align=center colspan=2 width="936" height="40" class="small"><br>
      <!--&nbsp;<br>
      &nbsp;<BR>--><font class ="small" >Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>.</font></TD></TR></TBODY></TABLE></center>
      </div>
</BODY>




