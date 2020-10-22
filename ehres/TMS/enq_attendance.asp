<!-- #INCLUDE FILE = "../global/ConnectStr.asp"-->
<!-- #INCLUDE FILE = "../global/AdoVbs.asp"-->
<!-- #INCLUDE FILE = "../global/inputSession.asp"-->
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

<% dim connect_string
       connect_string ="Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
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
      <P><IMG height=84 src="../Image/engtmsatt.gif" width=683 border=0 ><BR><BR>
        <FORM name=frmAttendance action=enq_attendance.asp method=post>
      <TABLE cellSpacing=0 width="100%" border=0>
        <TR>
           <TD WIDTH=4%></TD>
           <TD>
           <FONT class=small>Employee ID&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT>
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
	  			
	  			tmpEmpID = ""
	  			txtdate1 = ""
	  			txtdate2 = ""
	  			tmpEmpID = session("sscboempid")
			 	
			 	if Request.form("cboempid") = "" or tmpEmpID = "1" then 
	              response.write "<option selected value='1'>All Subordinate</option>"
	    
	            else  
	              response.write "<option value='1'>All Subordinate</option>"
	            end if 
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("empid")) = Request.form("cboempid") ) or ( trim(webdbRecordset.Fields("empid")) = tmpEmpID ) or ( trim(webdbRecordset.Fields("empid")) = session("sscboempidlv"))then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
 				    end if
 				    
 				    if tmpEmpID = "" OR tmpEmpID <> "" then
					      tmpEmpID = trim(webdbRecordset.Fields("empid"))
					      
					end if   
			  	 
				   webdbRecordset.MoveNext  
		        loop       
				Response.Write(tmpEmpID)
			%></select>&nbsp;&nbsp;&nbsp; 
           </TD>
        </TR>
        <tr><td WIDTH=4% height="20"></td></tr> 
        <TR>
          <TD WIDTH=4% ></TD>
          <TD>
                  <FONT class=small>Date (ddmmyyyy)</FONT> <INPUT 
                  style="FONT-SIZE: 8pt" size=16 name=txtDate1
                  <% if request.form("txtdate1")="" then
						response.write " value='" & session("ssDate1lv") & "'"
					 else
					    response.write " value='" & request.form("txtDate1") & "'"
	                 end if               
                  %> ><FONT class=small>&nbsp;&nbsp;to&nbsp;</FONT>
                  <INPUT style="FONT-SIZE: 8pt" size=15 name=txtDate2
                  <% if Request.Form("txtdate2")="" then
                  	    response.write " value='" & session("ssDate2lv") & "'"
                  	 else
                  	    response.write " value='" & request.form("txtDate2") & "'"
    
                  	 end if                  
                  %>><B>&nbsp;&nbsp;<INPUT onmouseover="this.style.cursor='hand';" style="FONT-SIZE: 8pt" onclick=Verify() type=button value=Search name=cmdSearch></A></B><input size=16 name =txtsearch type="hidden"></FORM></TD></TR></TABLE>&nbsp;</P>
      </TD></TR>
  <TR>
    <TD vAlign=top align=middle colspan=2 width="936" height="193">
      <div align="center">
        <center>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="1">
        <TBODY>
    <tr>
      <td bgcolor="#ffffff" width="30">&nbsp;</td>
      <td bgcolor="#f3f3f3" width="50"><font class="marineblack">EmpID</font></td>
      <td bgcolor="#f3f3f3" width="120"><font class="marineblack">Name</font></td>
      <td bgcolor="#f3f3f3" width="120"><font class="marineblack">Date In</font></td>
      <td bgcolor="#f3f3f3" width="120"><font class="marineblack">Date Out</font></td>
      <td bgcolor="#f3f3f3" width="40"><font class="marineblack"> Before OT</font></td>
      <td bgcolor="#f3f3f3" width="40"><font class="marineblack">  After OT</font></td>
      <td bgcolor="#f3f3f3" width="40"><font class="marineblack"> Late</font></td>
      <td bgcolor="#f3f3f3" width="80"><font class="marineblack">Leave Early</font></td>
      <td bgcolor="#f3f3f3" width="40"><font class="marineblack"> Shift</font></td>
      <td bgcolor="#f3f3f3" width="50"><font class="marineblack">Day Type</font></td>
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
		           
		            if Request.Form("txtsearch") = "Search" then
		                vtempdate1 = Request.Form("txtDate1")
						vtempdate2 = Request.Form("txtDate2")
						vtempcboempid = Request.Form("cboempid")     
						call inputleavebal(vtempdate1,vtempdate2,vtempcboempid)
						
				    End If
				    
					tempcboempid =  session("sscboempidlv")
					date1 = session("ssDate1lv")
					date2 = session("ssDate2lv")
					
                    if Request.Form("cboempid") ="1" or tempcboempid ="1" then                
						 ssql = "Exec sp_Wtms_selAttendance '" + Session("Regisno") + "','','" + Session("EmpID") + "','','" + session("ssdate1lv") + "','" + session("ssdate2lv") + "'"		
						 'Response.Write ssql
					else	
					     ssql = "Exec sp_Wtms_selAttendance '" + Session("Regisno") + "','EMPID','" + session("sscboempidlv") + "','','" + session("ssdate1lv") + "','" + session("ssdate2lv") + "'"
					     'Response.Write ssql
					end if
					
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
						 
			        page_count = rs.pagecount
			        
			        if 1 > current_page then current_page = 1
			        if current_page > page_count then current_page = page_count
					
					if rs.RecordCount >0 then
			        rs.AbsolutePage = current_page
			        		       
					colour = 0
					 
			        do while rs.AbsolutePage = current_page and not rs.EOF
			        
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			        
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("empid") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("empname") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("datein") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("dateout") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("bot") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("aot") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("late") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("leaveearly") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("shiftid") + "</td>"				        
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("daytypeid") + "</td>"
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
			 <table cellSpacing="0" cellPadding="1" border="0" width="90%" bordercolor="#808080">
			
		    <p align=center>
			
			<%Response.Write "<br>" 
			Response.Write "<td colspan=""4"" align=""center"">"
  ''''''''''''''''''''''''''''''''''''''''''''''paging records start'''''''''''''''''''''''''''''''''''''''''''''''''
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "First</font><font=""2""> | </font>"
			end if
  
			iF current_page >= 2 then
				Response.Write "<a href=""enq_attendance.asp?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""enq_attendance.asp?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""enq_attendance.asp?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""enq_attendance.asp?page="
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
      &nbsp;<BR>--><font class ="small" >Copyright © 1997-2005 SoftfAC
      Technology Sdn Bhd <i>All Rights Reserved</i>.</font></TD></TR></TBODY></TABLE></center>
      </div>
</BODY>




