<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
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
      <P><IMG height=87 src="../Image/engtmserr.gif" width=684 border=0 ><BR><BR>
      <TABLE cellSpacing=0 width="100%" border=0>
        
        <TR>
          <TD>
                  <FORM name=frmAttendance action=enq_atterror.asp method=post>
<font class="small">Type Of Error </font><select size="1" name="cboErrCode" class="small">
                 
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
                   <option selected value= "ALL">All Error</option>
                <%end if %>   
              
              </select>&nbsp;&nbsp;&nbsp;                   
                  <FONT class=small>Date (ddmmyyyy)</FONT> <INPUT 
                  style="FONT-SIZE: 8pt" size=16 name=txtDate1> <FONT 
                  class=small>to</FONT> <INPUT style="FONT-SIZE: 8pt" size=15 
                  name=txtDate2><B><INPUT onmouseover="this.style.cursor='hand';" style="FONT-SIZE: 8pt" onclick=Verify() type=button value=Search name=cmdSearch></B></FORM></TD></TR></TABLE>&nbsp;</P>
      </TD></TR>
  <TR>
    <TD vAlign=top align=middle colspan=2 width="936" height="193">
      <div align="center">
        <center>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="1">
        <TBODY>
    <tr>
    
    <% If request("cboErrCode") <> "" And request("cboErrCode") <> "ALL" Then %>
          <td bgcolor="#ffffff" width="30">&nbsp;</td>
	      <td width="100" bgcolor="#f3f3f3"><font class="marineblack"><b>Date</b></font></td>
	      <td width="130" bgcolor="#f3f3f3"><font class="marineblack"><b>Date In</b></font></td>
	      <td width="130" bgcolor="#f3f3f3"><font class="marineblack"><b>Date Out</b></font></td>
	      <td width="50" bgcolor="#f3f3f3"><font class="marineblack"><b>Shift In</b></font></td>
	      <td width="50" bgcolor="#f3f3f3"><font class="marineblack"><b>Shift Out</b></font></td>
	      <td width="30" bgcolor="#f3f3f3"><font class="marineblack"><b>Late</b></font></td></tr>
	<% Else%>
	      <td bgcolor="#ffffff" width="30">&nbsp;</td>
	      <td width="150" bgcolor="#f3f3f3"><font class="marineblack"><b>Type Of Error</b></font></td>
	      <td width="100" bgcolor="#f3f3f3"><font class="marineblack"><b>Date</b></font></td>
	      <td width="130" bgcolor="#f3f3f3"><font class="marineblack"><b>Date In</b></font></td>
	      <td width="130" bgcolor="#f3f3f3"><font class="marineblack"><b>Date Out</b></font></td>
	      <td width="50" bgcolor="#f3f3f3"><font class="marineblack"><b>Shift In</b></font></td>
	      <td width="50" bgcolor="#f3f3f3"><font class="marineblack"><b>Shift Out</b></font></td>
	      <td width="30" bgcolor="#f3f3f3"><font class="marineblack"><b>Late</b></font></td></tr>
	<% End If%>
    </tr>    

            <%
           		   Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		           Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           Set webdbCommand = Server.CreateObject("ADODB.Command")
  		           
  		           If Request("txtDate1") <> "" And Request("txtDate2") <> "" And Request("cboErrCode") <> ""  And Request("cboErrCode") <> "ALL" Then
			           ssql = "Exec sp_Wtms_selTMSError """ + Session("Regisno") + """, """ + Session("EmpID") + """, " + request("cboErrCode") + ", """ + request("txtDate1") + """, """ + request("txtDate2") + """, """ + request("cboErrCode") + """"
			           Response.Write SSQL
		           Elseif Request("txtDate1") <> "" And Request("txtDate2") <> "" And Request("cboErrCode") = "ALL" Then
			           ssql = "Exec sp_Wtms_selTMSError """ + Session("Regisno") + """, """ + Session("EmpID") + """, 0 , """ + request("txtDate1") + """, """ + request("txtDate2") + """, """ + request("cboErrCode") + """"
			           Response.Write SSQL
		           Elseif Request("txtDate1") = "" And Request("txtDate2") = "" And (Request("cboErrCode") = "" or Request("cboErrCode") = "ALL") Then
			           ssql = "Exec sp_Wtms_selTMSError """ + Session("Regisno") + """, """ + Session("EmpID") + """, 0, '', '', 'ALL'"
			           Response.Write SSQL
		           Elseif Request("txtDate1") = "" And Request("txtDate2") = "" And Request("cboErrCode") <> "" And Request("cboErrCode") <> "ALL" Then
			           ssql = "Exec sp_Wtms_selTMSError """ + Session("Regisno") + """, """ + Session("EmpID") + """, 0, '', '', ''"			           
			           Response.Write SSQL
		           End If

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

			          If request("cboErrCode") <> "ALL" and request("cboErrCode") <> "" Then
				         response.write "<tr>"
				         response.write "<td></td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("date") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("datein") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("dateout") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("shiftin") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("shiftout") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("late") + "</td></tr>"
				      ElseIf request("cboErrCode") = "ALL" or request("cboErrCode") = "" Then
				         response.write "<tr>"
				         response.write "<td></td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("description") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("date") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("datein") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("dateout") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("shiftin") + "</td>"
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("shiftout") + "</td>"				        
				         response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("late") + "</td></tr>"
				      End If
				   
				      webdbRecordset.MoveNext  
				      count = abs(count - 1)        
			        loop
			        
			        webdbRecordset.close
			        webdb.close      

			 %>
            </TBODY>
            </TABLE></center>
      </div>
    </TD></TR>
  <TR>
    <TD align=middle colspan=2 width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<BR>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </TD></TR></TBODY></TABLE></center>
</div>
</BODY>