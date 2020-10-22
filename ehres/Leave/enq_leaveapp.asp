<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/FormatDate.asp"-->

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

<title>Enquiry - Leave Application</title>

<!--Place this script anywhere in a page.-->
<!--NOTE: You do not need to modify this script.-->

	<script LANGUAGE="JavaScript">

	function Verify()
	{
		msg = "";
		m = true;
		n = true;

		m = CheckDate('txtDate');
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
	
	</script>

</head>

<body bgColor="#ffffff" topmargin="0" leftmargin="0">

<div align="center">
  <center>
<table cellSpacing="0" cellPadding="0" border="0" width="100%">
  <tbody>
  <tr>
    <td vAlign="top" align="center" colspan="2" width="936" bgcolor="#0099CC" height="27">
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
    </td></tr>
  <tr>
    <td vAlign="top" colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></td></tr>
  <tr>
    <td vAlign="top" align="center"><img alt="Main Menu" src="../Image/englsapp.gif" border="0" width="708" height="90"><br>
      &nbsp;</td></tr>


<table border="0" width="96%" cellspacing="0">
  <tr>
    <td width="100%" height="42">
      <table border="0" width="100%" cellspacing="1" height="16">
        <tr>
          <td width="100%" height="10">
            <form method="POST" action="enq_leaveapp.asp" name="frmLeaveApp">
              <p><select size="1" onchange="if(options[selectedIndex].value) top.location.href=('enq_leaveapp.asp?cboSelect=' + options[selectedIndex].value) " style="font-size: 8pt" name="cboSelect">
                <%  
					dim vSelect
			          vSelect = request("cboSelect")
					    
					   if vSelect  = "A" then
			            ' response.write "<Option Selected value = 'A'> All Employees </Option>"
			             response.write "<Option value = 'I'> Individual </Option>"
			          else
			            ' response.write "<Option value = 'A'> All Employees </Option>"
			             response.write "<Option Selected value = 'I'> Individual </Option>"
			          end if
			          					    
				  %>
              </select>
              
			<%If vSelect = "I" Or vSelect = "" Then%>
				 &nbsp;&nbsp;&nbsp;<font class="small">Status</font>
              <select size="1" name="cboStatus" style="font-size: 8pt">
				 
              <%  
					dim vStatus
					
					   vStatus = request.form("cbostatus")
					
					   if vStatus  = "R" then
			             response.write "<Option value = 'A'> Approved </Option>"
			             response.write "<Option value = 'P'> Pending </Option>"
			             response.write "<Option Selected value = 'R'> Rejected </Option>"
			          elseif vStatus  = "P" then
			             response.write "<Option value = 'A'> Approved </Option>"
			             response.write "<Option Selected value = 'P'> Pending </Option>"
			             response.write "<Option value = 'R'> Rejected </Option>"
			          else
			             response.write "<Option Selected value = 'A'> Approved </Option>"
			             response.write "<Option value = 'P'> Pending </Option>"
			             response.write "<Option value = 'R'> Rejected </Option>"
			          end if
				%>
              </select>&nbsp;&nbsp;&nbsp;<font class="small">Year</font>&nbsp;<select size="1" name="cboYear" style="font-size: 8pt">              

				<%    
				      Set webdb = Server.CreateObject("ADODB.Connection")
				          webdb.Open Session("ConnectStr")
				      Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
				      Set webdbCommand = Server.CreateObject("ADODB.Command")

				      ssql = "Exec sp_Wls_selLeaveTransaction '', '', 0, 'ENG', 'Y'"
				      
				             
				      Set webdbCommand.ActiveConnection = webdb
				          webdbCommand.CommandText = ssql
				          webdbRecordset.Open webdbCommand,,1 , 3
          
				 	   i = 1
				 	   Do Until webdbRecordset.EOF
				 	      If i = 2 and Request("cboYear") = "" Then
					         response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("year")) + "'>" + cstr(webdbRecordset.Fields("year")) + "</OPTION> "
					      Elseif Request("cboYear") = cstr(webdbRecordset.Fields("year")) then
				 	         response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("year")) + "'>" + cstr(webdbRecordset.Fields("year")) + "</OPTION> "
				 	      else   
					         response.write "<OPTION value='" + cstr(webdbRecordset.Fields("year")) + "'>" + cstr(webdbRecordset.Fields("year")) + "</OPTION> "
					      End If
					      i = i + 1
					      webdbRecordset.MoveNext
				      loop
				      webdbRecordset.close
				      webdb.close
				%>
              
              </select> 
				<input type="submit" value="Refresh" name="cmdRefresh" style="font-size: 8pt">              
			<%End If%>

		   <%

		   If vSelect = "A" Then%>              
              &nbsp;<font class="small">&nbsp;&nbsp;&nbsp; Date (ddmmyyyy)</font> 
              <input type="text" name="txtDate" size="16" style="font-size: 8pt" <% 
             if request.form("txtDate") <> "" then
                response.write " value='" & request.form("txtDate") & "'"              
             else   
                response.write " value='" & formatdate(now(),"ddmmyyyy") & "'"              
             end if   
          %>>
              <b><input type="button" value="Search" name="cmdSearch" onClick="Verify()" onmouseover="this.style.cursor='hand';" style="font-size: 8pt"></b></p>
          <%End If%>    
              
            </form>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%" height="100%" colspan="1">
    <table border="0" width="100%">
      <tr>
        <td width="86%">
        
	<%If vSelect = "I" Then%>
          <table cellSpacing="0" cellPadding="0" border="0" width="91%" bordercolor="#808080" height="1">
            <tr>
              <td height="20" width="12%" bgcolor="#F3F3F3"><font class="marineblack">Date Apply For</font></td>
              <td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Leave Type</font></td>
              <!--td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Department</font></td-->
              <td height="20" width="40%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reason</b></font></td>
              <td height="20" width="18%" bgcolor="#F3F3F3"><font class="marineblack">Period</font></td>
            </tr>
            <%   
            		 Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		           Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           Set webdbCommand = Server.CreateObject("ADODB.Command")

	 		        If Request("cboYear") = "" or Request("cboStatus") = "" Then
			           ssql = "Exec sp_Wls_selLeaveTransaction '" + Session("Regisno") + "','" + Session("EmpID") + "', 0, 'ENG', 'A'"
			        else
			           ssql = "Exec sp_Wls_selLeaveTransaction '" + Session("Regisno") + "','" + Session("EmpID") + "'," + request("cboyear") + ", 'ENG', '" + request.form("cboStatus") + "'"
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
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("dateapplyfor") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavetype") + "</td>"
'				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("dept") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("reason") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("period") + "</td></tr>"
				        webdbRecordset.MoveNext  
				        count = abs(count - 1)        
			        loop
			        webdbRecordset.close
			        webdb.close      
			 %>

          </table>
	<%End If%>

	<%If vSelect = "A" Then%>
          <table cellSpacing="0" cellPadding="0" border="0" width="91%" bordercolor="#808080" height="1">
            <tr>
			    <td height="20" width="8%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee ID</b></font></td>
			    <td height="20" width="25%" bgcolor="#F3F3F3"><font class="marineblack"><b>Name</b></font></td>
			    <td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Department</b></font></td>			    			    
			    <td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Leave Type</b></font></td>
			    <td height="20" width="5%" bgcolor="#F3F3F3"><font class="marineblack"><b>Day</b></font></td>
			    <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Status</b></font></td>            
            </tr>
            <%   
            		 Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		           Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           Set webdbCommand = Server.CreateObject("ADODB.Command")

				    If Request("txtDate") = "" Then
					    ssql = "Exec sp_Wls_selLeave """ + Session("Regisno") + """, '', '', 'DATEAPPLYFOR'"
					 Else   
					    ssql = "Exec sp_Wls_selLeave """ + Session("Regisno") + """, """ + request("txtDate") + """, '', 'DATEAPPLYFOR'"
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
			        
				        response.write "<tr>"
				        
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("empid") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("empname") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavetype") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("dept") + "</td>"												
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("days") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("status") + "</td></tr>"
				        webdbRecordset.MoveNext  
				        count = abs(count - 1)        
			        loop
			        webdbRecordset.close
			        webdb.close      
			 %>
          </table>	
	<%End If%>
	
        </td>
        <td width="7%"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<table border="0" width="96%">
  <tr>
    <td width="100%" align="center"><img border="0" src="/eHres/Image/dottedlinenav.gif" WIDTH="408" HEIGHT="4"></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<br>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>
<p>&nbsp;</p>
</html>