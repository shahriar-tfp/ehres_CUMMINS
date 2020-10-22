<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->

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

</head>

<body topmargin="0" leftmargin="0" bgColor="#ffffff">

<table border="0" width="96%" cellspacing="0" height="103">
  <tr>
    <td width="37%" height="78">
      <table border="0" width="50%" background="../Image/LeaveApp.gif" height="41">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>  
  <tr>  
    <td width="90%" height="78">
      <table border="0" width="100%" cellspacing="1">
        <tr>
          <td width="100%" height="10">
            <form method="POST" action="leaveapp2.asp" name="frmLeaveApp">
              <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="small">Status</font>
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
              
              </select><input name=txtAction type=hidden> 
				<input type="submit" value="Refresh" name="cmdRefresh" style="font-size: 8pt">              
              </p>
              
            </form>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%" height="21" colspan="2">

    <table border="0" width="100%">
      <tr>
        <td width="7%"></td>
        <td width="86%">
        
          <table cellSpacing="0" cellPadding="0" border="0" width="91%" bordercolor="#808080">
            <tr>
              <td height="20" width="4%" bgcolor="#F3F3F3"></td> 
              <td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Date Apply For</font></td>
              <td height="20" width="30%" bgcolor="#F3F3F3"><font class="marineblack">Leave Type</font></td>
              <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Period</font></td>
              <td height="20" width="50%" bgcolor="#F3F3F3"><font class="marineblack">Reason</font></td>
            </tr>
            <%   
            		 Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		           Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           Set webdbCommand = Server.CreateObject("ADODB.Command")

	 		        If Request("cboYear") = "" or Request("cboStatus") = "" Then
			           ssql = "Exec sp_Wls_selLeaveTransaction """ + Session("Regisno") + """, """ + Session("EmpID") + """, 0, 'ENG', 'A'"
			        else
			           ssql = "Exec sp_Wls_selLeaveTransaction """ + Session("Regisno") + """, """ + Session("EmpID") + """, " + request("cboyear") + ", 'ENG', """ + request.form("cboStatus") + """"
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
                       	response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("dateapplyfor") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavetype") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("period") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("reason") + "</td></tr>"
				        webdbRecordset.MoveNext  
				        count = abs(count - 1)        
			        loop
			        webdbRecordset.close
			        webdb.close      
			 %>

          </table>

	<%If vSelect = "" or vSelect = "A" Then%>
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
    <td width="100%" align="center"><img border="0" src="/ehres/Image/dottedlinenav.gif" WIDTH="408" HEIGHT="4"></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<br>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>

<p>&nbsp;</p>
</html>