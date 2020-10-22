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

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Enquiry - Leave Balance</title>
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">

<table border="0" width="96%">
  <tr>
    <td width="50%">
      &nbsp;
      <table border="0" width="100%" background="../Image/LeaveBal.gif" height="32">
        <tr>
          <td width="100%" height="28">&nbsp;</td>
        </tr>
      </table>
    </td>
    <td width="50%"></td>
  </tr>
  <tr>
    <td width="50%"></td>
    <td width="50%"></td>
  </tr>
</table>

<table border="0" width="96%">
  <tr>
    <td width="2%">&nbsp;</td>
    <td width="95%">
      <table cellSpacing="0" cellPadding="0" border="0" width="100%" bordercolor="#808080">
            <tr>
			    <td height="20" width="30%" bgcolor="#F3F3F3"><font class="marineblack"><b>Leave Type</b></font></td>
			    <td height="20" width="5%" bgcolor="#F3F3F3"><font class="marineblack"><b>Year</b></font></td>
			    <td height="20" width="18%" bgcolor="#F3F3F3"><font class="marineblack"><b>Total Entitlement</b></font></td>
			    <td height="20" width="16%" bgcolor="#F3F3F3"><font class="marineblack"><b>Entitle To Date</b></font></td>
			    <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Leave B/F</b></font></td>
			    <td height="20" width="7%" bgcolor="#F3F3F3"><font class="marineblack"><b>Day(s)</b></font></td>
			    <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Burn Leave</b></font></td>
			    <td height="20" width="12%" bgcolor="#F3F3F3"><font class="marineblack"><b>Balance</b></font></td>    
            </tr>
        <tr>
			<%    Set webdb = Server.CreateObject("ADODB.Connection")
					   webdb.Open Session("ConnectStr")
			      Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			      Set webdbCommand = Server.CreateObject("ADODB.Command")
			' add by tey sing yeh on Aug 02 2005
			datearray = split(session("currentDate"),"/")
			currentdate = datearray(1) & "/" & datearray(0) & "/" & datearray(2)
			currentdate = MonthName(Month(currentdate)) & " " & Day(currentdate) & " " & Year(currentdate)
			'end by tey sing yeh
			'ssql = "Exec sp_ls_selAllLeaveBal """ + Session("Regisno") + """, """ + Session("empid") + """, """ + Session("CurrentDate") + """, '', '0', 'RETRIEVE'"  'mark by tey sing yeh
  		    ssql = "Exec sp_ls_selAllLeaveBal """ + Session("Regisno") + """, """ + Session("empid") + """, """ + currentdate + """, '', '0', 'RETRIEVE'"  
			  	   Set webdbCommand.ActiveConnection = webdb
			  	       webdbCommand.CommandText = ssql
			  	       webdbRecordset.Open webdbCommand,,1 , 3
			  	       'set webdbRecordset = webdb.Execute(ssql)
			  	       'response.write ssql
			  	       
			  	       
			  	 colour = 0
	



 				  	Do Until webdbRecordset.EOF
				      if count = 1 then
				        colour = " bgcolor='#eeeeee'"
				      else
				         colour = ""
				     end if
				        
				      response.write "<tr>"				        
				      response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("description") + "</td>"
				      response.write "<td height='20' align='left'" + colour + "><font class='small'>" + cstr(webdbRecordset.Fields("entitlementyear")) + "</td>"
				      response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("actualentitlement") + "</td>"
				      response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("tolentitlement") + "</td>"
				      response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("carryforwarddays") + "</td>"
				      response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("daystakenexcp") + "</td>"				        
				      response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("burnleave") + "</td>"
				      response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("leavebalexcp") + "</td></tr>"
				      webdbRecordset.MoveNext  
				      count = abs(count - 1)        
		       loop  
		    
 			    webdbRecordset.close
			    webdb.close   
			set webdbRecordset = nothing
			set webdb = nothing

			%>
        </tr>
      </table>
    </td>
    <td width="3%">&nbsp;</td>
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
  
</body>

</html>