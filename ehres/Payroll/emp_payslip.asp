<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/ADOVBS.ASP" -->
<% dim connect_string 

connect_string = "Provider=SQLOLEDB.1;Persist Security Info=False;UID=WEBHR;PWD=password;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"

%>
 
 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 5.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

<body bgColor="#ffffff" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<div align="center">
  <center>
<table cellSpacing="0" cellPadding="0" border="0" width="100%">
  <tbody>
  <tr>
    <td vAlign="top" align="center" colspan="2" width="936" bgcolor="#0099CC" height="29">
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
  <!--<tr>
    <td vAlign="top" colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></td></tr>
  <tr>-->
  <TR>
    <TD vAlign=top colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></TD></TR>
  <TR>
    <TD vAlign=top align=center width="1" height="109"></TD>
  </center>
    <TD vAlign=top align=center width="100%" height="109">
      <p align="center"><IMG alt='Main Menu' 
      src="../Image/emppayslip.gif" 
     border=0 width="712" height="88"><!--<br>-->
     
     <form method="POST" action="emp_payslip.asp" name="frmpayslip">
    
     <%
       set myconn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.Recordset")    
		
		myconn.open connect_string
		               'sql ="exec sp_pr_SelPayrollDetails1 '16202-H','a003',1,1,'NAME'"        sql ="exec sp_pr_SelPayrollDetails1 '" + trim(Session("regisno")) + "','" + trim(Session("EmpID")) + "',1,1,'NAME'"              
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
        'response.write sql
        tempempid = rs("empid")
        tempempname = rs("empname") 
        'Response.Write (tempempid)
        'Response.Write (tempempname)           
     %>
    
     <table border ="0" width ="90%" cellSpacing="1" cellPadding="2" height="34">
     <!--<tr>
		<td height="28" width="20%"><font class=small>Employee ID : <%Response.Write(tempempid) %></td>
		<td height="28"><font class=small>Employee Name : <%Response.Write(tempempname)%></font></td>
     </tr>-->
     <tr>
		<td height="28" width="20%"><font class=small>Month : <select size="1" name=cbomonth style="font-size: 8pt">
       		<%  
				'dim vStatus1
					    if request("cbomonth") ="" then
					       vStatus = "7"
					    else
					       vStatus = request("cbomonth")
					    end if       
					    
					   if vStatus = "1" then
			             response.write "<Option Selected value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			          elseif vStatus  = "2" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option Selected value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			          elseif vStatus  = "3" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option Selected value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			           elseif vStatus  = "4" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option Selected  value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			          elseif vStatus  = "5" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option Selected value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			          elseif vStatus  = "6" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option Selected value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			          elseif vStatus  = "7" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option Selected value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			          elseif vStatus  = "8" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option Selected value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			          elseif vStatus  = "9" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option Selected value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"   
			          elseif vStatus  = "10" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option Selected value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			          elseif vStatus  = "11" then
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option Selected value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option value = '12'>December</Option>"
			          else vStatus  = "12" 
			             response.write "<Option value = '1'>January</Option>"
			             response.write "<Option value = '2'>February</Option>"
			             response.write "<Option value = '3'>March</Option>"
			             response.write "<Option value = '4'>April</Option>"
			             response.write "<Option value = '5'>May</Option>"
			             response.write "<Option value = '6'>June</Option>"
			             response.write "<Option value = '7'>July</Option>"
			             response.write "<Option value = '8'>August</Option>"
			             response.write "<Option value = '9'>September</Option>"
			             response.write "<Option value = '10'>Oktober</Option>"
			             response.write "<Option value = '11'>November</Option>"
			             'response.write "<Option value = 'A'>  </Option>"
			             response.write "<Option Selected value = '12'>December</Option>"
			          end if
			          
				%></select><!--<%Response.Write(tempempid) %>--></td>
			
		<td height="28"><font class=small> Year : <select size="1" name=cboYear style="font-size: 8pt">
       		<%    
				      Set webdb = Server.CreateObject("ADODB.Connection")
				          webdb.Open Session("ConnectStr")
				      Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
				      Set webdbCommand = Server.CreateObject("ADODB.Command")
					      
				      ssql = "Exec sp_Wls_selLeaveTransaction3 '', '', 0, 'ENG', 'Y'"
       
				      Set webdbCommand.ActiveConnection = webdb
				          webdbCommand.CommandText = ssql
				          webdbRecordset.Open webdbCommand,,1 , 3
					'response.write ssql
					  tempdate1 = year(now())
				 	   i = 1
				 	   Do Until webdbRecordset.EOF
				 	      If i = 2 and Request("cboYear") = "" Then  'cstr(webdbRecordset.Fields("year"))   'cstr(webdbRecordset.Fields("year"))
					         response.write "<OPTION Selected value='" + cstr(tempdate1) + "'>" + cstr(tempdate1) + "</OPTION> "
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
				</select>&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="Search" name="cmdSearch"  class="small"><!--<%Response.Write(tempempname)%>--></td>
     </tr>
     </table>
       
      <table cellSpacing="0" cellPadding="0" border="0" width="90%" bordercolor="#808080">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
				<tr>
					<td height="20" width="4%" bgcolor="#F3F3F3"></td> 
					<td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Process ID</font></td>
					<td height="20" width="20%" bgcolor="#F3F3F3" align="right"><font class="marineblack">Amount Earned</font></td>
					<td height="20" width="15%" bgcolor="#F3F3F3" align="right"><font class="marineblack">Amount Deducted</font></td>
					<td height="20" width="15%" bgcolor="#F3F3F3" align="right"><font class="marineblack">Net Salary</font></td>
					<!--<td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Total</font></td>-->			
				</tr>
				<!--<tr>
				<td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Total</font></td>
				</tr>-->
				
				
	<% dim tempdebit
	   dim tempcredit
	   dim total
	   dim totcredit 
	    
	    
	    set myconn = server.CreateObject("ADODB.Connection")
        set rs1 = server.CreateObject("ADODB.Recordset")    
		
		myconn.open connect_string
	    
	    tempdate = year(now())
	    'Response.Write (tempdate)        if Request.form("cboYear")="" then
           sql = "exec sp_pr_SelPayrollDetails1 '" + trim(Session("regisno")) + "','" + trim(Session("EmpID")) + "','" + cstr(tempdate) + "','" + vStatus + "','ALLOWANCE'"
        else             sql = "exec sp_pr_SelPayrollDetails1 '" + trim(Session("regisno")) + "','" + trim(Session("EmpID")) + "','" + request("cboYear") + "','" + vStatus + "','ALLOWANCE'"
        end if
        'sql = "exec sp_pr_SelPayrollDetails1 '" + trim(Session("regisno")"','" + trim(Session("EmpID")) + "','" + request("cboyear") + "'," + request("cbomonth") + ",'ALLOWANCE'"       
        'Response.Write sql       
        rs1.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
        'response.write sql
        tempempid = rs1(2)
        tempempname = rs1(3) 
       	
	           colour = 0
			   tempdebit = 0
			   tempcredit = 0		
			   total = 0
			        
			        do until rs1.EOF
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
				       
                        tempdebit =  tempdebit + cdbl (rs1("empdebit"))
                        'debit = cdbl(rs1("empdebit"))
                        'if debit ="0" then
                           'debit = " " 
                        'end if   
                                     
                        tempcredit =  tempcredit + cdbl (rs1("empcredit"))
                        
                        total = tempdebit - tempcredit
                         
				        response.write "<tr>"
                        response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs1(0) + "</td>"
				        response.write "<td height='20' align='right'" + colour + "><font class='small'>" + cstr(rs1("empdebit")) + "</td>"
				        response.write "<td height='20' align='right'" + colour + "><font class='small'>" + cstr(rs1("empcredit")) + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + tempid + "</td></tr>"
				        'response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("reason") + "</td></tr>" 'cstr(rs1("empdebit"))
				        rs1.MoveNext  
				        count = abs(count - 1)        
			        loop
		            'end if	
			        rs1.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
			        colour = ""
			        Response.write "<tr>"
			        response.write "<td height='20' width='4%'" + colour + "></td>"
			        response.write "<td height='20' width='4%'" + colour + "><font class='small'>Total</td>"
			        response.write "<td height='20' align='right'" + colour + "><font class='small'>" + cstr(tempdebit) + "</td>"
			        response.write "<td height='20' align='right'" + colour + "><font class='small'>" + cstr(tempcredit) + "</td>"
			        response.write "<td height='20' align='right'" + colour + "><font class='small'>" + cstr(total) + "</td></tr>"
		           
	%>		        			
          
     </div>
      
      <p></p>
      <table border="0" width="96%">
  <tr>
    <td width="100%" align="middle"><IMG border=0 height=4 src="/ehres/Image/dottedlinenav.gif" width=408></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small">
      &nbsp;<br>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>
      </div>
      <div align="center">      
  </table>     
</BODY>