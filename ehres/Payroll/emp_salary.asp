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
      src="../Image/empsalary.gif" 
     border=0 width="712" height="88"><br>
     <%
       set myconn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.Recordset")    
		
		myconn.open connect_string
		       
        'sql = "Exec sp_is_empprofile '" + trim(Session("EmpID")) + "','" + trim(Session("Regisno")) + "','Retrieve'"
        'sql = "Exec sp_sa_selEmpAddress 'ms0036','id'"
        'Response.Write sql       
        'sql = "exec sp_sa_seltypeofEmploy '"+ trim(Session("EmpID")) + "','employ'"
        sql = "exec sp_sa_selEmpSalaryweb '"+ trim(Session("EmpID")) + "','ID'"
        'sql = "exec sp_sa_seltypeofEmploy 'ms0022','employ'"
                  
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText
         'response.write sql
        'if rs.RecordCount <> 0 then
        do until rs.EOF 
           tempbasicsalary = rs("salary")
           temppaymode = rs("paymode")
           tempbankname = rs("bankname")
           tempbranch = rs("branch")
           tempbankaccno = rs("bankaccno")
           tempepfcode = rs("epfcode")
           tempepfno = rs("epf")
           tempsocsocategory = rs("socsodesc")
           tempsocso = rs("socso")
           temppcbno = rs("pcb")
           tempnoofchild = rs("noofchild")
           tempworkingstatus = rs("workingstatus")    'Appraisal
           tempamanah = rs("asb")
           temptabung = rs("tb")
 
           rs.MoveNext  
		   count = abs(count - 1)        
	       loop
		            'end if	
		   rs.close
		   set rs = nothing
		   myconn.close
		   set myconn = nothing   
        'end if   
           'tempstatus = rs("jobstatus")
           'temptype = rs("emptype")
     %>
    
     <!--<table border ="0" width ="90%" cellSpacing="1" cellPadding="2" height="34">
     <tr>
		<td height="28"><font class="marineblack">Employee ID :</font><font class=small>&nbsp;<%Response.Write tempempid %></font></td>
		<td height="28"><font class="marineblack">Employee Name :</font><font class=small>&nbsp;<%Response.Write tempempname%></font></td>
     </tr>
     </table>--> 
  </TD></TR>
  <center>
  <TR>
    <TD vAlign=top align=center colspan=2 width="716" height="193">
      <div align="center"> 
      &nbsp; 
      </div>
      <div align="center"> 
      <table border="0" width="90%" cellSpacing="1" cellPadding="2">
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Basic Salary</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempbasicsalary%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Socso Category</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempsocsocategory%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Pay Mode</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temppaymode%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Socso No</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempsocso%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Bank Name</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempbankname%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Income Tax No</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temppcbno%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Branch</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempbranch%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;No Of Children</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempnoofchild%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Bank Account No</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempbankaccno%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Wife Working Status</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempworkingstatus%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;EPF Code</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempepfcode%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Amanah Saham No</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempamanah%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;EPF No</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempepfno%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Tabung Haji No</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temptabung%></font></td>
        </tr>
        <!--<tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Line</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write templine%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Date Confirmed<i> (dd/mm/yyyy)</i></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempdateconfirm%></font></td>
        </tr>-->
      </table>
      
     </div>
      
      <p>&nbsp;</p>
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