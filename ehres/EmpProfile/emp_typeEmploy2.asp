<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/ADOVBS.ASP" -->
<% dim connect_string 

connect_string = Session("ConnectStr")%>
 
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
      src="../Image/emptypeemploy.gif" 
     border=0 width="712" height="88"><br>
     <%
       set myconn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.Recordset")    
		
		myconn.open connect_string
		       
        'sql = "Exec sp_is_empprofile '" + trim(Session("EmpID")) + "','" + trim(Session("Regisno")) + "','Retrieve'"
        'sql = "Exec sp_sa_selEmpAddress 'ms0036','id'"
        'Response.Write sql       
        sql = "exec sp_sa_seltypeofEmploy '"+ trim(Session("EmpID")) + "','employ'"
        'sql = "exec sp_sa_seltypeofEmploy 'ms0022','employ'"
                  
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
        'Response.Write sql
        if rs.RecordCount >0 then
           'tempempid =rs("empid")
           'tempempname = rs("empname")
           tempdivision = rs("division")
           tempdepartment = rs("department")
           tempsection = rs("section")
           templine = rs("line")
           tempgroup = rs("group")
           tempprocess = rs("process")
           temptmsid = rs("tms")
           tempfinid = rs("finance")
           tempjobgrade = rs("jobgrade")
           tempjobtitle = rs("jobtitle")
           tempstaterecruited = rs("StateRecruited")
           tempappraissal = rs("appraisal")    'Appraisal
           tempdatejoin = rs("datejoin")
           tempdateconfirm = rs("dateconfirm")
           tempstatus = rs("jobstatus")
           temptype = rs("emptype")
	end if
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
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Division</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempdivision%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Group</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempgroup%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Department</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempdepartment%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Process</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempprocess%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Section</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempsection%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Finance Code</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempfinid%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;TMS ID</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temptmsid%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;State Recruited</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempstaterecruited%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Job Grade</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempjobgrade%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Status</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempstatus%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Job Title</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempjobtitle%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Type</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temptype%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Appraisal Grade</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempappraissal%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Date Joined<i>  (dd/mm/yyyy)</i></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempdatejoin%></font></td>
        </tr>
        <tr>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Line</font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write templine%></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Date Confirmed<i> (dd/mm/yyyy)</i></font></td>
          <td width="25%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempdateconfirm%></font></td>
        </tr>
      </table>
      
     </div>
      
      <p>&nbsp;</p>
      <table border="0" width="96%">
  <tr>
    <td width="100%" align="middle"><IMG border=0 height=4 src="/ehres/Image/dottedlinenav.gif" width=408></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small">
      &nbsp;<br>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All 
      Rights Reserved</i>. </td></tr>
</table>
      </div>
      <div align="center">      
  </table>     
</BODY>