<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/ADOVBS.ASP" -->

<% dim connect_string 

	connect_string = "Provider=SQLOLEDB.1;Persist Security Info=False;UID=HRISMGR;PWD=TIGER;Initial catalog=HRDB_CSEM;Data Source=DESKTOP-SQCF4E5\DEV2017;Connect Timeout=900000"



%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>eHRES</TITLE>
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
      src="../Image/empprofile1.gif" 
     border=0 width="712" height="88"><br>
     <% 
        set myconn = server.CreateObject("ADODB.Connection")      
 	set rs = server.CreateObject("ADODB.Recordset")    			
	myconn.open connect_string		               
	sql = "Exec sp_is_empprofile '" + trim(Session("EmpID")) + "','" + trim(Session("Regisno")) + "','Retrieve'"       
        'sql = "Exec sp_is_empprofile 'ms0036','185612-k','Retrieve'"       
        'Response.Write sql              
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
      	' Response.Write sql
        'if rs.RecordCount > 0 then
           tempempid =rs("empid")
           tempempname = rs("empname")
           tempempname2 = rs("empname2")
           tempfirstname = rs("firstname")
           tempmiddlename = rs("middlename")
           templastname = rs("lastname")
           tempinitial = rs("initial")
           tempalias = rs("alias")
           tempoldic = rs("oldnric")
           tempnewic = rs("newnric")
           tempdob = rs("dob")
           temppob = rs("pob")
           tempiccolor = rs("nriccolor")
           tempreligion = rs("religion")
           temprace = rs("race")
           tempmaritalstatus = rs("maritalstatus")
           tempsex = rs("sex")
           tempcitizenship = rs("citizenship")
           temppassno = rs("passportno")
           temppassexp = rs("passexp")
           tempworkpermitno = rs("workpermitno")
           tempworkpertmitexp = rs("workpermitexp")
           tempmail = rs("email")
           temphandphone = rs("handphone")
      
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
    <TD vAlign=top align=center colspan=2 width="714" height="193">
      <div align="center">
      
      <p align="left">&nbsp;
      
      </div>
    </center>
      <div align="center">
       
      <table border="0" width="88%" cellSpacing="1" cellPadding="2">
        <tr>
          <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Name 2&nbsp;</font></td>
          <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempempname2 %></font></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Alias&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempalias%></font></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee"align="left"><font class="small">&nbsp;First Name&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempfirstname%></font></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Religion&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempreligion%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Middle Name&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempmiddlename%></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Race&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temprace%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Last Name&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write templastname%></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Citizenship&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempcitizenship%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;New Ic No&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempnewic%></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Initial&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempinitial%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Old Ic No&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempoldic%></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Passport No&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temppassno%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Ic Color&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempiccolor%></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Passport Expired Date<i>(mm/dd/yyyy)</i>&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temppassexp%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Date Of Birth&nbsp;<i>(mm/dd/yyyy)</i></font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempdob%></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Work Permit No&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempworkpermitno%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Place Of Birth&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temppob%></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Work Permit Expired Date <i>(mm/dd/yyyy)</i>&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempworkpertmitexp%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Sex&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempsex%></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Hand Phone No&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write temphandphone%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Marital Status&nbsp;</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempmaritalstatus%></td>
          <!--</tr>
          <tr>-->
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Email Address</font></td>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write tempmail%></td>
          </tr>
          <!--<tr>
            <td width="50%">&nbsp;</td>
            <td width="50%">&nbsp;</td>
          </tr>-->
        </table>
       <p>&nbsp;</p>
      <!--<p></p>
      <p></p>-->
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
      
      <!--<p align="left">&nbsp;-->
      
      </div>
      <div align="center">
      
      <!--<p align="left">&nbsp;-->
      
      </div>
    </center>
    </TD></TR>
  <TR>
    <!--<TD align=middle colspan=2 width="714" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<BR>Copyright © 1997-2000 Software
      Factory Sdn Bhd <i>All Rights Reserved</i>. </TD>--></TR></TBODY></TABLE>
</div>
</BODY>