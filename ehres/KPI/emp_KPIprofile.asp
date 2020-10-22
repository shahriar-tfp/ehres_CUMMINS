<!-- #include virtual ="/ehres/global/ConnectStrKPI.asp"-->
<!-- #include virtual ="/ehres/global/ADOVBS.ASP" -->
<% dim connect_string 

connect_string = Session("ConnectStr")%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

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
		       
        sql = "Exec sp_kpi_selempinfo '" + trim(Session("EmpID")) + "'"
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
      	'Response.Write sql + "//"
      	'response.Write connect_string 
      	'response.End 
        'if rs.RecordCount > 0 then
           'tempempid =rs("empid")
           'tempempname = rs("empname")
           'tempempname2 = rs("empname2")
           'tempfirstname = rs("firstname")
           'tempmiddlename = rs("middlename")
           'templastname = rs("lastname")
           'tempinitial = rs("initial")
           'tempalias = rs("alias")
           'tempoldic = rs("oldnric")
           'tempnewic = rs("newnric")
           'tempdob = rs("dob")
           'temppob = rs("pob")
           'tempiccolor = rs("nriccolor")
           'tempreligion = rs("religion")
           'temprace = rs("race")
           'tempmaritalstatus = rs("maritalstatus")
           'tempsex = rs("sex")
           'tempcitizenship = rs("citizenship")
           'temppassno = rs("passportno")
           'temppassexp = rs("passexp")
           'tempworkpermitno = rs("workpermitno")
           'tempworkpertmitexp = rs("workpermitexp")
           'tempmail = rs("email")
           'temphandphone = rs("handphone")
           empname = rs("empname")
           position = rs("Position")
           datejoin = rs("Datejoin")
           dateconfirm = rs("DateConfirm")
           SupName = rs("SupName")
           SupPosition = rs("SupPosition")
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
          <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Name :&nbsp;</font></td>
          <td width="44%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write empname %></font></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Position :&nbsp;</font></td>
            <td width="44%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write Position%></font></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee"align="left"><font class="small">&nbsp;Date Of Commencement(DD/MM/YYYY) :&nbsp;</font></td>
            <td width="44%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write Datejoin%></font></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Date Of Confirmation(DD/MM/YYYY) :&nbsp;</font></td>
            <td width="44%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write DateConfirm%></td>
          </tr>
          <tr>
			<td width="88%" bgcolor="#eeeeee" align="left" colspan=2><font class="small">&nbsp;<b><u>Immediate Superior</u></b></font></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Name :&nbsp;</font></td>
            <td width="44%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write SupName%></td>
          </tr>
          <tr>
            <td width="22%" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Position :&nbsp;</font></td>
            <td width="44%" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<%response.write SupPosition%></td>
          </tr>
        </table>
       <p>&nbsp;</p>
      <!--<p></p>
      <p></p>-->
      <table border="0" width="96%">
  <tr>
    <td width="100%" align="middle"><IMG border=0 height=4 src="/eHres/Image/dottedlinenav.gif" width=408></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small">
      &nbsp;<br>Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
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
      &nbsp;<BR>Copyright © 1997-2006 Software
      Factory Sdn Bhd <i>All Rights Reserved</i>. </TD>--></TR></TBODY></TABLE>
</div>
</BODY>




