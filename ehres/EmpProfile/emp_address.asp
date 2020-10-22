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
      src="../Image/empaddress.gif" 
     border=0 width="712" height="88"><br>
    
     <% 
        set myconn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.Recordset")    
		
		myconn.open connect_string
		       
        'sql = "Exec sp_is_empprofile '" + trim(Session("EmpID")) + "','" + trim(Session("Regisno")) + "','Retrieve'"
        'sql = "Exec sp_is_empprofile 'ms0036','185612-k','Retrieve'"
        sql = "Exec sp_sa_selEmpAddress1 '" + trim(Session("EmpID")) + "','id'"
        'Response.Write sql       
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
      	'Response.Write sql
        'if rs.RecordCount > 0 then
         tempempid = rs("empid")
         tempempname = rs("empname")
         praddress = rs("presentadd1")
         prpostcode = rs("presentpostcode")
         prcity = rs("presentcity")
         prstate = rs("presentstate")
         prtelno = rs("presenttelno")
         prcountry = rs("presentcountry")
         pmaddress = rs("permanentadd1")
         pmpostcode = rs("permanentpostcode")
         pmcity = rs("permanentcity")
         pmstate = rs("permanentstate")
         pmtelno = rs("permanenttelno")
         pmcountry = rs("permanentcountry")
	%>
     
     
  </TD></TR>
  <center>
  
  <TR>
    <TD vAlign=top align=center colspan=2 width="90%" height="193">
      <div align="center">
      
      <!--<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;-->
      
      </div>
      <div align="center">
      <table border="0" width="90%" cellspacing="0" cellpadding="0" align=center >
        <tr>
          <td width="100%" bgcolor="#0099CC">    <!--#000080-->
            <b><font face=Verdana size=2 color="#FFFFFF">&nbsp;Present Address</font></b></td>
        </tr>
        <tr>
          <td width="80%">&nbsp;
            <table border="0" width="100%" height="100" cellspacing="2" cellpadding="1">
              <tr>
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Address</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write praddress%></font></td>
                <!--<td width="50%" height="19"></td>
                <td width="50%" height="19"></td>-->
              </tr>
              <tr>
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Post Code</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write prpostcode%></font></td>
              </tr>
              <tr>   
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Country</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write prcountry%></font></td>
              </tr>
              <tr>
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;State</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write prstate%></td>
              </tr>
              <tr>  
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Tel No</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write prtelno%></td>
              </tr>
              <tr>
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;City</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write prcity%></font></td>
                <!--<td width="50%" height="19"></td>
                <td width="50%" height="19"></td>-->
              </tr>
            </table>
            
          </td>
        </tr>
      </table>
      <!--<br>-->
      <P></P>
      
     <div align="center">
      <table border="0" width="90%" cellspacing="0" cellpadding="0" align=center>
        <tr>
          <td width="100%" bgcolor="#0099CC">
            <b><font face=Verdana size=2 color="#FFFFFF">&nbsp;Permanent Address</font></b></td>
        </tr>
        <tr>
          <td width="80%">&nbsp;
          <table border="0" width="100%" height="100" cellspacing="2" cellpadding="1">
              <tr>
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Address</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write pmaddress%><font></td>
                <!--<td width="50%" height="19"></td>
                <td width="50%" height="19"></td>-->
              </tr>
              <tr>
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Post Code</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write pmpostcode%></font></td>
              </tr>
              <tr>  
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Country</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write pmcountry%></font></td>
              </tr>
              <tr>
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;State</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write pmstate%></font></td>
             </tr>
             <tr>   
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;Tel No</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write pmtelno%></font></td>
              </tr>
              <tr>
                <td width="20%" height="19" bgcolor="#eeeeee" align="left"><font class="small">&nbsp;City</font></td>
                <td width="60%" height="19" bgcolor="#eeeeee" align="left"><font class=small>&nbsp;<% Response.Write pmcity%></font></td>
                <!--<td width="50%" height="19"></td>
                <td width="50%" height="19"></td>-->
              </tr>
            </table>
            <p>&nbsp;</p>
          </td>
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
      
      
</center>
      
      
</BODY>