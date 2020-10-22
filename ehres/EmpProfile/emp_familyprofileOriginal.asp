<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/ADOVBS.ASP" -->
<% dim connect_string 

connect_string = Session("ConnectStr")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<script language="JavaScript1.2">

/*
Highlight Table Cells Script- 
Last updated: 99/01/21
© Dynamic Drive (www.dynamicdrive.com)
For full source code, installation instructions,
100's more DHTML scripts, and Terms Of
Use, visit dynamicdrive.com
*/

function changeto(highlightcolor){
source=event.srcElement
if (source.tagName=="tr"||source.tagName=="table")
return
while(source.tagName!="td")
source=source.parentElement
if (source.style.backgroundColor!=highlightcolor&&source.id!="ignore")
source.style.backgroundColor=highlightcolor
}

function changeback(originalcolor){
if (event.fromElement.contains(event.toElement)||source.contains(event.toElement)||source.id=="ignore")
return
if (event.toElement!=source)
source.style.backgroundColor=originalcolor
}
</script>
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
      src="../Image/empfamilyprofile.gif" 
     border=0 width="712" height="88"><br>
     
    
     <!--<table border ="0" width ="90%" cellSpacing="1" cellPadding="2" height="34">
     <tr>
		<td height="28"><font class="marineblack">Employee ID :</font><font class=small>&nbsp;<%response.Write tempempid %></font></td>
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
      <FONT color=#ff0080><STRONG>
      <!--onMouseover="changeto('lightgreen')" 
onMouseout="changeback('white')"--></STRONG></FONT>&gt;
      <table cellSpacing="0" cellPadding="0" border="0" width="92%" onMouseover="changeto('lightgreen')" onMouseout="changeback('white')" bordercolor="#808080">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
				<tr>
					<!--<td height="20" width="4%" bgcolor="#F3F3F3"></td>--> 
					<td height="20" width="20%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Name</font></td>
					<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Date of Birth</font></td>
					<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">New IC No</font></td>
					<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Relation</font></td>
					<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Status</font></td>
					<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Working Status</font></td>
					<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Company Name</font></td>
					<!--<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Occupation</font></td>-->
					<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Income Tax No</font></td>
					<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Income Tax Branch</font></td>
					<td height="20" width="10%" bgcolor="#F3F3F3" align="left"><font class="marineblack">Benifits(%)</font></td>
					<!--<td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Total</font></td>-->			
				</tr>
	 
	 <%
        set myconn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.Recordset")    
		
		myconn.open connect_string
		       
        'sql = "Exec sp_is_empprofile '" + trim(Session("EmpID")) + "','" + trim(Session("Regisno")) + "','Retrieve'"
        'sql = "Exec sp_sa_selEmpAddress 'ms0036','id'"
        'Response.Write sql       
        sql = "Exec sp_is_selFamilyInfoweb '" + trim(Session("EmpID")) + "','eng'"                'sql = "Exec sp_is_empprofile 'ms0036','185612-k','Retrieve'"
        'Response.Write sql       
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
      	
      	do until rs.EOF
			if count = 1 then
				 colour = " bgcolor='#eeeeee'"
		    else
				 colour = ""
		    end if	
		    			           
				        response.write "<tr>"
                        'response.write "<td height='20' width='4%'" + colour + "></td> "
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("name") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("dob") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("newic") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("relation") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("status") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("workingstatus") + "</td>"
				        'response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("occupation") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("company") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("pcbno") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("pcbbranch") + "</td>"
				        'response.write "<td height='20' align='right'" + colour + "><font class='small'>" + rs("newic") + "</td>"
				        response.write "<td height='20' align='left'" + colour + "><font class='small'>" + cstr(rs("benefit")) + "</td></tr>"
				        'response.write "<td height='20' align='left'" + colour + "><font class='small'>" + rs("reason") + "</td></tr>" 'cstr(rs1("empdebit"))
				        rs.MoveNext  
				        count = abs(count - 1)        
			        loop
		            'end if	
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
     %>			
      
      </table>
      
     </div>
      
      <p>&nbsp;</p>
      <table border="0" width="96%" height="48">
  <tr>
    <td width="100%" align="middle" height="6"><IMG border=0 height=4 src="/ehres/Image/dottedlinenav.gif" width=408></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="34" class="small">
      &nbsp;<br>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All 
      Rights Reserved</i>. </td></tr>
</table>
      </div>
      <div align="center">      
  </table>     
</BODY>
</HTML>