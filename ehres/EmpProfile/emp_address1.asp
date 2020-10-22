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
        sql = "Exec sp_sa_selEmpAddress 'ms0036','id'"
        'Response.Write sql       
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText
        
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
     
     <% 
        set myconn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.Recordset")    
		
		myconn.open connect_string
		       
        'sql = "Exec sp_is_empprofile '" + trim(Session("EmpID")) + "','" + trim(Session("Regisno")) + "','Retrieve'"
        sql = "Exec sp_is_empprofile 'ms0036','185612-k','Retrieve'"
        'Response.Write sql       
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
      	
        'if rs.RecordCount > 0 then
           tempempid =rs("empid")
           tempempname = rs("empname")
	%>
     
     <table border ="0" width ="90%" cellSpacing="1" cellPadding="2" height="34">
     <tr>
		<td height="28"><font class=small>Employee ID : <%Response.Write tempempid %></td>
		<td height="28"><font class=small>Employee Name : <%Response.Write tempempname%></font></td>
     </tr>
     </table> 
     <!--<table border ="0" width ="90%">
     <tr>
     <td>
      &nbsp;<font class=small>Employee ID </font><select size="1" name="cboempid" style="font-size: 8pt">
      <%response.write "<option selected value=" + trim(tempempid) + ">"  + " " + trim(tempempid) + "</option>"%>
      </select>&nbsp;&nbsp;&nbsp;<font class=small>Employee Name</font><select size="1" name="cboStatus" style="font-size: 8pt">
      <%response.write "<option selected value=" + trim(tempempname) + ">"  + " " + trim(tempempname) + "</option>"%>
      </select>&nbsp;<input type="button" value="Search" name="cmdSearch" onClick="Verify()" onmouseover="this.style.cursor='hand';" class="small"></p>
     </td></tr>
     </table>--> 
  </TD></TR>
  <center>
  <TR>
    <TD vAlign=top align=center colspan=2 width="90%" height="193">
      <div align="center">
      
      <!--<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;-->
      
      </div>
      <div align="center">
      <P>&nbsp;</P>
      <table border="0" width="90%" cellspacing="0" cellpadding="0" align=center >
        <tr>
          <td width="100%" bgcolor=#000080>
            <b><font face=Verdana size=2 color="#FFFFFF">&nbsp;Present Address</font></b></td>
          </td>
        </tr>
        <tr>
          <td width="100%"><br></br>
            <table border="0" width="100%">
              <tr>
                <td width="18%"><font face=Verdana size=2 >Address</font></td>
                <td width="82%"><input type="text" name="txtpraddress" size="81" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & praddress & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%">
              <tr>
                <td width="17%"><font face=Verdana size=2 >Post Code</font></td>
                <td width="33%"><input type="text" name="txtprpostcode" size="31" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & prpostcode & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
                <td width="12%" align="center"><font face=Verdana size=2 >City</font></td>
                <td width="38%"><input type="text" name="txtprcity" size="35" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & prcity & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%">
              <tr>
                <td width="17%"><font face=Verdana size=2 >State</font></td>
                <td width="33%"><select size="1" name="cbostate" style="font-size: 8pt">
                <%response.write "<option selected value=" + trim(prstate) + ">"  + " " + trim(prstate) + "</option>"%> 
                </select></td>
                <td width="12%" align="right"><font face=Verdana size=2 >Country</font></td>
                <td width="38%"><select size="1" name="cbocountry" style="font-size: 8pt">
                <%response.write "<option selected value=" + trim(prcountry) + ">"  + " " + trim(prcountry) + "</option>"%>
                </select></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%" height="22">
              <tr>
                <td width="17%" height="16"><font face=Verdana size=2 >Tel No</font></td>
                <td width="33%" height="16"><input type="text" name="txtprtelno" size="31" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & prtelno & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
                <td width="12%" height="16"></td>
                <td width="38%" height="16"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <!--<br>-->
      <P>&nbsp;</P>
      <table border="0" width="90%" cellspacing="0" cellpadding="0" align=center >
        <tr>
          <td width="100%" bgcolor=#000080 height="50%"><!--<input type="checkbox" name="C1" value="ON">-->
            <b><font face=Verdana size=2 color="#FFFFFF">&nbsp;Permanent Address</font></b></td>
          </td>
        </tr>
        <tr>
          <td width="100%"><br></br>
            <table border="0" width="100%">
              <tr>
                <td width="18%" ><font face=Verdana size=2 >Address</font></td>
                <td width="82%"><input type="text" name="txtpmaddress" size="81" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & pmaddress & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%">
              <tr>
                <td width="17%"><font face=Verdana size=2 >Post Code</font></td>
                <td width="33%"><input type="text" name="txtpmpostcode" size="31" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & pmpostcode & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
                <td width="12%" align="center"><font face=Verdana size=2 >City</font></td>
                <td width="38%"><input type="text" name="txtpmcity" size="35" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & pmcity & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%">
              <tr>
                <td width="17%"><font face=Verdana size=2 >State</font></td>
                <td width="33%"><select size="1" name="cboStatus" style="font-size: 8pt">
                <%response.write "<option selected value=" + trim(pmstate) + ">"  + " " + trim(pmstate) + "</option>"%> 
                </select></td>
                <td width="12%" align="right"><font face=Verdana size=2 >Country</font></td>
                <td width="38%"><select size="1" name="cboStatus" style="font-size: 8pt">
                <%response.write "<option selected value=" + trim(pmcountry) + ">"  + " " + trim(pmcountry) + "</option>"%> 
                </select></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%" height="22">
              <tr>
                <td width="17%" height="16"><font face=Verdana size=2 >Tel No</font></td>
                <td width="33%" height="16"><input type="text" name="txtpmtelno" size="31" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & pmtelno & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
                <td width="12%" height="16"></td>
                <td width="38%" height="16"></td>
              </tr>
            </table>
          </td>
        </tr>
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
      
      
</BODY>