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
    <td vAlign="top" align="middle" colspan="2" width="936" bgcolor="#0099cc" height="29">
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
    </td></tr><!--<tr>
    <td vAlign="top" colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></td></tr>
  <tr>-->
  <TR>
    <TD vAlign=top colspan="2" width="100%" height="21" class="small" align="middle">
      <p align="right"><A href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <A href="../signout.asp"><font color="#000000">Logout</font></a></p></TD></TR>
  <TR>
    <TD vAlign=top align=middle width="1" height="109"></TD>
  </center>
    <TD vAlign=top align=middle width="100%" height="109">
      <p align="center"><IMG height=88 alt 
      ="Main 
     Menu" src="../Image/englsbal.gif" width=712 border=0 ><br>
     
     <% 
        set myconn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.Recordset")
				myconn.open connect_string
		       
        'sql = "Exec sp_sa_seltypeofEmploy '" + trim(Session("EmpID")) + "','employ'"        sql = "exec sp_sa_seltypeofEmploy 'ms0022','employ'"                  
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
        
        'if rs.RecordCount >0 then
           tempempid =rs("empid")
           tempempname = rs("empname")
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
           tempstaterecruited = rs("staterecruited")
           'tempappraissal = rs("appraisal")
           tempdatejoin = rs("datejoin")
           tempdateconfirm = rs("dateconfirm")
           tempstatus = rs("jobstatus")
           temptype = rs("emptype")
       'end if
       
       Response.Write tempempid    
         
     'Set webdb = Server.CreateObject("ADODB.Connection")
				'webdb.Open Session("ConnectStr")
			    'Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    'Set webdbCommand = Server.CreateObject("ADODB.Command")
                'ssql ="Exec sp_is_empprofile '" + Session("Regisno") + "','" + trim(Session("EmpID")) + "','Retrieve'"
     
     'Set webdbCommand.ActiveConnection = webdb
			  	'webdbCommand.CommandText = ssql
     %>
     <table border ="1" width ="90%" cellspacing="0" cellpadding="0">
     <tr>
     <td>
      &nbsp;Employee ID <select size="1" name="cboempid" style="FONT-SIZE: 8pt">
      <%response.write "<option selected value=" + trim(tempempid) + ">"  + " " + trim(tempempid) + "</option>"%>
      </select>&nbsp;&nbsp;&nbsp; Employee Name <select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
      <%response.write "<option selected value=" + trim(tempempname) + ">"  + " " + trim(tempempname) + "</option>"%>
      </select>&nbsp;<input type="button" value="Search" name="cmdSearch" onClick="Verify()" onmouseover="this.style.cursor='hand';" class="small">
            <P></P></td></tr></table></p>
     </TD></TR>
  <center>
  <TR>
    <TD vAlign=top align=middle colspan=2 width="714" height="193">
      <div align="center">
      
      <p align="left">&nbsp;&nbsp;&nbsp;&nbsp;</p>
      
      </div>
      <div align="center">
      
      <!--<p align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;--> 
      <div align="center">
      <table border="0" width="110%" cellspacing="0" cellpadding="0" align=center>
      <tr><td>
      <table border="0" width="90%" cellspacing="0" cellpadding="0" height="104" align=center>
        <tr>
          <td width="90%">
            <table border="1" width="100%" cellspacing="0" cellpadding="0" align=center>
              <tr>
                <td width="25%">Division<!--</td>
                <td width="25%">--><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempdivision) + ">"  + " " + trim(tempdivision) + "</option>"%>
                </select></td>
                <td width="25%">Line</td>
                <td width="25%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(templine) + ">"  + " " + trim(templine) + "</option>"%>
                </select></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%" height="21">
            <table border="1" width="100%" cellspacing="0" cellpadding="0" align=center>
              <tr>
                <td width="25%">Department</td>
                <td width="25%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempdepartment) + ">"  + " " + trim(tempdepartment) + "</option>"%>
                </select></td>
                <td width="25%">Group</td>
                <td width="25%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempgroup) + ">"  + " " + trim(tempgroup) + "</option>"%>
                </select></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%" height="21">
            <table border="1" width="100%" cellspacing="0" cellpadding="0" align=center>
              <tr>
                <td width="25%">Section</td>
                <td width="25%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempsection) + ">"  + " " + trim(tempsection) + "</option>"%>
                </select></td>
                <td width="25%">Process</td>
                <td width="25%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempprocess) + ">"  + " " + trim(tempprocess) + "</option>"%>
                </select></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%" height="21">
          </td>
        </tr>
        <!--<tr>
          <td width="100%" height="21">
          </td>
        </tr>-->
      </table>
      <br>
      <table border="0" width="90%" cellspacing="0" cellpadding="0" align=center>
        <tr>
          <td width="100%">
            <table border="1" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="25%">TMS ID</td>
                <td width="25%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(temptmsid) + ">"  + " " + trim(temptmsid) + "</option>"%>
                </select></td>
                <td width="25%">Finance Code</td>
                <td width="25%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempfinid) + ">"  + " " + trim(tempfinid) + "</option>"%>
                </select></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <br></br>
      <!--<p>&nbsp;</p>-->
      <table border="1" width="90%" cellspacing="0" cellpadding="0" align=center>
        <tr>
          <td width="100%">
            <table border="0" width="100%">
              <tr>
                <td width="27%">Job Grade</td>
                <td width="23%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempjobgrade) + ">"  + " " + trim(tempgrade) + "</option>"%>
                </select></td>
                <td width="18%">State Recruited</td>
                <td width="32%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempstaterecruited) + ">"  + " " + trim(tempstaterecruited) + "</option>"%>
                </select></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="1" width="100%">
              <tr>
                <td width="27%">Job Title</td>
                <td width="23%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempjobtitle) + ">"  + " " + trim(tempjobtitle) + "</option>"%>
                </select></td>
                <td width="18%">Status</td>
                <td width="32%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempstatus) + ">"  + " " + trim(tempstatus) + "</option>"%>
                </select></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%">
              <tr>
                <td width="27%">Appraisal Grade</td>
                <td width="23%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
      </select></td>
                <td width="18%">Type</td>
                <td width="32%"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(temptype) + ">"  + " " + trim(temptype) + "</option>"%>
                </select></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%">
              <tr>
                <td width="31%">Date Joined<font size="2"><i> (dd/mm/yyyy)</i></font></td>
                <td width="43%"><input name="T1" size="37" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & tempdatejoin & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
                <td width="6%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%" height="22">
              <tr>
                <td width="31%" height="16">Date Confirmed<font size="2"><i> (dd/mm/yyyy)</i></font></td>
                <td width="43%" height="16"><input name="T1" size="37"<%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & tempdateconfrim & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
                <td width="2%" height="16"></td>
                <td width="25%" height="16"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      </td></tr>
      </table>
     </div>
      </div>
      <div align="center">
      
      <p align="left"></p>
      
      </div>
      <div align="center">
      <p align="left"></p>
      </div>
      <div align="center">
      <p align="left"></p>
      </div>
      <CENTER>
      <p></p>
      </CENTER>
    </TD></TR>
  <TR>
    <TD align=middle colspan=2 width="714" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<BR>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All 
    Rights Reserved</i>. </TD></TR></TBODY></TABLE>
</div></CENTER>
</body>