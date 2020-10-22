<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/ADOVBS.ASP" -->
<% dim connect_string 

connect_string = Session("ConnectStr")%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>eHRES</TITLE>
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
      src="../Image/englsbal.gif" 
     border=0 width="712" height="88"><br>
     
     <% 
        set myconn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.Recordset")
		
		myconn.open connect_string
		       
        sql = "Exec sp_sa_seltypeofEmploy '" + trim(Session("EmpID")) + "','employ'"
        'sql = "exec sp_sa_seltypeofEmploy 'ms0022','employ'"
                  
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
     <table border ="1" width ="78%" cellspacing="0" cellpadding="0">
     <tr>
     <td align="left">
      &nbsp;Employee ID <select size="1" name="cboempid" style="FONT-SIZE: 8pt">
      <%response.write "<option selected value=" + trim(tempempid) + ">"  + " " + trim(tempempid) + "</option>"%>
      </select>&nbsp;&nbsp;&nbsp; Employee Name <select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
      <%response.write "<option selected value=" + trim(tempempname) + ">"  + " " + trim(tempempname) + "</option>"%>
      </select>&nbsp;<input type="button" value="Search" name="cmdSearch" onClick="Verify()" onmouseover="this.style.cursor='hand';" class="small">
            <P></P></td></tr></table>
     </TD></TR>
  <center>
  <TR>
    <TD vAlign=top align=middle colspan=2 width="714" height="193">
      <div align="center">
      
      <p align="left">&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;</p>
      
      </div>
    </CENTER>
      <div align="center">
      
      <table border="1" width="90%">
        <tr>
          <td width="100%" align="center">
            <table border="1" width="73%" height="375">
              <tr>
                <td width="50%" height="19">
                  <p align="right">Division</td>
  <center>
  <td width="50%" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempdivision) + ">"  + " " + trim(tempdivision) + "</option>"%>
                </select></td>
                </tr>
              </CENTER>
              <tr>
                <td width="50%" height="19">
                  <p align="right">Department</td>
  <center>
  <td width="50%" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempdepartment) + ">"  + " " + trim(tempdepartment) + "</option>"%>
                </select></td>
                </tr>
                <tr>
                  <td width="50%" align="right" height="19">Section</td>
                </CENTER>
                <td width="50%" align="right" height="19">
                  <p align="left"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempsection) + ">"  + " " + trim(tempsection) + "</option>"%>
                </select></td>
              </tr>
  <center>
  <tr>
    <td width="50%" align="right" height="19">TMS ID</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(temptmsid) + ">"  + " " + trim(temptmsid) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Job Grade</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempjobgrade) + ">"  + " " + trim(tempgrade) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Job Title</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempjobtitle) + ">"  + " " + trim(tempjobtitle) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Appraisal Grade</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
      </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Line</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(templine) + ">"  + " " + trim(templine) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Group</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempgroup) + ">"  + " " + trim(tempgroup) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Process</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempprocess) + ">"  + " " + trim(tempprocess) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Finance Code</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempfinid) + ">"  + " " + trim(tempfinid) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">State Recruited</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempstaterecruited) + ">"  + " " + trim(tempstaterecruited) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Status</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(tempstatus) + ">"  + " " + trim(tempstatus) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Type</td>
    <td width="50%" align="left" height="19"><select size="1" name="cboStatus" style="FONT-SIZE: 8pt">
                <%response.write "<option selected value=" + trim(temptype) + ">"  + " " + trim(temptype) + "</option>"%>
                </select></td>
  </tr>
  <tr>
    <td width="50%" align="right" height="19">Date Joined<font size="2"><i>  (dd/mm/yyyy</i></font></td>
  </CENTER>
  <td width="50%" align="right" height="19">
    <p align="left"><input name="T1" size="20" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & tempdateconfrim & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
              </tr>
  <center>
  <tr>
    <td width="50%" align="right" height="19">Date Confirmed<font size="2"><i> (dd/mm/yyyy)</i></font></td>
  </CENTER>
  <td width="50%" align="right" height="19">
    <p align="left"><input name="T1" size="20" <%
             
             'if request.form("txtname2") ="" then 'temptypecomm ="Search" or request.form("txtPartNo") ="" then
                response.write " value='" & tempdateconfrim & "'"                      
             'else 
                'response.write " value='" & request("txtpartno") & "'"
             'end if %>></td>
              </tr>
            </table>
  <center>
  <p>&nbsp;</p>
  <p>&nbsp;</td>
          </tr>
        </table>
      
      </div>
      <div align="center">
      
      <p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
      
      </div>
      <div align="center">
      
      <p align="left">&nbsp;</p>
      
      </div>
      <div align="center">
      
      <p align="left">&nbsp;</p>
      
      </div>
      <div align="center">
      
      <p align="left">&nbsp;</p>
      
      </div>
      <div align="center" style="width: 863; height: 387">
      
      <!--<p align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;--> 
      </div>
      </td></tr>
      </table>
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
      <CENTER></CENTER>
<table>
  <TR>
    <TD align=middle colspan=2 width="714" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<BR>Copyright © 1997-2000 Software
      Factory Sdn Bhd <i>All Rights Reserved</i>. </TD></TR></CENTER>
</body>




