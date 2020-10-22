<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/AdoVbs.asp"-->
<!-- #include virtual ="/ehres/global/inputSession.asp"-->
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>
<% dim connect_string
       connect_string =Session("ConnectStr")
%>
<SCRIPT LANGUAGE="vbScript">
	function Update()
		if msgbox("Are you sure you want to save? (Yes/No)",vbyesno,"Confirmation") = vbyes then
			document.frmRQInfo.submit()
		end if
	end function

	function validate()
		dim validated
		validated = true
		if Len(document.frmRQInfo.P4.value) <> 1 or (ucase(document.frmRQInfo.P4.value) <> "M" and ucase(document.frmRQInfo.P4.value) <> "F") then
			if msgbox("Sex Must be (M/F) value",vbOKOnly,"Error") = vbOK then
				validated = False
			end if
		elseif len(document.frmRQInfo.P6.value) <> 1 or (ucase(document.frmRQInfo.P6.value) <> "S" and ucase(document.frmRQInfo.P6.value) <> "M" and ucase(document.frmRQInfo.P6.value) <> "D" and ucase(document.frmRQInfo.P6.value) <> "W") then
			if msgbox("Marital Status Must be (S/M/D/W) value",vbOKOnly,"Error") = vbok then
				validated = false
			end if
		elseif UCASE(document.frmRQInfo.P8.value) <> "MALAYSIAN" and UCASE(document.frmRQInfo.P8.value) <> "VIETNAMESE" and UCASE(document.frmRQInfo.P8.value) <> "PERMT. RESIDENT" and UCASE(document.frmRQInfo.P8.value) <> "JAPANESE" and UCASE(document.frmRQInfo.P8.value) <> "INDONESIAN" and UCASE(document.frmRQInfo.P8.value) <> "SINGAPOREAN" then
			if msgbox("Citizenship was not correct,Example:(VIETNAMESE,MALAYSIAN)",vbOKOnly,"Error") = vbok then
				validated = false
			end if
		end if

		if validated = true then
			call update()
		end if
	end function
	
	</SCRIPT>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div align="center">
  <center>
<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%" height="392">
  <TBODY>
    <TD vAlign=top align=middle width="27" height="109"></TD>
    <TD vAlign=top align=middle width="907" height="109">
      <P><IMG height=84 src="../Image/engrecruitment.gif" width=683 border=0 ><BR><BR>
      <TABLE cellSpacing=0 width="100%" border=0>
  <TR>
    <TD vAlign=top align=middle colspan=2 width="936" height="193">
      <div align="center">
        <center>
			<FORM name=frmRQInfo action=UpdateRQInfo.asp method=post>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="1">
        <TBODY>
    <tr>
      <td bgcolor="#ffffff" width="30">&nbsp;</td>
      <td bgcolor="#f3f3f3" width="240"><font class="marineblack">Personal Info</font></td>
      <td bgcolor="#f3f3f3" width="400"><font class="marineblack"></font></td>
    </tr>

            <%     
                    set myconn = server.CreateObject("ADODB.Connection")
			        set rs = server.CreateObject("ADODB.Recordset")
		                myconn.open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=SA;PWD=;Initial catalog=HRDB_SNE;Server=(local);Connect Timeout=900000"
		                
					ssql = "select * from rq_label where type = '1'"
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
					 
			        page_count = rs.pagecount
			        		       
					colour = 0
					rowno = 0 
			        do while not rs.EOF
			        
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			            rowno = rowno + 1
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("description") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=hidden name=PH" + cstr(rowno) + " value='" + rs("ID") + "'><input size=90 type=text name='P" + cstr(rowno) + "'></td>"
				        Response.Write "</tr>"
				       
				        rs.Movenext  
			        loop
			        response.write "<input type=hidden name=txtRowNo1 value=" + cstr(rowno) + ">"
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing 
			       
			%>           
		<tr>
      <td bgcolor="#ffffff" width="30">&nbsp;</td>
      <td bgcolor="#f3f3f3" width="240"><font class="marineblack">Education Info</font></td>
      <td bgcolor="#f3f3f3" width="400"><font class="marineblack"></font></td>
    </tr>

            <%     
                    set myconn = server.CreateObject("ADODB.Connection")
			        set rs = server.CreateObject("ADODB.Recordset")
		                myconn.open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=SA;PWD=;Initial catalog=HRDB_SNE;Server=(local);Connect Timeout=900000"
		                
					ssql = "select * from rq_label where type = '2'"
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
					 
			        page_count = rs.pagecount
			        		       
					colour = 0
					rowno = 0 
			        do while not rs.EOF
			        
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			            rowno = rowno + 1
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("description") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=hidden name=EH" + cstr(rowno) + " value='" + rs("ID") + "'><input size=90 type=text name='E" + cstr(rowno) + "'></td>"
				        Response.Write "</tr>"
				       
				        rs.Movenext  
			        loop
			        response.write "<input type=hidden name=txtRowNo2 value=" + cstr(rowno) + ">"
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing 
			       
			%>
			<tr>
      <td bgcolor="#ffffff" width="30">&nbsp;</td>
      <td bgcolor="#f3f3f3" width="240"><font class="marineblack">Experience Info</font></td>
      <td bgcolor="#f3f3f3" width="400"><font class="marineblack"></font></td>
    </tr>

            <%     
                    set myconn = server.CreateObject("ADODB.Connection")
			        set rs = server.CreateObject("ADODB.Recordset")
		                myconn.open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=SA;PWD=;Initial catalog=HRDB_SNE;Server=(local);Connect Timeout=900000"
		                
					ssql = "select * from rq_label where type = '3'"
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
					 
			        page_count = rs.pagecount
			        		       
					colour = 0
					rowno = 0 
			        do while not rs.EOF
			        
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			            rowno = rowno + 1
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("description") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=hidden name=EXH" + cstr(rowno) + " value='" + rs("ID") + "'><input size=90 type=text name='EX" + cstr(rowno) + "'></td>"
				        Response.Write "</tr>"
				       
				        rs.Movenext  
			        loop
			        response.write "<input type=hidden name=txtRowNo3 value=" + cstr(rowno) + ">"
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing 
			       
			%>           
		<tr>
      <td bgcolor="#ffffff" width="30">&nbsp;</td>
      <td bgcolor="#f3f3f3" width="240"><font class="marineblack">Additional Info</font></td>
      <td bgcolor="#f3f3f3" width="400"><font class="marineblack"></font></td>
    </tr>

            <%     
                    set myconn = server.CreateObject("ADODB.Connection")
			        set rs = server.CreateObject("ADODB.Recordset")
		                myconn.open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=SA;PWD=;Initial catalog=HRDB_SNE;Server=(local);Connect Timeout=900000"
		                
					ssql = "select * from rq_label where type = '4'"
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
					 
			        page_count = rs.pagecount
			        		       
					colour = 0
					rowno = 0 
			        do while not rs.EOF
			        
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			            rowno = rowno + 1
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("description") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=hidden name=AH" + cstr(rowno) + " value='" + rs("ID") + "'><input size=90 type=text name='A" + cstr(rowno) + "'></td>"
				        Response.Write "</tr>"
				       
				        rs.Movenext  
			        loop
			        response.write "<input type=hidden name=txtRowNo4 value=" + cstr(rowno) + ">"
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing 
			       
			%>           
            </TBODY>
            </TABLE><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="1" ID="Table1">
        <TBODY>
			<tr><td bgcolor="#ffffff" width="100">&nbsp;</td><td bgcolor="#ffffff" width="240">&nbsp;</td></tr>
			<tr><td bgcolor="#ffffff" width="100">&nbsp;</td>
			<td bgcolor="#ffffff" width="240"><input type=button name=cmdSave value=Save onclick=Validate()></td>
			</tr>
			</TBODY></table></form></center>
      </div>
    </TD></TR>
    <center>  
    <TD align=center colspan=2 width="936" height="40" class="small"><br>
      <!--&nbsp;<br>
      &nbsp;<BR>--><font class ="small" >Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>.</font></TD></TR></TBODY></TABLE></center>
      </div>
</BODY>




