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
       
	<SCRIPT LANGUAGE="vbscript">
<!--
	function Back()
		document.frmPersonalInfo.action = "../Maintenance.asp"
		document.frmPersonalInfo.submit()
	end function
	
	function Change()
		document.frmPersonalInfo.txtaction.value = ""
		document.frmPersonalInfo.action = "PersonalProfileVw.asp"
		document.frmPersonalInfo.submit()
	end function
	
	function Update()
		document.frmPersonalInfo.txtaction.value = "UPD"
		document.frmPersonalInfo.action = "UpdateHQRQInfo.asp"
		document.frmPersonalInfo.submit()
	end function
	
	function Delete()
		if msgbox("Are you sure you want to Delete? (Yes/No)",vbyesno,"Confirmation") = vbyes then 
			document.frmPersonalInfo.txtaction.value = "DEL"
			document.frmPersonalInfo.action="UpdateHQRQInfo.asp"
			document.frmPersonalInfo.submit()
		end if
	End function	
	
	function check()
		if document.frmPersonalInfo.Hire.checked = true then
			document.frmPersonalInfo.Hire.value = 1
		else
			document.frmPersonalInfo.Hire.value = 0
		end if
		document.frmPersonalInfo.action="PersonalProfileVw.asp"
		document.frmPersonalInfo.submit()
	end function
	
	function import()
		document.frmPersonalInfo.action = "ImportPersonalInfo.asp"
		document.frmPersonalInfo.submit()
	end function
	// -->
	</SCRIPT>

<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div align="center">
  <center>
<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%" height="392">
  <TBODY>
  <TR>
    <TD vAlign=top align=middle width="27" height="109"></TD>
    <TD vAlign=top align=middle width="907" height="109">
      <P><IMG height=84 src="../Image/engRecruitment.gif" width=683 border=0 ><BR><BR>
      <FORM name=frmPersonalInfo action=UpdateHQRQInfo.asp method=post>
      <TABLE cellSpacing=0 width="100%" border=0>
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <TD>
           <FONT class=small><b>Recruitment ID &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
            <select name=cboRQID style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px"  onchange=Change()>  
       			<%  dim tmpRQID
       			

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="select RQID from rq_Profile where flag = 0 group by RQID order by RQID"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
			  	
	  			
	  			tmpRQID = ""
	  			tmpRQID = Request.form("cboRQID")
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("RQID")) = Request.form("cboRQID") ) or ( trim(webdbRecordset.Fields("RQID")) = tmpRQID )then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("RQID")) + ">"  + " " + trim(webdbRecordset.Fields("RQID")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("RQID")) + ">"  + " " + trim(webdbRecordset.Fields("RQID")) + "</option>"
 				    end if
 				    
 				    if tmpRQID = "" then
					      tmpRQID = trim(webdbRecordset.Fields("RQID"))
					      
					end if   
				   webdbRecordset.MoveNext  
		        loop       
			%></select>&nbsp;&nbsp;&nbsp; 
           </TD>
        </TR>
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <% if request.Form("Hire") = "1" then %>
           <TD><input type=checkbox name=Hire value=Hire onclick=check() checked><font class=small>Hire</font></td></tr>
           <%else%>
           <TD><input type=checkbox name=Hire value=Hire onclick=check()><font class=small>Hire</font></td></tr>
           <%end if%>
           <% if request.Form("hire") = "1" then
           response.Write "<tr><td WIDTH='10%'>&nbsp;</TD>"
           response.Write "<td>Employee ID :&nbsp;&nbsp;<input type=text name=empid size=50></td></tr>"
           end if
           %>
			
  <tr>
  <td>&nbsp;</td>
  </tr>
  <TR>
    <TD vAlign=top align=middle colspan=3 width="936" height="193" >
      <div align="center">
        <center>
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
		                
					ssql = "exec sp_WRQ_SelRQProfile '" + tmpRQID + "','1'"
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
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=hidden name=PH" + cstr(rowno) + " value='" + rs("ID") + "'><input size=90 type=text name='P" + cstr(rowno) + "' value='" + rs("Answer") +"'></td>"
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
		                
					ssql = "exec sp_WRQ_SelRQProfile '" + tmpRQID + "','2'"
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
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=hidden name=EH" + cstr(rowno) + " value='" + rs("ID") + "'><input size=90 type=text name='E" + cstr(rowno) + "' value='" + rs("Answer") +"'></td>"
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
		                
					ssql = "exec sp_WRQ_SelRQProfile '" + tmpRQID + "','3'"
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
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=hidden name=EXH" + cstr(rowno) + " value='" + rs("ID") + "'><input size=90 type=text name='EX" + cstr(rowno) + "' value='" + rs("Answer") +"'></td>"
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
		                
					ssql = "exec sp_WRQ_SelRQProfile '" + tmpRQID + "','4'"
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
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=hidden name=AH" + cstr(rowno) + " value='" + rs("ID") + "'><input size=90 type=text name='A" + cstr(rowno) + "' value='" + rs("Answer") +"'></td>"
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
            </TABLE></center>
      </div>
    </TD></TR>
    <tr>
          <td height="19"></td>
        </tr>
        <tr>
          <td width="6%" height="19"></td>
          <% if request.Form("Hire") = "1" then%>
          <td width="94%" height="19"><input type=hidden name=txtAction value= ><input type=button value="Import" name=cmdImport onclick="Import()" class="small"><input type="button" value="Delete" name="cmdDelete" onclick="Delete()" class="small"><input type="button" value="Back" name="cmdBack" onclick="Back()" class="small"></td>
          <%else%>
          <td width="94%" height="19"><input type=hidden name=txtAction value= ><input type=button value="Update" name=cmdUpdate onclick="Update()" class="small"><input type="button" value="Delete" name="cmdDelete" onclick="Delete()" class="small"><input type="button" value="Back" name="cmdBack" onclick="Back()" class="small"></td>
          <%end if%>
        </tr>
    <center>  </form>
    <TD align=center colspan=2 width="936" height="40" class="small"><br>
      <!--&nbsp;<br>
      &nbsp;<BR>--><font class ="small" >Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>.</font></TD></TR></TBODY></TABLE></center>
      </div>
</BODY>




