<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<%
Response.Buffer = true
%>

<HTML><HEAD><TITLE>Approval - Staff Confirmation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252"><LINK 
href="../css/login.css" type=text/css rel=stylesheet>
<SCRIPT language=javascript type=text/javascript>
<!--
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos)
{
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=no,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);
if(win.focus){win.focus();}}
// -->
</SCRIPT>
<!--Place this script anywhere in a page.--><!--NOTE: You do not need to modify this script.-->
<SCRIPT language=JavaScript>

	function Verify()
	{
		msg = "";
		m = true;
		n = true;

		m = CheckDate('txtDate1');
		if (!m)
		{
			window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
			document.forms[0].reset()
		}
		

		m = CheckDate('txtDate2');
		if (!m)
		{
			window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
			document.forms[0].reset()
		}
		else
		{
		    
		    eval("document.forms[0]." + "txtPage" + ".value = 'NEW'");
		    document.forms[0].submit();
		}		
	}

	function CheckDate(x)
	{
		o = true;

		if ( eval("document.forms[0]." + x + ".value.length == 8") )
		{
			day = eval("document.forms[0]." + x + ".value.substring(0,2)");
			month = eval("document.forms[0]." + x + ".value.substring(2,4)");
			year = eval("document.forms[0]." + x + ".value.substring(4,8)");
			o = o && CheckDay(day, month, year);
			o = o && (month < 13);
		}
		else
		{
		  if ( eval("document.forms[0]." + x + ".value.length == 10") )
		  {
			day = eval("document.forms[0]." + x + ".value.substring(0,2)");
			month = eval("document.forms[0]." + x + ".value.substring(3,5)");
			year = eval("document.forms[0]." + x + ".value.substring(6,10)");
			o = o && CheckDay(day, month, year);
			o = o && (month < 13);
		  }
                  else 
                  { 
 		    if ( eval("document.forms[0]." + x + ".value.length == 0") )
                    {
                        o = false;
                    }
                    else 
                    {
			o = false;
                    }
                  }
		}
	
		if (o) return true;
		else return false;
	}

	function CheckDay(dd, mm, yy)
	{
		MaxDay = new Array (31,28,31,30,31,30,31,31,30,31,30,31);
	
		if (yy%4 == 0) MaxDay[1]++;

		if (dd <= MaxDay[mm-1]) return true;
	}
	
	</SCRIPT>


<script language="vbscript">
<!--
function CheckStatus(vRow)
    
 	if document.All("cbostatus" + vrow).value = "C" or document.All("cbostatus" + vrow).value = "O" then
	   ssql= " document.frmStaffConfirm.D" + vRow + ".disabled=false"
	else   
	   ssql= " document.frmStaffConfirm.D" + vRow + ".disabled=true"
	end if   

	execute ssql
    
end function


function DisableAll()
    
	dim rowcount 
	dim ssql
	dim maxrow
    dim approve
	
	maxrow = document.frmStaffConfirm.All("txtRowNo").value
	
	do until rowcount = cint(maxrow)
  	   rowcount = rowcount + 1    
	   ssql= " document.frmStaffConfirm.D" + cstr(rowcount-1) + ".disabled=true"
	   execute ssql
	loop
    
end function



function Update()
	document.frmStaffConfirm.submit()
end function
		    
// -->
</script>



<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginwidth="0" marginheight="0">
<DIV align=center>
<CENTER>
<%

dim maxrow 
dim rowcount
			 
if Request.form("txtAction")="UPD"then
   Set webdb = Server.CreateObject("ADODB.Connection")
 	   webdb.Open Session("ConnectStr")
   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
   Set webdbCommand = Server.CreateObject("ADODB.Command")
   Set webdbCommand.ActiveConnection = webdb
	
   maxrow = request.form("txtRowNo")

   do until rowcount = cint(maxrow)
      rowcount = rowcount + 1
      
      if Request.Form("cbostatus"+ cstr(rowcount-1)) <> "" then
         if Request.Form("D"+ cstr(rowcount-1)) = "" then
         ssql = "Exec sp_Wis_InsUpdDelEmpConfirm """ & _
                      Session("EmpID") & """,""" & _
   	                  Request.Form("O"+ cstr(rowcount-1)) & """,""" & _
	                  "01/01/1900" & """,""" & _
	                  Request.Form("cbostatus"+ cstr(rowcount-1)) & """"
         else
         ssql = "Exec sp_Wis_InsUpdDelEmpConfirm """ & _
                      Session("EmpID") & """,""" & _
   	                  Request.Form("O"+ cstr(rowcount-1)) & """,""" & _
	                  Request.Form("D"+ cstr(rowcount-1)) & """,""" & _
	                  Request.Form("cbostatus"+ cstr(rowcount-1)) & """"
	     webdbCommand.CommandText = ssql
		 webdb.Execute webdbCommand.CommandText
		 end if
	  end if
   loop
		      
   'response.redirect "/ehres/main.asp"
end if	
%>

<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD vAlign=top align=middle width=936 bgColor=#0099cc colSpan=2 height=29>
      <DIV align=center>
      <CENTER>
      <TABLE width="100%" border=0>
        <TBODY>
        <TR>
          <TD width="3%"></TD>
          <TD width="23%"><FONT class=marinewhite>Employee ID : <%response.write session("EmpID")%></FONT></TD>
          <TD width="74%"><FONT class=marinewhite>Name : <%response.write session("EmpName")%>  
      </FONT></TD></TR></TBODY></TABLE></CENTER></DIV></TD></TR>
  <TR>
    <TD class=small vAlign=top align=middle width="100%" colSpan=2 height=21>
      <P align=right><A href="http://10.10.10.4/ehres/main.asp"><FONT 
      color=#000000>Home</FONT></A>&nbsp;&nbsp;&nbsp; |&nbsp;&nbsp;&nbsp; <A 
      href="http://10.10.10.4/ehres/signout.asp"><FONT 
      color=#000000>Logout</FONT></A></P></TD></TR>
  <TR>
    <TD vAlign=top align=middle width=907 height=109><IMG 
      alt="Main Menu" src="../Image/engisapp.gif" 
      border=0><BR>&nbsp;</TD></TR>
  <TR>
    <TD align=right width="100%"></TD></TR></TBODY></TABLE>
<TABLE height=121 cellSpacing=0 width="96%" border=0>
  <TBODY>
  <TR>
    <TD width="164%" colSpan=2 height=44>
      <TABLE height=1 width="100%" border=0>
        <TBODY>
        <TR>
          <TD width="100%" height=1>
            <FORM name=frmStaffConfirm action=app_confirm.asp method=post>
            <%
               if request("txtpage") = "NEW" then
                  Response.Write "<input type='hidden' name=txtPage values =NEW>" 
               else
                  Response.Write "<input type='hidden' name=txtPage values =OLD>" 
               end if   
            %>
            <P>&nbsp;<FONT class=small>&nbsp;&nbsp;&nbsp;&nbsp; Date Form (dd/mm/yyyy)&nbsp;<INPUT class=small size=9 
            name=txtDate1 <% if Request("txtDate1") <> "" then Response.Write " value = " + Request("txtDate1") end if%> >&nbsp;&nbsp;&nbsp; to&nbsp;&nbsp; (dd/mm/yyyy)<INPUT class=small 
            size=9 name=txtDate2 <% if Request("txtDate2") <> "" then Response.Write " value = " + Request("txtDate2") end if%>></FONT>&nbsp;&nbsp; <B><INPUT class=small onmouseover="this.style.cursor='hand';" onclick=Verify() type=button value=Search name=cmdSearch></B></P>
            <TABLE cellSpacing=0 cellPadding=0 width="1063" border=0>
              <TBODY>
              <tr>
                <TD height="20" align=left width="12%" bgColor=#f3f3f3><font class="marineblack"><b>Status</b></font></TD>
                <TD height="20" align=left width="7%" bgColor=#f3f3f3><font class="marineblack"><b>Emp ID</b></font></TD>
                <TD height="20" align=left width="20%" bgColor=#f3f3f3><FONT class=marineblack><B>Name</B></FONT></TD>
                <TD height="20" align=left width="15%" bgColor=#f3f3f3><font class="marineblack"><b>Department</b></font></TD>
                <TD height="20" align=left width="10%" bgColor=#f3f3f3><font class="marineblack"><b>Job Grade</b></font></TD>
                <TD height="20" align=left width="16%" bgColor=#f3f3f3><font class="marineblack"><b>Position</b></font></TD>
                <TD height="20" align=left width="10%" bgColor=#f3f3f3><font class="marineblack"><b>Date Join</b></font></TD>
                <TD height="20" align=left width="10%" bgColor=#f3f3f3><font class="marineblack"><b>Est. Date Confirm</b></font></TD>
              </tr>
              <%
              
                dim ssql, count, colour
                dim rec, recCount, recNo
                dim mode, cbopage

				Set webdb = Server.CreateObject("ADODB.Connection")
				    webdb.Open Session("ConnectStr")
				Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
				Set webdbCommand = Server.CreateObject("ADODB.Command")

		        If Request("txtDate1") <> "" And Request("txtDate2") <> "" Then
				   ssql = "Exec sp_wis_selconfimemployee """ + Request("txtDate1") + """,""" + Request("txtDate2") + """,""" + session("EmpID") + """"
				else   
				   ssql = "Exec sp_wis_selconfimemployee '','',""" + session("EmpID") + """"
				end if   
				
		        Set webdbCommand.ActiveConnection = webdb
				    webdbCommand.CommandText = ssql
				    webdbRecordset.Open webdbCommand, , 1  , 3 

                colour = 0
                
                if not webdbRecordset.EOF then
                   recCount = webdbRecordset.fields(7)
                end if
                
                rec = 1            
                
                cbopage = request("cbopage")
                
                if request("txtpage") = "NEW" then
                   cbopage = ""
                end if   

	 		    if cbopage <> ""  then
	 		       recNo = (20 * (cint(cbopage)-1)) + 1
	 		    else   
	 		       recNo = 1    
	 		    end if 


	 		    
	 	        Do Until webdbRecordset.EOF or rec = recNo + 20
	 	           if count = 1 then
				      colour = " bgcolor='#eeeeee'"
				   else
				      colour = ""
				   end if
                   
                   if rec >=recNo then   
				      Response.Write "<TR><font class=small>"
                      Response.Write "<TD height='25' align=left " + colour + "><font class=small><select onchange='CheckStatus(""" + cstr(rec-recno) + """)' name='cboStatus" + cstr(rec-recNo) + "'>"
                      Response.Write "<option value='' selected> </option>"
                      Response.Write "<option value='C'>Confirm</option>"
                      Response.Write "<option value='X1'>Extend 1 month</option>"
                      Response.Write "<option value='X2'>Extend 2 months</option>"
                      Response.Write "<option value='X3'>Extend 3 months</option>"
                      Response.Write "<option value='X4'>Extend 4 months</option>"
                      Response.Write "<option value='X5'>Extend 5 months</option>"
                      Response.Write "<option value='X6'>Extend 6 months</option>"
                      Response.Write "<option value='O'>Terminate</option>"
                      Response.Write "</select></TD>"
                      Response.Write "<TD height='25' " + colour + "><font class=small>" + webdbRecordset.fields(0) + " <input type='hidden' name=O" + cstr(rec-recNo) + " value=" + webdbRecordset.fields(0) + "></TD>"
                      Response.Write "<TD height='25' " + colour + "><font class=small>" + webdbRecordset.fields(1) + "</TD>"
                      Response.Write "<TD height='25' " + colour + "><font class=small>" + webdbRecordset.fields(2) + "</TD>"
                      Response.Write "<TD height='25' " + colour + "><font class=small>" + webdbRecordset.fields(3) + "</TD>"
                      Response.Write "<TD height='25' " + colour + "><font class=small>" + webdbRecordset.fields(4) + "</TD>"
                      Response.Write "<TD height='25' " + colour + "><font class=small>" + webdbRecordset.fields(5) + "</TD>"
                      'Response.Write "<TD height='25' " + colour + "><font class=small>" + webdbRecordset.fields(6) + "</TD>"
                      Response.Write "<TD height='25' " + colour + "><font class=small><input maxlength=10 type='text' size='10' name=D" + cstr(rec-recNo) + " value='" + cstr(webdbRecordset.fields(6)) + "'> </TD>"
                      
                      Response.Write "</font></TR>"
                   end if
                   webdbRecordset.MoveNext
                   count = abs(count - 1)

                   rec = rec+ 1
                loop
                
                
                if recCount > 20 then
                   Response.Write "<tr><TD> <input type=hidden name=txtRowNo value=" + cstr(rec-recNo ) + "></TD><TD><input type=hidden name=txtAction value='UPD'</TD><TD><B><INPUT class=small onclick=Update() type=button value=Update name=cmdUpdate></B></TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD>"
                
                   Response.Write "<TD height='25' align=left ><font class=small><select name='cboPage' "
                   Response.Write "onchange='document.forms[0].submit();'"
                   Response.Write ">"

                                      
                   if reccount mod 20 > 0 then
                      mode = 1
                   else   
                      mode = 0
                   end if
                      
                   count = 1
                   do until count >= (recCount/20)  + mode
                      if cstr(count) = cbopage then
                         Response.Write "<option value='" + cstr(count) + "' selected>Page "+ cstr(count) + "</option>"
                      else   
                         Response.Write "<option value='" + cstr(count) + "' >Page "+ cstr(count) + "</option>"
                      end if    
                      count = count + 1
                   loop
                   Response.Write "</select></TD>"
                   Response.Write "</tr>"

                end if   
              %>
              </TBODY></TABLE>
              </FORM></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></CENTER></DIV></BODY></HTML>
