<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->

<%
Response.Buffer = true

Dim ssql
Dim i
Dim colour
Dim count

%>

<html>

<head>
<link rel="stylesheet" type="text/css" HREF="../css/login.css">

<script language="javascript" type="text/javascript">
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
</script>

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<!--
			document.frmApplyLeave.txtDate1.value = ""
			document.frmApplyLeave.txtDate1.focus() -->
<title>Leave Application</title>
	<script LANGUAGE="JavaScript">

	function Verify()
	
	{
		msg = "";
		m = true;
		n = true;
		lockdate =""
		
		//document.write(lockdate)
		m = CheckDate('txtDate1');
		if (!m)
		{
			window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
			document.frmapplyleave.txtDate1.value = ""
			document.frmapplyleave.txtDate1.focus() 
		}
		
		if (m)
		{
		   m = CheckDate('txtDate2');
		   if (!m)
		      {
			    window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
   			    document.frmapplyleave.txtDate2.value = ""
			    document.frmapplyleave.txtDate2.focus()

		      }
	    }	
	    var lockdate = new Date(document.frmapplyleave.txtLockDate.value);
        var dateFrom = new Date(document.frmapplyleave.txtDate1.value.substr(2,2) + "/" + document.frmapplyleave.txtDate1.value.substr(0,2) + "/" + document.frmapplyleave.txtDate1.value.substr(4,4));
        var dateTo = new Date(document.frmapplyleave.txtDate2.value.substr(2,2) + "/" + document.frmapplyleave.txtDate2.value.substr(0,2) + "/" + document.frmapplyleave.txtDate2.value.substr(4,4))
        //document.write (dateFrom);
                
	    if (m)
	    {
	      if (dateFrom < lockdate || dateTo < lockdate )
	      {
	        window.alert("Cannot Apply Leave Before LockDate!");
	        document.frmapplyleave.txtDate1.value = ""
	        document.frmapplyleave.txtDate2.value = ""
	        document.frmapplyleave.txtDate1.focus() 
	      }
	      else 
		    document.forms[0].submit();
	    }
	}

	function CheckDate(x)
	{
		o = true;
		
		if ( eval("document.forms[0]." + x + ".value.length == 8"))
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
	
	</script>
		
</head>
<%

   Set webdb = Server.CreateObject("ADODB.Connection")
       webdb.Open Session("ConnectStr")
   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
   Set webdbCommand = Server.CreateObject("ADODB.Command")
   
   ssql ="select lockdate from sa_locktransaction" 

   Set webdbCommand.ActiveConnection = webdb
   webdbCommand.CommandText = ssql
   webdbRecordset.Open webdbCommand,,1 , 3
   
   'if Request.Form("txtlockdate") ="Pass" then
'		Response.Write "pass1"
'   end if	
   tempLockDate = webdbRecordset.fields("lockdate")
      
%>
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
       
  </font></td>
                  <td width="37%"><font class="marinewhite">Name : <%   '   changePass.asp
          response.write session("EmpName")
                    %>
                    </font></td>
		<td width="37%"><font class="marinewhite">Organisation Name : <%   '   changePass.asp
          response.write session("Organname")
                    %>  

    
    </font>
          </td>
        </tr>
      </table>
        </center>
      </div>
    </td></tr>
  <tr>
    <td vAlign="top" colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></td></tr>
  <tr>
    <td vAlign="top" align="center" width="907" height="109"><img src="../Image/englsappl.gif" border="0" width="703" height="87"><br>
      &nbsp;</td></tr>

  <tr>
    <td width="100%" align="right"></td>
  </tr>
  
  
</table>

<table border="0" width="96%">
    <td width="97%" height="28" align="left">    
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.txtDate1.value == "")
  {
    alert("Please enter a value for the \"txtDate1\" field.");
    theForm.txtDate1.focus();
    return (false);
  }

  if (theForm.txtDate1.value.length < 8)
  {
    alert("Please enter at least 8 characters in the \"txtDate1\" field.");
    theForm.txtDate1.focus();
    return (false);
  }

  if (theForm.txtDate1.value.length > 10)
  {
    alert("Please enter at most 10 characters in the \"txtDate1\" field.");
    theForm.txtDate1.focus();
    return (false);
  }

  if (theForm.txtDate2.value == "")
  {
    alert("Please enter a value for the \"txtDate2\" field.");
    theForm.txtDate2.focus();
    return (false);
  }

  if (theForm.txtDate2.value.length < 8)
  {
    alert("Please enter at least 8 characters in the \"txtDate2\" field.");
    theForm.txtDate2.focus();
    return (false);
  }

  if (theForm.txtDate2.value.length > 10)
  {
    alert("Please enter at most 10 characters in the \"txtDate2\" field.");
    theForm.txtDate2.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="insapplyleave.asp" name="frmapplyleave" >
<!--
<script language="vbscript">

function CheckAdvance() 
	
	ssql="if " + "document.frmApplyLeave.chkAdvance.checked = false then" + chr(10) 
	ssql= ssql + " document.frmApplyLeave.txtAdvance.value=0" + chr(10)
	ssql=ssql + "end if"
	
	execute ssql
	
	ssql=""
	ssql="if " + "document.frmApplyLeave.chkAdvance.checked = true then" + chr(10) 
	ssql= ssql + " document.frmApplyLeave.txtAdvance.value=1" + chr(10)
	ssql=ssql + "end if"
		
	execute ssql

'	document.frmApplyLeave.submit()

End function	
// -->
<%
function CheckOthers() 
	
	ssql="if " + "document.frmApplyLeave.chkOthers.checked = false then" + chr(10) 
	ssql= ssql + " document.frmApplyLeave.txtOthers.value=0" + chr(10)
	ssql=ssql + "end if"

	execute ssql
	
	ssql=""
	ssql="if " + "document.frmApplyLeave.chkOthers.checked = true then" + chr(10) 
	ssql= ssql + " document.frmApplyLeave.txtOthers.value=1" + chr(10)
	ssql=ssql + "end if"
	
	execute ssql
    
'	document.frmApplyLeave.submit()
End function	%>
     
<!-- remark by wong 28/05/2001-->
<!-- <p><input type="checkbox" onclick="CheckAdvance()" name="chkAdvance" value="ON">&nbsp;<font class="small"> Apply Advance Leave if leave balance is not        available.</font> <input type="hidden" name="txtAdvance" value="0"></p>  -->
<!-- end remark by wong 28/05/2001-->
        <p><input  type="checkbox" onclick = "CheckOthers()" name="chkOthers" value="ON" checked>&nbsp; <font class="small">Apply
        others leave type except Advanced Leave if leave balance is not available.</font> <input type="hidden" name="txtOthers" value="0">       
        </p>
        
        <p>&nbsp; </p>
        
        <p><font class="small">Leave Type&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <select size="1" name="cboLeaveType" class="small">

                <%
                   dim LeaveType
                   dim selected
                   
                   selected = false
                   
                   LeaveType = request.form("cboLeaveType")

                   Set webdb = Server.CreateObject("ADODB.Connection")
                       webdb.Open Session("ConnectStr")
                   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
                   Set webdbCommand = Server.CreateObject("ADODB.Command")
                   
				  ssql = "Exec sp_Wls_selLeaveIDDesc """ + Session("Regisno") + """,""" + Session("EmpID") + """, """ + Session("CurrentDate") + """, '','1','0','INDIVIDUAL'"

				      Set webdbCommand.ActiveConnection = webdb
				          webdbCommand.CommandText = ssql
                                         webdbRecordset.Open webdbCommand,,1 , 3
          
					   i = 1
					   count = 1
	  					
				       Do Until webdbRecordset.EOF
				          If LeaveType = "" and count = 1 Then
				 	          response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("leaveid")) + "'>" + cstr(webdbRecordset.Fields("description")) + "</OPTION> "
							Elseif LeaveType = webdbRecordset.Fields("leaveid") then
					          response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("leaveid")) + "'>" + cstr(webdbRecordset.Fields("description")) + "</OPTION> "
					          selected = true
						 	else   
						       response.write "<OPTION value='" + cstr(webdbRecordset.Fields("leaveid")) + "'>" + cstr(webdbRecordset.Fields("description")) + "</OPTION> "
							End If
							i = i + 1
							count = count + 1
							webdbRecordset.MoveNext
					   loop
                                      webdbRecordset.close
                                      webdb.close
                     
                   
				   %>
                &nbsp;
        </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Day&nbsp;&nbsp; <select size="1" name="cboDay" class="small">

                <%
                   dim Period
                   dim selectedP
                   
                   selectedP = false
                   
                   Period = request.form("cboDay")
                   
                   Set webdb = Server.CreateObject("ADODB.Connection")
                       webdb.Open Session("ConnectStr")
                   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
                   Set webdbCommand = Server.CreateObject("ADODB.Command")
                   
					  ssql = "Exec sp_Wls_selLeaveIDDesc '', '', '', '','','','PERIOD'"
						
				      Set webdbCommand.ActiveConnection = webdb
				          webdbCommand.CommandText = ssql
				          webdbRecordset.Open webdbCommand,,1,3
          
					   i = 1
					   count = 1
	  					
				       Do Until webdbRecordset.EOF
				          If Period = "" and count = 1 Then
				 	          response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("referenceid")) + "'>" + cstr(webdbRecordset.Fields("description")) + "</OPTION> "
							Elseif Period = webdbRecordset.Fields("referenceid") then
					          response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("referenceid")) + "'>" + cstr(webdbRecordset.Fields("description")) + "</OPTION> "
					          selectedP = true
						 	else   
						       response.write "<OPTION value='" + cstr(webdbRecordset.Fields("referenceid")) + "'>" + cstr(webdbRecordset.Fields("description")) + "</OPTION> "
							End If
							i = i + 1
							count = count + 1
							webdbRecordset.MoveNext
					   loop
				       webdbRecordset.close
				       webdb.close
				   %>
        
        </select>&nbsp;&nbsp;</font> </p>
        
        <p><font class="small">Date Apply For(DDMMYYYY)&nbsp; <!--webbot bot="Validation" B-Value-Required="TRUE" I-Minimum-Length="8" I-Maximum-Length="10" --> <input type="text" name="txtDate1" size="20" class="small" maxlength="10">&nbsp;&nbsp;&nbsp;
        to&nbsp;&nbsp; <!--webbot bot="Validation" B-Value-Required="TRUE" I-Minimum-Length="8" I-Maximum-Length="10" --> <input type="text" name="txtDate2" size="20" class="small" maxlength="10"></font> </p>
        <p><font class="small"><I>( The maximum characters You can Enter is 255 Characters )</I></font></p>
        <p><font class="small">Reason&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <!--<input type="text" name="txtReason" size="61" maxlength="255" class="small">&nbsp;-->
        <textarea rows="5" class ="small" name="txtReason" cols="50"></textarea>&nbsp;&nbsp;&nbsp;
        <!--<input type="text" name="txtReason" size="61" maxlength="255" class="small">-->
        <input type="button" value="Submit" name="cmdSubmit" onClick="Verify()" onmouseover="this.style.cursor='hand';" class="small"></font> </p>
        <input type="hidden" name="txtLockDate" class="small"
        <%				
					response.write " value='" & tempLockDate & "'"
					%>></font> </p>
		     		 
        <p></p>
      </form>
      <p></td>
  </tr>
  <tr>
    <td width="100%" height="28" colspan="2"></td>
  </tr>
  <tr>
    <td width="100%" height="28" colspan="2" align="left">
    <a href="leavebalance2.asp" onclick="NewWindow(this.href,'LeaveBalance','700','480','yes','center');return false" onfocus="this.blur()"><font class="marineblue"><u>Leave Balance</u></font></a>    
    &nbsp;
    <a href="leaveapp2.asp" onclick="NewWindow(this.href,'LeaveBalance','600','350','yes','center');return false" onfocus="this.blur()"><font class="marineblue"><u>Leave Application</u></font></a>    
    </td>
  </tr>
  <tr>
    <td width="100%" height="28" colspan="2"></td>
  </tr>
</table>

<table border="0" width="96%">
  <tr>
    <td width="100%" align="center"><img border="0" src="/eHres3/Image/dottedlinenav.gif" WIDTH="408" HEIGHT="4"></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="10" class="small"><br>
      &nbsp;<br>
      &nbsp;<br>Copyright © 1997-2000 Software
      Factory Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>

</body>

</html>
