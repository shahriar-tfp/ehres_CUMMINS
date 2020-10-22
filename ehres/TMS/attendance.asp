<!-- #INCLUDE FILE = "../global/ConnectStr.asp"-->
<%
Response.Buffer = true


Dim ssql
dim status
dim datein
dim dateout
dim timein
dim timeout1
dim regisno1
dim empid1
'regisno1 = Session("Regisno")
'empid1 = Session("EmpID")
'datein = Request.Form("txtdatein")
'dateout = Request.Form("txtdateout")
'timein= Request.Form("txttimein")
'timeout1= Request.Form("txttimeout")

  ' Set webdb = Server.CreateObject("ADODB.Connection")
  '     webdb.Open Session("ConnectStr")
   'Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  ' Set webdbCommand = Server.CreateObject("ADODB.Command")
   
  ' ssql ="Insert into etimesheet values("&regisno1&","&empid1&","&datein&","&dateout&","&timein&","&timeout1&")"

 '  Set webdbCommand.ActiveConnection = webdb
 '  webdbCommand.CommandText = ssql
  ' webdbRecordset.Open webdbCommand,,1 , 3
   
   'if Request.Form("txtlockdate") ="Pass" then
'		Response.Write "pass1"
'   end if	
  ' tempLockDate = webdbRecordset.fields("lockdate")
      


  
      ssql = "Exec sp_web_getsheetdata """ + Session("Regisno") + """,""" + Session("EmpID") + """, """ _
			 
			  + request("txtdatein") + """, """ + request("txtdateout") + """, """ _
			 + request("txttimein") + """, """ + request("txttimeout") + """,  'ADD'"			 

	   Set webdb = Server.CreateObject("ADODB.Connection")
	  	   webdb.Open Session("ConnectStr")
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	   Set webdbCommand = Server.CreateObject("ADODB.Command")

	   Set webdbCommand.ActiveConnection = webdb
		   webdbCommand.CommandText = ssql		   
	  webdbRecordset.Open webdbCommand,,1 , 3
	 

		 %>
		 
		 
<html>
<link rel="stylesheet" type="text/css" HREF="../css/login.css">
<title>Redirect to Main Page</title>
<script langauage="JavaScript">
function Redirect()
{
	location.href= "/eHRES/tms/i-attendance.asp"
}
function RedirectWithDelay()
{
	window.setTimeout("Redirect();", 1000);
}
</script>
<body bgcolor="#ffffff" onload="RedirectWithDelay();">

<div align="center">
  <center>
  <table border="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td width="100%">
        <p align="center"><font class="bigmarineblue">Insert Attendance
        Successfully</font></td>
    </tr>
  </table>
  </center>
</div>


</body>
</html>

		 
		 
		 