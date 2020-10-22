<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<%
Response.Buffer = true
'Response.CacheControl = "private"

Dim ssql
Dim fromdate
Dim todate
Dim reason     
Dim status
Dim ApplyAdvance
Dim ApplyOthers

%>
<%		  
	  ssql = ""   
	if request("txtReason") = "" and request("txtDate1") <> "" and request("txtDate2") <> "" then		
       ssql = "Exec sp_Wls_ApplyLeave """ + Session("Regisno") + """,""" + Session("EmpID") + """, """ _
			  + request("cboLeaveType") + """,""" + request("txtDate1") + """, """ _
			  + request("txtDate2") + """, """ + request("cboDay") + """, '', " + request("txtOthers") + ", " _
			  + "0" + ", 'ADD'"
    
    'Response.Write SSQL
'Response.End
		
	elseif request("txtReason") <> "" and request("txtDate1") <> "" and request("txtDate2") <> "" then
       ssql = "Exec sp_Wls_ApplyLeave """ + Session("Regisno") + """,""" + Session("EmpID") + """, """ _
			  + request("cboLeaveType") + """,""" + request("txtDate1") + """, """ _
			  + request("txtDate2") + """, """ + request("cboDay") + """, """ _
			  + request("txtReason") + """, " + request("txtOthers") + ", " + "0" + ", 'ADD'"			 
	'Response.Write ssql

    end if	
	
  	
    if ssql <> "" then
	   Set webdb = Server.CreateObject("ADODB.Connection")
	   webdb.Open Session("ConnectStr")
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")

	   Set webdbCommand = Server.CreateObject("ADODB.Command")

	   Set webdbCommand.ActiveConnection = webdb
		   webdbCommand.CommandText = ssql		   
		   webdbRecordset.Open webdbCommand,,1 , 3
		   
      do until webdbRecordset.EOF 
        status = webdbrecordset.Fields(0)          
         webdbRecordset.MoveNext
       loop                                                 '// marked by ang on  25/7/2005                         
	
   
           webdbRecordset.close      
          webdb.Close
       
         Set webdb = nothing
          Set webdbRecordset = nothing
         Set webdbCommand = nothing                             '// marked by ang on 25/7/2005


 '        Response.Write(status)
'Response.end
		if status = "1" then %>
		  <script language="javascript">
		  <!--
			alert('Transaction Fail. Total Day Out Of Range.')
			window.history.go(-1)
		  //-->
		  </script>

<%		elseif status = "7" then %>
		  <script language="javascript">
		  <!--
			alert('Total Leave Applied Exceed Leave Balance!')
			window.history.go(-1)
		  //-->
		  </script>

		  
<%		elseif status = "2" then %>
		  <script language="javascript">
		  <!--
			alert('Transaction Fail. TMSID Not Found!')
			window.history.go(-1)
		  //-->
		  </script>
<%		elseif status = "4" then %>
		  <script language="javascript">
		  <!--
			alert('Transaction not fully success, some of the date is not allow to apply leave!')
			window.history.go(-1)
		  //-->
		  </script>

<%		elseif status = "6" then %>
		  <script language="javascript">
		 <!--
			alert('Transaction fail. Apply future date is not allow for this leave type!')
			window.history.go(-1)
		  //-->
		  </script>
<%		elseif status = "5" then %>
		  <script language="javascript">
		  <!--
			alert('Duplicate date found!')
			window.history.go(-1)
		  //-->
		  </script>

<%		elseif status = "99" then %>
		 <script language="javascript">		  
			alert('Incomplete Transaction!')
			window.history.go(-1)
		   </script>  
		  			
<%     
       end if       

    else %>
		  <script language="javascript">
		  <!--
			window.history.go(-1)
		  //-->
		  </script>    
<%  end if 
%>


<html>
<link rel="stylesheet" type="text/css" HREF="../css/login.css">
<title>Redirect to Main Page</title>
<script langauage="JavaScript">
function Redirect()
{
	location.href= "/ehres/main.asp"
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
        <p align="center"><font class="bigmarineblue">Leave Apply
        Successfully</font></td>
    </tr>
  </table>
  </center>
</div>


</body>
</html>



