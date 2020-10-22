<!-- #INCLUDE FILE = "../global/ConnectStr.asp"-->

<%
Response.Buffer = true
Dim ssql
Dim i
Dim colour
Dim count
Dim rowno
Dim ApproveRow
Dim maxrow
Dim rowcount
Dim POH
Dim AOT
Dim Upd
%>

<html>


<head>
<link rel="stylesheet" type="text/css" HREF="../css/login.css">

<title>Overtime Approval System</title>

<!--Place this script anywhere in a page.-->
<!--NOTE: You do not need to modify this script.-->

	<SCRIPT LANGUAGE="JavaScript">

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
		else
		{
			document.forms[0].submit();
		}
		

		m = CheckDate('txtDate2');
		if (!m)
		{
			window.alert("Invalid Date Entry. Please make sure the date format is [ddmmyyyy] and date is valid.");
			document.forms[0].reset()
		}
		else
		{
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
                        o = true;
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

</head>

<body topmargin="0" leftmargin="0" bgColor="#ffffff"> <!--background="../Image/bg-schome.jpg"-->

<table cellSpacing="0" cellPadding="0" border="0" width="100%">
  <tbody>
  <tr>
    <td vAlign="top" align="center" colspan="2" width="936" bgcolor="#0099CC" height="29">
      <div align="center">
        <center>
      <table border="0" width="100%">
        <!--<tr>
          <td width="3%">
          </td>-->
          <td width="23%">
		<font class="marinewhite">Employee ID : 
			<%    
				response.write session("EmpID")
			%>  
       
		</font>
        </td>
        <td width="74%"><font class="marinewhite"> Name :
       <%    
          response.write session("EmpName")
       %>  
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Date :
      <SCRIPT LANGUAGE = "JavaScript">
	

	// Array of day names
	var dayNames = new Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday");

	var monthNames = new Array("January","February","March","April","May","June","July",
	                           "August","September","October","November","December");

	var dt = new Date();
	var y  = dt.getYear();

	// Y2K compliant
	if (y < 1000) y +=1900;

    document.writeln;
	document.write(dayNames[dt.getDay()] + ", " + monthNames[dt.getMonth()] + " " + dt.getDate() + ", " + y);

	// 
	    </SCRIPT>
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
    <td vAlign="top" align="center" width="907" height="109"><img alt="Main Menu" src="../Image/appTmsOT1.gif" border="0" width="703" height="87"><br>
      &nbsp;</td></tr>
  <tr>
    <td width="100%" align="right"></td>
  </tr>
  
</table>
  <!--<tr>
    <td width="164%" height="44" colspan="2">-->
      <table border="0" width="100%" height="1">
        <tr>
          <td width="100%" height="1">
			<form method="POST" action="tmsapproval.asp" name="frmOTApproval" >
              <p>&nbsp;<font class="small">&nbsp;&nbsp;&nbsp;&nbsp; Status</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <select size="1" name="cboStatus" onchange="if(options[selectedIndex].value) top.location.href=('tmsapproval.asp?cboStatus=' + options[selectedIndex].value)" style="font-size: 8pt">
              
<script language="vbscript">
<!--
function UpdateOT() 
	
	dim rowcount 
	dim ssql
	dim maxrow
	dim approve
	
	maxrow = document.frmOTApproval.txtRowNo.value
		
	do until rowcount = cint(maxrow)
	rowcount = rowcount + 1
	
	ssql="if " + "document.frmOTApproval.A" + cstr(rowcount) + ".checked = false then" + chr(10) 
	ssql= ssql + " document.frmOTApproval.N" + cstr(rowcount) + ".value=0" + chr(10) 
	ssql=ssql + "end if"
		
	execute ssql
	
	ssql=""
	ssql="if " + "document.frmOTApproval.R" + cstr(rowcount) + ".checked = false then" + chr(10) 
	ssql= ssql + " document.frmOTApproval.D" + cstr(rowcount) + ".value=0" + chr(10) 
	ssql=ssql + "end if"
		
	execute ssql
	
	loop

	frmOTApproval.txtAction.value="UPD"
	document.frmOTApproval.submit()
End function	

function UncheckA(vRow)
	ssql= " document.frmOTApproval.A" + vRow + ".checked=false"
	execute ssql
end function

function UncheckR(vRow)
	ssql= " document.frmOTApproval.R" + vRow + ".checked=false"
	execute ssql
end function

function checkAllApprove()
   document.frmOTApproval.txtAppAll.value="CHECK"
   document.frmOTApproval.submit()
end function

function checkAllReject()
   document.frmOTApproval.txtRejectAll.value="CHECK"
   document.frmOTApproval.submit()
end function

function uncheckAll()
	document.All("txtAction").value="UNCHECK"
	document.frmOTApproval.submit()
end function

function UncheckPOH()
	ssql= " document.frmOTApproval.chkPOH.checked=false"
	execute ssql
	
	document.frmOTApproval.txtAOH1.value="AOH"
	document.frmOTApproval.txtPOH.value=""
end function

function UncheckAOH()
	ssql= " document.frmOTApproval.chkAOH.checked=false"
	execute ssql

	document.frmOTApproval.txtPOH1.value="POH"
end function


// -->
</script>
				 
              <%  
					dim vStatus
					
					   if request("cbostatus") <> "" then
					      vStatus = request("cbostatus")
					   elseIf request.form("cboStatus") = "" Then
					      vStatus = "PREOTPENDING"
					   Else
					      vStatus = request.form("cbostatus")
					   End If
					   
					   if vStatus  = "PREOTAPPROVED" then
			             response.write "<Option value = 'PREOTPENDING'> Pre OT Approval (Pending) </Option>"
			             response.write "<Option selected value = 'PREOTAPPROVED'> Pre OT Approval </Option>"
			             response.write "<Option value = 'OTDONE'> OT Done (Discrepancy) </Option>"
			             response.write "<Option value = 'NOPREOTDONE'> OT Done (No Pre OT Approval) </Option>"
			             response.write "<Option value = 'APPROVEDOT'> Approved OT </Option>"
					   elseif vStatus  = "OTDONE" then
			             response.write "<Option value = 'PREOTPENDING'> Pre OT Approval (Pending) </Option>"
			             response.write "<Option value = 'PREOTAPPROVED'> Pre OT Approval </Option>"
			             response.write "<Option selected value = 'OTDONE'> OT Done (Discrepancy) </Option>"
			             response.write "<Option value = 'NOPREOTDONE'> OT Done (No Pre OT Approval) </Option>"
			             response.write "<Option value = 'APPROVEDOT'> Approved OT </Option>"
					   elseif vStatus  = "NOPREOTDONE" then
			             response.write "<Option value = 'PREOTPENDING'> Pre OT Approval (Pending) </Option>"
			             response.write "<Option value = 'PREOTAPPROVED'> Pre OT Approval </Option>"
			             response.write "<Option value = 'OTDONE'> OT Done (Discrepancy) </Option>"
			             response.write "<Option selected value = 'NOPREOTDONE'> OT Done (No Pre OT Approval) </Option>"
			             response.write "<Option value = 'APPROVEDOT'> Approved OT </Option>"
					   elseif vStatus  = "APPROVEDOT" then
			             response.write "<Option value = 'PREOTPENDING'> Pre OT Approval (Pending) </Option>"
			             response.write "<Option value = 'PREOTAPPROVED'> Pre OT Approval </Option>"
			             response.write "<Option value = 'OTDONE'> OT Done (Discrepancy) </Option>"
			             response.write "<Option value = 'NOPREOTDONE'> OT Done (No Pre OT Approval) </Option>"
			             response.write "<Option selected value = 'APPROVEDOT'> Approved OT </Option>"			             
					   else
			             response.write "<Option Selected value = 'PREOTPENDING'> Pre OT Approval (Pending) </Option>"
			             response.write "<Option value = 'PREOTAPPROVED'> Pre OT Approval </Option>"
			             response.write "<Option value = 'OTDONE'> OT Done (Discrepancy) </Option>"
			             response.write "<Option value = 'NOPREOTDONE'> OT Done (No Pre OT Approval) </Option>"
			             response.write "<Option value = 'APPROVEDOT'> Approved OT </Option>"
			          end if
				%>
              </select>

			  <% if vStatus <> "PREOTPENDING" and vStatus <> "NOPREOTDONE" then %>                            
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;			  
                  <font class="small">&nbsp;&nbsp;
                  Date <input type="text" name="txtDate1" size="9" class="small" 
                  <% 
	                 if request.form("txtDate1") <> "" then
	                    response.write " value='" & request.form("txtDate1") & "'"              
	                 end if   
	              %>   
              
                  >&nbsp;&nbsp;&nbsp; to&nbsp;&nbsp; <input type="text" name="txtDate2" size="9" class="small" 
                  <% 
	                 if request.form("txtDate2") <> "" then
	                    response.write " value='" & request.form("txtDate2") & "'"              
	                 end if   
	              %>   
              
                  ></font></p>
              
                  <p>&nbsp;&nbsp;&nbsp;&nbsp; <font class="small"> Search By </font><b><select size="1" name="cboSearchBy" class="small">

                  <%  
			     		dim SearchBy
					
				  	   SearchBy = request.form("cboSearchBy")
					
					   if SearchBy  = "DEPT" then
			             response.write "<Option value = ''>  </Option>"
			             response.write "<Option value = 'EMPID'> Employee </Option>"
			             response.write "<Option value = 'EMPGROUP'> Employee Group</Option>"			             
			             response.write "<Option Selected value = 'DEPT'> Department </Option>"
			          elseif SearchBy = "EMPID" Then
			             response.write "<Option value = ''>  </Option>"			          
			             response.write "<Option Selected value = 'EMPID'> Employee </Option>"
			             response.write "<Option value = 'EMPGROUP'> Employee Group </Option>"			             
			             response.write "<Option value = 'DEPT'> Department </Option>"
			          elseif SearchBy = "EMPGROUP" Then
			             response.write "<Option value = ''>  </Option>"			          
			             response.write "<Option value = 'EMPID'> Employee </Option>"
			             response.write "<Option selected value = 'EMPGROUP'> Employee Group </Option>"			             
			             response.write "<Option value = 'DEPT'> Department </Option>"			             
			          else
			             response.write "<Option Selected value = ''>  </Option>"			          
			             response.write "<Option value = 'EMPID'> Employee </Option>"
			             response.write "<Option value = 'EMPGROUP'> Employee Group </Option>"
			             response.write "<Option value = 'DEPT'> Department </Option>"
			          end if
			     	%>
              
                 </select></b>
              
	               <font class="small">&nbsp;&nbsp;&nbsp;&nbsp; Employee ID/
                   Employee Group ID / Department ID </font><b>&nbsp;<input type="text" name="txtID" size="8" class="small" 
                   <% 
   		              if request.form("txtID") <> "" then
   		                 response.write " value='" & request.form("txtID") & "'"
   		              end if
		           %>   
              
	               >&nbsp;&nbsp;<input type="button" value="Search" name="cmdSearch" onClick="Verify()" onmouseover="this.style.cursor='hand';" class="small"></b></p>             
              <!--<br>&nbsp;</br>
              <!--<p>&nbsp;</p>-->              
            <% else %>
	            <!--<p>&nbsp;</p>-->
	             <br>&nbsp;</br>
			  <% end if %>
              
              <table border="0" width="100%" cellSpacing="0" cellPadding="0">
				 <%if vStatus = "PREOTPENDING" then%>
		            <tr>
	                  <!--<td width="4%"></td>-->	            
					    <td align='center' width="7%" bgcolor="#F3F3F3"><font class="marineblack"><b>Approve</b></font></td>
					    <td align='center' width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reject</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee ID</b></font></td>
					    <td width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee Name</b></font></td>
		              <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Department ID</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Applied</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>POH</b></font></td>
					    <td width="8%" bgcolor="#F3F3F3"><font class="marineblack"><b>Amount (RM)</b></font></td>
		            </tr>
				 <%elseif vStatus = "PREOTAPPROVED" then%>
		            <tr>
	                  <!--<td width="4%">&nbsp;</td>-->
					    <td align='center' width="7%" bgcolor="#F3F3F3"><font class="marineblack"><b>Change</b></font></td>
					    <td align='center' width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reject</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee ID</b></font></td>
					    <td width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee Name</b></font></td>
		              <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Department ID</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Applied</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>POH</b></font></td>
					    <td width="8%" bgcolor="#F3F3F3"><font class="marineblack"><b>Amount (RM)</b></font></td>
		            </tr>
				 <%elseif vStatus = "OTDONE" then%>
   		         <tr>
	                 <!--<td width="4%">&nbsp;</td>-->
					    <td align='center' width="7%" bgcolor="#F3F3F3"><font class="marineblack"><b>Approve</b></font></td>
					    <td align='center' width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reject</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee ID</b></font></td>
					    <td width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee Name</b></font></td>
		              <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Department ID</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Applied</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>POH</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>AOH</b></font></td>
					    <td width="8%" bgcolor="#F3F3F3"><font class="marineblack"><b>Amount (RM)</b></font></td>
		            </tr>
				 <%elseif vStatus = "NOPREOTDONE" then%>
   		         <tr>
	                 <!--<td width="4%">&nbsp;</td>-->
					    <td align='center' width="7%" bgcolor="#F3F3F3"><font class="marineblack"><b>Approve</b></font></td>
					    <td align='center' width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reject</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee ID</b></font></td>
					    <td width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee Name</b></font></td>
		              <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Department ID</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Applied</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>POH</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>AOH</b></font></td>
					    <td width="8%" bgcolor="#F3F3F3"><font class="marineblack"><b>Amount (RM)</b></font></td>
		            </tr>
				 <%elseif vStatus = "APPROVEDOT" then%>
   		         <tr>
	                 <!--<td width="4%">&nbsp;</td>-->
					    <td align='center' width="7%" bgcolor="#F3F3F3"><font class="marineblack"><b>Change</b></font></td>
					    <td align='center' width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reject</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee ID</b></font></td>
					    <td width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Employee Name</b></font></td>
		              <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Department ID</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Applied</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>POH</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>AOH</b></font></td>
					    <td width="8%" bgcolor="#F3F3F3"><font class="marineblack"><b>Amount (RM)</b></font></td>
		            </tr>
				 <%end if%>
				 
            <%
               dim vID
               dim vStatusID
               
               if vStatus = "PREOTAPPROVED" or vStatus = "OTDONE" or vStatus = "APPROVEDOT" then
                  vStatusID = "Y"
               else
                  vStatusID = "P"
               end if
               
               if SearchBy = "" then
               	 vID = "%"
               else
               	 vID = request("txtID") + "%"
               end if

            		 Set webdb = Server.CreateObject("ADODB.Connection")
   		         webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
  					 Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		             Set webdbCommand = Server.CreateObject("ADODB.Command")

				if request("txtDate1") = "" and request("txtDate2") = "" then
					 ssql = "Exec sp_Wtms_OTApproval '" + Session("Regisno") + "', '" + vID + "','','01/01/1900','01/01/1900', '" _
					 		  + Session("EmpID") + "', 'ENG', '" + vStatusID + "', '" + vStatus +"', '" + SearchBy + "'"
              'Response.Write ssql
				elseif request("txtDate1") <> "" and request("txtDate2") <> "" then
					 ssql = "Exec sp_Wtms_OTApproval '" + Session("Regisno") + "', '" + vID + "', '', '" _
					 		  + request("txtDate1") + "', '" + request("txtDate2") + "', '" _
					 		  + Session("EmpID") + "', 'ENG', '" + vStatusID + "', '" + vStatus +"', '" + SearchBy + "'"
				elseif request("txtDate1") <> "" and request("txtDate2") = "" then
					 ssql = "Exec sp_Wtms_OTApproval '" + Session("Regisno") + "', '" + vID + "', '', '" _
					 		  + request("txtDate1") + "', '01/01/1900', '" + Session("EmpID") + "', 'ENG', '" _
					 		  + vStatusID + "', '" + vStatus +"', '" + SearchBy + "'"
				elseif request("txtDate1") = "" and request("txtDate2") <> "" then
					 ssql = "Exec sp_Wtms_OTApproval '" + Session("Regisno") + "', '" + vID + "', '', '01/01/1900', '" _
					 		  + request("txtDate2") + "', '" + Session("EmpID") + "', 'ENG', '" + vStatusID + "', '" + vStatus +"', '" + SearchBy + "'"					 		  
				end if
				'Response.Write ssql
			        Set webdbCommand.ActiveConnection = webdb
			            webdbCommand.CommandText = ssql
			            webdbRecordset.Open webdbCommand,,1 , 3

					 colour = 0
					' paste here
			 if NOT webdbRecordset.EOF or webdbRecordset.BOF then  
			   If vStatus = "PREOTPENDING" or vStatus = "PREOTAPPROVED" Then
			        
			           Do Until webdbRecordset.EOF
					       rowno = rowno + 1			        
			           
				           if count = 1 then
				              colour = " bgcolor='#eeeeee'"
 				           else
				              colour = ""
				           end if

					 		 if Request.Form("txtAppAll")= "CHECK" then
							    AstrCheck="checked"
							 else
							    AstrCHECK= ""     
							 end if


					 		 if Request.Form("txtRejectAll")= "CHECK" then
							    RstrCheck="checked"
							 else
							    RstrCHECK= ""     
							 end if
							
						   temp	= webdbRecordset.Fields("empid")
				           response.write "<tr>"
		                  'response.write "<td>&nbsp;</td>"
						   response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' onclick='UncheckR(""" + cstr(rowno) + """)' name=A" + cstr(rowno) + " value='ON' " + Astrcheck + "></font></td>"
						   response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' onclick='UncheckA(""" + cstr(rowno) + """)' name=R" + cstr(rowno) + " value='ON' " + Rstrcheck + "></font></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("empid") + "<input type='hidden' name=E" + cstr(rowno) + " value= " + webdbRecordset.Fields("empid") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("empname") + "<input type='hidden' name=N" + cstr(rowno) + " value= " + webdbRecordset.Fields("empname") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("deptid") + "<input type='hidden' name=D" + cstr(rowno) + " value= " + webdbRecordset.Fields("deptid") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("dateapplied") + "<input type='hidden' name=DA" + cstr(rowno) + " value= " + webdbRecordset.Fields("dateapplied") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("poh") + "<input type='hidden' name=PO" + cstr(rowno) + " value= " + webdbRecordset.Fields("poh") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("amount") + "<input type='hidden' name=RM" + cstr(rowno) + " value= " + webdbRecordset.Fields("amount") + "></td></tr>"
				           webdbRecordset.MoveNext  
				           count = abs(count - 1)
			           loop
			           response.write "<input type=hidden name=txtRowNo value=" + cstr(rowno) + ">"
		  				 
			           webdbRecordset.close
			           webdb.close

					 ElseIf vStatus = "OTDONE" or vStatus = "NOPREOTDONE" or vStatus = "APPROVEDOT" Then
			           Do Until webdbRecordset.EOF
					       rowno = rowno + 1			        
			           
				           if count = 1 then
				              colour = " bgcolor='#eeeeee'"
 				           else
				              colour = ""
				           end if

					 		 if Request.Form("txtAppAll")= "CHECK" then
							    AstrCheck="checked"
							 else
							    AstrCHECK= ""     
							 end if


					 		 if Request.Form("txtRejectAll")= "CHECK" then
							    RstrCheck="checked"
							 else
							    RstrCHECK= ""     
							 end if
                           
                           temp = webdbRecordset.Fields("empid") 
				           response.write "<tr>"
		                  'response.write "<td>&nbsp;</td>"
						    response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' onclick='UncheckR(""" + cstr(rowno) + """)' name=A" + cstr(rowno) + " value='ON' " + Astrcheck + "></font></td>"
						    response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' onclick='UncheckA(""" + cstr(rowno) + """)' name=R" + cstr(rowno) + " value='ON' " + Rstrcheck + "></font></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("empid") + "<input type='hidden' name=E" + cstr(rowno) + " value= " + webdbRecordset.Fields("empid") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("empname") + "<input type='hidden' name=N" + cstr(rowno) + " value= " + webdbRecordset.Fields("empname") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("deptid") + "<input type='hidden' name=D" + cstr(rowno) + " value= " + webdbRecordset.Fields("deptid") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("dateapplied") + "<input type='hidden' name=DA" + cstr(rowno) + " value= " + webdbRecordset.Fields("dateapplied") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("poh") + "<input type='hidden' name=PO" + cstr(rowno) + " value= " + webdbRecordset.Fields("poh") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("aot") + "<input type='hidden' name=AO" + cstr(rowno) + " value= " + webdbRecordset.Fields("aot") + "</td>"				           
				           response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields(8) + "<input type='hidden' name=RM" + cstr(rowno) + " value= " + webdbRecordset.Fields(8) + "></td></tr>"
				           webdbRecordset.MoveNext  
				           count = abs(count - 1)
			           loop
		  				 response.write "<input type=hidden name=txtRowNo value=" + cstr(rowno) + ">"
			           
			           webdbRecordset.close
			           webdb.close
					 End If
			end if
			   if Request.form("txtAction")="UPD" then	
				   Set webdb = Server.CreateObject("ADODB.Connection")
				   		webdb.Open Session("ConnectStr")
				   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
				   Set webdbCommand = Server.CreateObject("ADODB.Command")
				   Set webdbCommand.ActiveConnection = webdb

				   if request("txtPOH") <> "" and request.form("txtPOH1") = "POH" then
				      POH = request("txtPOH") 
				   end if

				   maxrow = request.form("txtRowNo")
				   approve ="false"

				     do until rowcount = cint(maxrow)
				     	 rowcount = rowcount + 1

'				     	 if POH = "" then
'				     	    POH = Request("PO" + cstr(rowcount))
'				     	 end if

				     	 if (vStatus = "OTDONE" or vStatus = "NOPREOTDONE" or vStatus = "APPROVEDOT") And request.form("txtAOH1") = "AOH" then
				     	    POH = Request("AO" + cstr(rowcount))
				     	 end if

				     	 if vStatus = "OTDONE" or vStatus = "NOPREOTDONE" then
				     	    AOH = Request("AO" + cstr(rowcount))
				     	 else
				     	    AOH = "00:00"
				     	 end if

				        if Request.Form("N" + cstr(rowcount)) <> "0" and (vStatus = "PREOTPENDING" or vStatus = "NOPREOTDONE" ) then
						   approve = "true"

						   ssql = "Exec sp_Wtms_updOTApproval '" + Session("Regisno") + "', '" _
						           + Request("E" + cstr(rowcount)) + "', '" + Request("DA" + cstr(rowcount)) + "', '', '" _
						           + POH + "', '" + AOH + "', '" + Session("EmpID") + "', 'ENG', 'Y'"
						   webdbCommand.CommandText = ssql
						   webdb.Execute webdbCommand.CommandText
				        end if


				        if Request.Form("N" + cstr(rowcount)) <> "0" and vStatus = "PREOTAPPROVED" then
						   approve = "true"

						   ssql = "Exec sp_Wtms_updOTApproval '" + Session("Regisno") + "', '" _
						           + Request("E" + cstr(rowcount)) + "', '" + Request("DA" + cstr(rowcount)) + "', '', '" _
						           + POH + "', '" + AOH + "', '" + Session("EmpID") + "', 'ENG', 'C'"
						   webdbCommand.CommandText = ssql
						   webdb.Execute webdbCommand.CommandText
				        end if

				        if Request.Form("N" + cstr(rowcount)) <> "0" and vStatus = "OTDONE" then
						   approve = "true"

						   ssql = "Exec sp_Wtms_updOTApproval '" + Session("Regisno") + "', '" _
						           + Request("E" + cstr(rowcount)) + "', '" + Request("DA" + cstr(rowcount)) + "', '', '" _
						           + POH + "', '" + AOH + "', '" + Session("EmpID") + "', 'ENG', 'C2'"
						   webdbCommand.CommandText = ssql
						   webdb.Execute webdbCommand.CommandText

				        end if
				        
				        if Request.Form("N" + cstr(rowcount)) <> "0" and vStatus = "APPROVEDOT" then
						   approve = "true"

						   ssql = "Exec sp_Wtms_updOTApproval '" + Session("Regisno") + "', '" _
						           + Request("E" + cstr(rowcount)) + "', '" + Request("DA" + cstr(rowcount)) + "', '', '" _
						           + POH + "', '" + AOH + "', '" + Session("EmpID") + "', 'ENG', 'C'"
						   webdbCommand.CommandText = ssql
						   webdb.Execute webdbCommand.CommandText
				        end if

				        if Request.Form("D" + cstr(rowcount)) <> "0" then
						   approve = "true"
						   ssql = "Exec sp_Wtms_updOTApproval '" + Session("Regisno") + "', '" _
						   		    + Request.form("E" + cstr(rowcount)) + "', '" + Request.form("DA" + cstr(rowcount)) + "', '', '" _
				                  + POH + "', '" + AOH + "', '" + Session("EmpID") + "', 'ENG', 'N'"
						   webdbCommand.CommandText = ssql
						   webdb.Execute webdbCommand.CommandText
				        end if			      
			         loop
				  
		          if approve = "true" then
		             response.redirect "tmsapproval.asp"
		          end if
		    end if	
		%>
							 
                <tr>
                  <td width="6%">&nbsp;</td>
                  <td width="50%" colspan="5">&nbsp;</td>
                  <td width="44%" colspan="4">&nbsp;</td>
                </tr>
                <tr>
                  <!--<td width="6%"></td>-->
                  <td width="50%" colspan="5"><font class="small">&nbsp;&nbsp; 
                    <input type="checkbox" name="chkPOH" onclick="UnCheckAOH()" value="ON" class="small">
                    POH&nbsp;</font>
                    <input type="text" name="txtPOH" size="11" class="small">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    
                    <% if vStatus <> "PREOTPENDING" and vStatus <> "PREOTAPPROVED" then %>
                       <input type="checkbox" name="chkAOH" onclick="UnCheckPOH()" value="ON" class="small">
                       <font class="small">POH = AOH</font></td>
                    <% end if %>
                    
                  <td width="44%" colspan="4"></td>
                </tr>
                <tr>
                  <td width="6%">&nbsp;</td>
                  <td width="50%" colspan="5">&nbsp;</td>
                  <td width="44%" colspan="4">&nbsp;</td>
                </tr>
                <tr>
                  <td width="6%">&nbsp;</td>
                  <td width="50%" colspan="5">
        
        		  <% if vStatus = "PREOTPENDING" or vStatus = "OTDONE" or vStatus = "NOPREOTDONE" then%>
		            <input type="button" value="Select All Approve" name="cmdSelectAllA" onclick="checkAllApprove()" class="small">
		         <% end if %>
		         
		         <% if vStatus = "PREOTAPPROVED" or vStatus = "APPROVEDOT" then %>
		            <input type="button" value="Select All Change" name="cmdSelectAllC" onclick="checkAllApprove()" class="small">
		         <% end if %>

		         <input type="button" value="Select All Reject" name="cmdSelectAllR" onclick="checkAllReject()"  class="small">
		         <% if temp <> "" then %>
		         <input type="submit" value="Update" name="cmdUpdate"	onclick="UpdateOT()" class="small">
				 <% end if %>   	
		         <input type=hidden name=txtAction>		         
		         <input type=hidden name=txtAppAll>
		         <input type=hidden name=txtRejectAll>		         
		         <input type=hidden name=txtPOH1>
		         <input type=hidden name=txtAOH1>		         
        
                  </td>
                  <td width="44%" colspan="4">&nbsp;</td>
                </tr>
                <tr>
                  <td width="6%">&nbsp;</td>
                  <td width="50%" colspan="5"></td>
                  <td width="44%" colspan="4"></td>
                </tr>
              </table>
              
              <p>&nbsp;</p>
              
            </form>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%" height="21" colspan="2">

    </td>
  </tr>
</table>
<table border="0" width="96%">
  <tr>
    <td width="100%" align="center"><img border="0" src="../Image/dottedlinenav.gif"></td>
  </tr>
  <tr>
    <td width="100%" align="center">Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </td>
  </tr>
</table>
<p>&nbsp;</p>
</html>