<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/AdoVbs.asp"-->
<!-- #include virtual ="/ehres/global/inputSession.asp"-->

<%session.Timeout=20%>
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
Dim page_size
Dim current_page
dim vStatus
Dim myconn
Dim rs
Dim sql
Dim page_count
Dim connect_string

connect_string =Session("ConnectStr")
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

<title>Leave Approval System</title>

<!--Place this script anywhere in a page.-->
<!--NOTE: You do not need to modify this script.-->

	<script LANGUAGE="JavaScript">

	function Verify()
	{
		msg = "";
		m = true;
		n = true;
        
        document.frmLeaveApproval.txtSearch.value ="Search";
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
	
	</script>

</head>

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
    </td></tr>
  <tr>
    <td vAlign="top" colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></td></tr>
  <tr>
    <td vAlign="top" align="center" width="907" height="109"><img alt="Main Menu" src="../Image/englsappr.gif" border="0" width="703" height="87"><br>
      &nbsp;</td></tr>

  <tr>
    <td width="100%" align="right"></td>
  </tr>
  
  
</table>

<table border="0" width="96%" cellspacing="0" height="121">
  <tr>
    <td width="164%" height="44" colspan="2">
      <table border="0" width="100%" height="1">
        <tr>
          <td width="100%" height="1">

<form method="POST" action="app_approval.asp" name="frmLeaveApproval">
              <p>&nbsp;<font class="small">&nbsp;&nbsp;&nbsp;&nbsp; Status</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <select size="1" onchange="if(options[selectedIndex].value) top.location.href=('app_approval.asp?cboStatus=' + options[selectedIndex].value)" name="cboStatus" style="font-size: 8pt">
              
<script language="vbscript">
<!--
function ApproveRejectLeave() 
	
	dim rowcount 
	dim ssql
	dim maxrow
   dim approve
	
	maxrow = document.frmLeaveApproval.txtRowNo.value
'   maxrow = approverow
		
	do until rowcount = cint(maxrow)
	rowcount = rowcount + 1
	
	ssql="if " + "document.frmLeaveApproval.A" + cstr(rowcount) + ".checked = false then" + chr(10) 
	ssql= ssql + " document.frmLeaveApproval.N" + cstr(rowcount) + ".value=0" + chr(10) 
	ssql=ssql + "end if"
		
	execute ssql
	
	ssql=""
	ssql="if " + "document.frmLeaveApproval.R" + cstr(rowcount) + ".checked = false then" + chr(10) 
	ssql= ssql + " document.frmLeaveApproval.D" + cstr(rowcount) + ".value=0" + chr(10) 
	ssql=ssql + "end if"
		
	execute ssql
	
	loop

	frmLeaveApproval.txtAction.value="UPD"
	document.frmLeaveApproval.submit()
End function	

function RejectLeave() 
	
	dim rowcount 
	dim ssql
	dim maxrow
   dim approve
	
	maxrow = document.frmLeaveApproval.txtRowNo.value
'   maxrow = approverow
		
	do until rowcount = cint(maxrow)
	rowcount = rowcount + 1
	
	ssql="if " + "document.frmLeaveApproval.R" + cstr(rowcount) + ".checked = false then" + chr(10) 
	ssql= ssql + " document.frmLeaveApproval.D" + cstr(rowcount) + ".value=0" + chr(10) 
	ssql=ssql + "end if"
		
	execute ssql
	loop

	frmLeaveApproval.txtAction.value="UPD"
	document.frmLeaveApproval.submit()
End function	

function UncheckA(vRow)

'	ssql= " document.All(""A" + vRow + """).checked=false"
	ssql= " document.frmLeaveApproval.A" + vRow + ".checked=false"
	execute ssql

end function

function UncheckR(vRow)
'	ssql= " document.All(""R" + vRow + """).checked=false"
	ssql= " document.frmLeaveApproval.R" + vRow + ".checked=false"
	execute ssql

end function

function checkAllApprove()
'	document.All("txtAction").value="CHECK"

	dim rowcount 
	dim ssql
	dim maxrow
    dim approve
	
	maxrow = document.frmLeaveApproval.txtRowNo.value
		
	do until rowcount = cint(maxrow)
  	   rowcount = rowcount + 1
  	   
  	   if document.frmLeaveApproval.cboStatus.value  = "R" or document.frmLeaveApproval.cboStatus.value  = "P" then
	      ssql="document.frmLeaveApproval.R" + cstr(rowcount) + ".checked = false "
	      execute ssql
       end if
       
  	   if document.frmLeaveApproval.cboStatus.value  = "A" or document.frmLeaveApproval.cboStatus.value  = "P" then
	      ssql="document.frmLeaveApproval.A" + cstr(rowcount) + ".checked = true "
	      execute ssql
	   end if   
	loop
	
'    document.frmLeaveApproval.txtAppAll.value="CHECK"
'	document.frmLeaveApproval.submit()
end function

function checkAllReject()
'	document.All("txtAction").value="CHECK"
'    document.frmLeaveApproval.txtRejectAll.value="CHECK"
'	document.frmLeaveApproval.submit()

	dim rowcount 
	dim ssql
	dim maxrow
    dim approve
	
	maxrow = document.frmLeaveApproval.txtRowNo.value
		
	do until rowcount = cint(maxrow)
  	   rowcount = rowcount + 1
  	   if document.frmLeaveApproval.cboStatus.value  = "A" or document.frmLeaveApproval.cboStatus.value  = "P" then
  	      ssql="document.frmLeaveApproval.R" + cstr(rowcount) + ".checked = true "
	      execute ssql
	   end if
	   
	   if document.frmLeaveApproval.cboStatus.value  = "R" or document.frmLeaveApproval.cboStatus.value  = "P" then
	      ssql="document.frmLeaveApproval.A" + cstr(rowcount) + ".checked = false "
	      execute ssql
	   end if   
	loop

end function

function uncheckAll()
	document.All("txtAction").value="UNCHECK"
	document.frmLeaveApproval.submit()
end function

// -->
</script>
				 
              <%  
				'dim vStatus1
					    tempstatus= session("ssStatus")
					    if tempStatus="" then 
					       tempStatus="P"
					    end if   
					    vStatus = tempstatus
					    If request("cboStatus") <> "" Then
					       vStatus = request("cboStatus")
					    elseIf request.form("cboStatus") ="" and tempstatus="P" Then 
						   vStatus ="P"   					    
					    Else
					      vStatus=tempstatus
					    End If
						'Response.Write(tempStatus)
						'Response.Write(vStatus)		
					   if vStatus  = "R" then
			             response.write "<Option value = 'A'> Approved </Option>"
			             response.write "<Option value = 'P'> Pending </Option>"
			             response.write "<Option Selected value = 'R'> Rejected </Option>"
			          elseif vStatus  = "A" then
			             response.write "<Option Selected value = 'A'> Approved </Option>"
			             response.write "<Option value = 'P'> Pending </Option>"
			             response.write "<Option value = 'R'> Rejected </Option>"
			          else
			             response.write "<Option value = 'A'> Approved </Option>"
			             response.write "<Option Selected value = 'P'> Pending </Option>"
			             response.write "<Option value = 'R'> Rejected </Option>"
			         end if
				%>
              </select>
              
              <% If vStatus = "A" or vStatus= "R" Then%>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="small">&nbsp;&nbsp;
              Date Apply For&nbsp;<input type="text" name="txtDate1" size="9" class="small" <% 
	              tempssdate1=session("ssdate1")
	             if request.form("txtDate1") <> "" or tempssdate1 ="01/01/1900" then
	                response.write " value='" & request.form("txtDate1") & "'"
	             else
	                Response.Write " value='" & session("ssDate1") & "'"               
	             end if   
	          %>>&nbsp;&nbsp;&nbsp;
              to&nbsp;&nbsp; <input type="text" name="txtDate2" size="9" class="small" <% 
	              tempssdate2=session("ssdate2")
	             if request.form("txtDate2") <> "" or tempssdate2 ="01/01/1900" then
	                response.write " value='" & request.form("txtDate2") & "'"
	             else
	                Response.Write " value='" & session("ssDate2") & "'"                     
	             end if   
	          %>></font></p>
	         
              <p>&nbsp;&nbsp;&nbsp;&nbsp; <font class="small"> Search By </font><b><select size="1" name="cboSearchBy" class="small">

              <%  
					dim SearchBy
					
					   'SearchBy = request.form("cboSearchBy") 'or session("ssSearchBy") 
					    SearchBy = session("ssSearchBy")
					       'if SearchBy = "" then 
					           'SearchBy = request.form("cboSearchBy")
					       'end if   
					   
					  if SearchBy  = "DEPT" then
			             response.write "<Option value = ''>  </Option>"
			             response.write "<Option value = 'EMPID'> Employee </Option>"
			             response.write "<Option Selected value = 'DEPT'> Department </Option>"
			          elseif SearchBy = "EMPID" Then
			             response.write "<Option value = ''>  </Option>"			          
			             response.write "<Option Selected value = 'EMPID'> Employee </Option>"
			             response.write "<Option value = 'DEPT'> Department </Option>"
			          else
			             response.write "<Option Selected value = ''>  </Option>"			          
			             response.write "<Option value = 'EMPID'> Employee </Option>"
			             response.write "<Option value = 'DEPT'> Department </Option>"
			          end if
				%>
              
              </select></b>
              
	               <font class="small">&nbsp;&nbsp;&nbsp;&nbsp; Employee ID / Department ID
              </font><b>&nbsp;<input type="text" name="txtID" size="8" class="small" <% 
   		              tempID = session("ssID")
   		              if request.form("txtID") <> "" or tempID ="%" then
   		                 response.write " value='" & request.form("txtID") & "'"
   		              else
   		                 response.write " value='" & tempID & "'" 
   		              end if
		           %>>&nbsp;&nbsp;<input type="button" value="Search" name="cmdSearch" onClick="Verify()" onmouseover="this.style.cursor='hand';" class="small" ><input type=hidden name="txtSearch" size=8></b></p>
	               
              <% 
              else
                 Response.Write "<p><b></b></p>"
              end if%>
              
              <table cellSpacing="0" cellPadding="0" border="0" width="100%">
				 <%if vStatus = "P" then%>
		             <tr>
					    <td align="center" width="3%" bgcolor="#F3F3F3"><font class="marineblack"><b>Approve</b></font></td>
					    <td align="center" width="5%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reject</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Apply On</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Apply For</b></font></td>
                        <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>EmpID</b></font></td>
					    <td width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Name</b></font></td>
					    <td width="9%" bgcolor="#F3F3F3"><font class="marineblack"><b>Department</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>Leave ID</b></font></td>
					    <td width="7%" bgcolor="#F3F3F3"><font class="marineblack"><b>Day</b></font></td>
					    <td width="18%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reason</b></font></td>
		            </tr>
				 <%elseif vStatus = "A" then%>
		            <tr>
					    <td align="center" width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reject</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Apply On</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Apply For</b></font></td>
		                <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>EmpID</b></font></td>
					    <td width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Name</b></font></td>
					    <td width="9%" bgcolor="#F3F3F3"><font class="marineblack"><b>Department</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>Leave ID</b></font></td>
					    <td width="7%" bgcolor="#F3F3F3"><font class="marineblack"><b>Day</b></font></td>
					    <td width="18%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reason</b></font></td>
		            </tr>
				 <%elseif vStatus = "R" then%>
   		         <tr>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Apply On</b></font></td>
					    <td width="10%" bgcolor="#F3F3F3"><font class="marineblack"><b>Date Apply For</b></font></td>
 		                <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>EmpID</b></font></td>
 					    <td width="15%" bgcolor="#F3F3F3"><font class="marineblack"><b>Name</b></font></td>
				  		<td width="9%" bgcolor="#F3F3F3"><font class="marineblack"><b>Department</b></font></td>
					    <td width="6%" bgcolor="#F3F3F3"><font class="marineblack"><b>Leave ID</b></font></td>
					    <td width="7%" bgcolor="#F3F3F3"><font class="marineblack"><b>Day</b></font></td>
					    <td width="18%" bgcolor="#F3F3F3"><font class="marineblack"><b>Reason</b></font></td>
		            </tr>            
				 <%end if%>
			
			<!--<% if request("txtSearch") ="Search" or vStatus="" then 'vStatus="P"
			       tempvStatus = vStatus
			       tempDate1 = request("txtdate1")
			       tempDate2 = request("txtdate2")
			       tempVID = request(vID)
			       tempSearchBy = request("cboSearchBy")
			       call inputsessionApp(tempvStatus,tempDate1,tempDate2,tempVID,tempSearchBy)
			   end if 
			Response.Write"hello"
			Response.Write(tempSearchBy)
			Response.Write(tempVID)
		
			%> -->      
				 
            <%
              IF vStatus="P" or vStatus="R" or vStatus="A" then
               dim vID
               
               if SearchBy = "" then
               	 vID = "%"
               else
               	 vID = request("txtID") '+ "%"
               end if
			   sessionValue= session("ssStatus") 
			   if request("txtSearch") ="Search" or (vStatus="P" and sessionValue <>"P") or (vStatus="R" and sessionValue <>"R") or (vStatus="A" and sessionValue <>"A") then 'vStatus="P"
			       tempvStatus = vStatus
			       tempDate1 = request("txtdate1")
			       tempDate2 = request("txtdate2")
			       tempVID = vID
			       tempSearchBy = request("cboSearchBy")
			       call inputsessionApp(tempvStatus,tempDate1,tempDate2,tempVID,tempSearchBy)
			   end if 
			'Response.Write"hello"
			'Response.Write(tempSearchBy)
			'Response.Write(tempVID) 
			    			    
			   page_size =5
			
			   if request("page") = "" then
				current_page = 1
			   else
				current_page = CInt(request("page"))
			   end if
			   		   	
					set myconn = server.CreateObject("ADODB.Connection")
			        set rs = server.CreateObject("ADODB.Recordset")
		                myconn.open connect_string
            	    
				    ssql = "Exec sp_Wls_LeaveApproval '" + Session("Regisno") + "', '" + Session("EmpID") + "', 'ENG', '" _
					         + session("ssStatus") + "', '" + session("ssID") + "', '', '" + session("ssDate1") + "', '" + session("ssDate2") + "', '" + session("ssSearchBy") + "'" 
				     
				    'Response.Write ssql
				   rs.cursorlocation = adUseClient
		           rs.pagesize = page_size      
				 
				   rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 					 
			     
			        'page_count = rs.RecordCount         
					page_count = rs.pagecount
			         
			        if 1 > current_page then current_page = 1
			         	 
			        if current_page > page_count then current_page = page_count				    
					if rs.RecordCount = 0 then
					   current_page =1
					   page_count = 1
					end if 
					  	
					if rs.RecordCount > 0 then	          
			        rs.AbsolutePage = current_page
					'Response.Write"hello"
				    end if
			   	    
 '*****************************   vStatus = "P"   **********************************************************                                 
				   If vStatus = "P" then
			          
					 colour = 0
					
			        do while rs.AbsolutePage = current_page and not rs.EOF
			           'Do Until webdbRecordset.EOF
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
							 end if  'regisno=" + Session("Regisno") + "?
						   temp = rs("applyon")	
				           response.write "<tr>"
					       response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' onclick='UncheckR(""" + cstr(rowno) + """)' name=A" + cstr(rowno) + " value='ON' " + Astrcheck + "></font></td>"
					       response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' onclick='UncheckA(""" + cstr(rowno) + """)' name=R" + cstr(rowno) + " value='ON' " + Rstrcheck + "></font></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("applyon") + "<input type='hidden' name=O" + cstr(rowno) + " value= " + rs("applyon") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("applyfor") + "<input type='hidden' name=F" + cstr(rowno) + " value= " + rs("applyfor") + "></td>"
					       response.write "<td align='left'" + colour + "><font class='small'>" +  "<a href=""leavebalance17.asp?employeeid=" + UCase(rs("empid")) + """onclick=""NewWindow(this.href,'LeaveBalance','700','480','yes','center');return false""onfocus=""this.blur()"">" + rs("empid") + "</a>"  +  "<input type='hidden' name=I" + cstr(rowno) + " value= " + rs("empid") + "></td>"	   						   'response.write "<td align='left'" + colour + "><font class='small'>" +"<a href=""leavebalance3.asp?"employeeid="+webdbRecordset.Fields("empid")+""">"+ webdbRecordset.Fields("empid") + "</a>" + "<input type='hidden' name=I" + cstr(rowno) + " value= " + webdbRecordset.Fields("empid") + "></td>"
				           'Pending
				           'onclick=""NewWindow(this.href,'LeaveBalance','700','480','yes','center');return false""
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("empname") + "<input type='hidden' name=N" + cstr(rowno) + " value= " + rs("empname") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("deptid") + "<input type='hidden' name=D" + cstr(rowno) + " value= " + rs("deptid") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("leaveid") + "<input type='hidden' name=L" + cstr(rowno) + " value= " + rs("leaveid") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("period") + "<input type='hidden' name=P" + cstr(rowno) + " value= " + rs("periodid") + "></td>"
				           
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("reason") + "<input type='hidden' name=S" + cstr(rowno) + " value= " + rs("reason") + "></td></tr>"
				           rs.MoveNext  
				           count = abs(count - 1)
			           loop
	  				 response.write "<input type=hidden name=txtRowNo value=" + cstr(rowno) + ">"
'	  				 ApproveRow = rowno
		  			 rs.close
			         set rs = nothing
			         myconn.close
			         set myconn = nothing 	 
			         'webdbRecordset.close
			         'webdb.close
'******************************************   end vStatus ="P"    *****************************************************************************
					
'******************************************    vStatus ="A"    ********************************************************************************					
					 ElseIf vStatus = "A" then
					   dim tmpApplyfor
					   dim tmpLockDate
			           do while rs.AbsolutePage = current_page and not rs.EOF 'Do Until webdbRecordset.EOF
					       rowno = rowno + 1			        
			           
				           if count = 1 then
				              colour = " bgcolor='#eeeeee'"
 				           else
				              colour = ""
				           end if
				           
				           temp = rs("applyon")	
				           response.write "<tr>"
				           tmpApplyfor = mid(rs("applyfor"),4,2) + "/" + mid(rs("applyfor"),1,2) + "/" + mid(rs("applyfor"),7,4)
				           tmpLockDate = mid(rs("lockdate"),4,2) + "/" + mid(rs("lockdate"),1,2) + "/" + mid(rs("lockdate"),7,4)    
				           if cdate(tmpApplyfor) < cdate(tmpLockDate) then
						     response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' disabled=true name=R" + cstr(rowno) + " value='ON' " + strcheck + "></font></td>"
						   else
						     response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' name=R" + cstr(rowno) + " value='ON' " + strcheck + "></font></td>"
						   end if
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("applyon") + "<input type='hidden' name=O" + cstr(rowno) + " value= " + rs("applyon") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("applyfor") + "<input type='hidden' name=F" + cstr(rowno) + " value= " + rs("applyfor") + "></td>"
				           
				           response.write "<td align='left'" + colour + "><font class='small'>" +  "<a href=""leavebalance17.asp?employeeid=" + UCase(rs("empid"))+ """onclick=""NewWindow(this.href,'LeaveBalance','700','480','yes','center');return false""onfocus=""this.blur()"">" + rs("empid") + "</a>"  +  "<input type='hidden' name=I" + cstr(rowno) + " value= " + rs("empid") + "></td>"
				           'Approved
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("empname") + "<input type='hidden' name=N" + cstr(rowno) + " value= " + rs("empname") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("deptid") + "<input type='hidden' name=D" + cstr(rowno) + " value= " + rs("deptid") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("leaveid") + "<input type='hidden' name=L" + cstr(rowno) + " value= " + rs("leaveid") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("period") + "<input type='hidden' name=P" + cstr(rowno) + " value= " + rs("periodid") + "></td></tr>"
				           'response.write "<td align='left'" + colour + "><font class='small'>" + rs("reason") + "<input type='hidden' name=S" + cstr(rowno) + " value= " + rs("reason") + "></td></tr>"
				           rs.MoveNext  
				           count = abs(count - 1)
			           loop
		  				 response.write "<input type=hidden name=txtRowNo value=" + cstr(rowno) + ">"
			            rs.close
			           set rs = nothing
			            myconn.close
			           set myconn = nothing 
			           'webdbRecordset.close
			           'webdb.close
'*****************************************     end vStatus="A"  ************************************************************************************************************************************			           

'*****************************************     vStatus="R"    **************************************************************************************************************************************					
					 ElseIf vStatus = "R" then 			 
					  'else
			          do while rs.AbsolutePage = current_page and not rs.EOF 'Do Until webdbRecordset.EOF
					       rowno = rowno + 1			        
			           
				           if count = 1 then
				              colour = " bgcolor='#eeeeee'"
 				           else
				              colour = ""
				           end if
			               temp = rs("applyon")	
				           response.write "<tr>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("applyon") + "<input type='hidden' name=O" + cstr(rowno) + " value= " + rs("applyon") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("applyfor") + "<input type='hidden' name=F" + cstr(rowno) + " value= " + rs("applyfor") + "></td>"
				           
				           response.write "<td align='left'" + colour + "><font class='small'>" + "<a href='leavebalance17.asp?employeeid=" + UCase(rs("empid"))+ "'onclick=""NewWindow(this.href,'LeaveBalance','700','480','yes','center');return false""onfocus=""this.blur()"">" + rs("empid") + "</a>"  + "<input type='hidden' name=I" + cstr(rowno) + " value= " + rs("empid") + "></td>"
				           'Reject
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("empname") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("deptid") + "</td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("leaveid") + "<input type='hidden' name=L" + cstr(rowno) + " value= " + rs("leaveid") + "></td>"
				           response.write "<td align='left'" + colour + "><font class='small'>" + rs("period") + "<input type='hidden' name=P" + cstr(rowno) + " value= " + rs("periodid") + "></td></tr>"
				           'response.write "<td align='left'" + colour + "><font class='small'>" + rs("reason") + "<input type='hidden' name=S" + cstr(rowno) + " value= " + rs("reason") + "></td></tr>"
				           rs.MoveNext  
				           count = abs(count - 1)
			           loop
			            rs.close
			           set rs = nothing
			            myconn.close
			           set myconn = nothing 
			           'webdbRecordset.close
			           'webdb.close
'**************************************************    end vStaus ="R"   *************************************************************************************************************************************			           
					 End If	
				end if	 				 
			 %>
			 

			 <%
			   if Request.form("txtAction")="UPD" and vStatus = "P" then
	
				   Set webdb = Server.CreateObject("ADODB.Connection")
				   		webdb.Open Session("ConnectStr")
				   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
				   Set webdbCommand = Server.CreateObject("ADODB.Command")
				   Set webdbCommand.ActiveConnection = webdb
	
'				   maxrow = approverow
				   maxrow = request.form("txtRowNo")
				   approve ="false"

			      do until rowcount = cint(maxrow)
				      rowcount = rowcount + 1

			      if Request.Form("N" + cstr(rowcount)) <> "0" then
					  approve = "true"
			      	  ssql = "Exec sp_Wls_updLeaveApproval '" & Session("Regisno") & "', '" _
			      	  		  & Request.Form("O"+ cstr(rowcount)) & "', '" & Request.Form("F"+ cstr(rowcount)) & "', '" _
			      	  		  & Session("CurrentDate") & "','" & Request.Form("I"+ cstr(rowcount)) & "' , '" _
			      	         & Request.Form("L"+ cstr(rowcount)) & "', '" & Request.Form("P"+ cstr(rowcount)) & "', 'ENG','" _
			      	         & Session("EmpID") & "', '" + "A" + "'"
			      	         
					  webdbCommand.CommandText = ssql
					  webdb.Execute webdbCommand.CommandText
			      end if
			      

			      if Request.Form("D" + cstr(rowcount)) <> "0" then
					  approve = "true"
			      	  ssql = "Exec sp_Wls_updLeaveApproval '" & Session("Regisno") & "', '" _
			      	  		  & Request.Form("O"+ cstr(rowcount)) & "', '" & Request.Form("F"+ cstr(rowcount)) & "', '" _
			      	  		  & Session("CurrentDate") & "','" & Request.Form("I"+ cstr(rowcount)) & "' , '" _
			      	         & Request.Form("L"+ cstr(rowcount)) & "', '" & Request.Form("P"+ cstr(rowcount)) & "', 'ENG','" _
			      	         & Session("EmpID") & "', '" + "R" + "'"
					  webdbCommand.CommandText = ssql
					  webdb.Execute webdbCommand.CommandText
			      end if				      
		      loop
		      
		      if approve = "true" then
		         response.redirect "/ehres3/Leave/updateSuccess.asp"
		      end if
		    end if	
		%>

			 <%
			   if Request.form("txtAction")="UPD" and vStatus = "A" then
	
				   Set webdb = Server.CreateObject("ADODB.Connection")
				   		webdb.Open Session("ConnectStr")
				   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
				   Set webdbCommand = Server.CreateObject("ADODB.Command")
				   Set webdbCommand.ActiveConnection = webdb
	
				   maxrow = request.form("txtRowNo")
				   approve ="false"

			      do until rowcount = cint(maxrow)
				      rowcount = rowcount + 1			      

			      if Request.Form("D" + cstr(rowcount)) <> "0" then
					  approve = "true"
			      	  ssql = "Exec sp_Wls_updLeaveApproval '" & Session("Regisno") & "', '" _
			      	  		  & Request.Form("O"+ cstr(rowcount)) & "', '" & Request.Form("F"+ cstr(rowcount)) & "', '" _
			      	  		  & Session("CurrentDate") & "','" & Request.Form("I"+ cstr(rowcount)) & "' , '" _
			      	         & Request.Form("L"+ cstr(rowcount)) & "', '" & Request.Form("P"+ cstr(rowcount)) & "', 'ENG','" _
			      	         & Session("EmpID") & "', '" + "R" + "'"
					  webdbCommand.CommandText = ssql
					  webdb.Execute webdbCommand.CommandText
			      end if	
		      loop
		      
		      if approve = "true" then
		         response.redirect "/ehres3/Leave/updateSuccess.asp"
		      end if
		    end if	
		%>	   	
							 
                <tr>
                  <td width="6%">&nbsp;</td>
                  <td width="50%" colspan="5">&nbsp;</td>
                  <td width="44%" colspan="4">&nbsp;</td>
                </tr>
                <tr>
                  <td width="6%">&nbsp;</td>
                  <td width="50%" colspan="5">&nbsp;</td>
                  <td width="44%" colspan="4">&nbsp;</td>
                </tr>
                <tr>
                  <td width="6%">&nbsp;</td>
                  <td width="50%" colspan="5">
        
        		  <% if vStatus = "P" and temp <> "" then%>
		            <input type="button" value="Select All Approve" name="cmdSelectAllA" onclick="checkAllApprove()" class="small">
        		  <% end if %>
        		  
        		  <% if (vStatus = "P" and temp <> "") or (vStatus = "A" and temp <> "") then%>
		            <input type="button" disabled=true value="Select All Reject" name="cmdSelectAllR" onclick="checkAllReject()" class="small">
		            <input type="submit" value="Update" name="cmdUpdate" <% end if%>
		             <%if vStatus = "P" and temp <> "" then%> onclick="ApproveRejectLeave()" <%elseif vStatus = "A" and temp <> "" then%> onclick="RejectLeave()" <%end if%> <% if (vStatus = "P" and temp <> "") or (vStatus = "A" and temp <> "") then%> class="small">
        		  <% end if %></TD>				    	
				 <!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-->
				 <!--employeeid=" + Request.Form("empid") +-->
				 <!--<% response.write "<a href = 'leavebalance3.asp?" + "'onclick=""NewWindow(this.href,'LeaveBalance','700','480','yes','center');return false""onfocus=""this.blur()"">"%><font class="marineblue">Leave Balance</font></a>-->	
		         <input type="hidden" name="txtAction">
		         <input type="hidden" name="txtAppAll">
		         <input type="hidden" name="txtRejectAll">
		         
                  <!--</td>-->
                  <td width="44%" colspan="4"><% response.write "<a href = 'leavebalance3.asp?" + "'onclick=""NewWindow(this.href,'LeaveBalance','700','480','yes','center');return false""onfocus=""this.blur()"">"%><font class="marineblue">Leave Balance</font></a>	&nbsp;</td>
                </tr>
                <tr>
                  <td width="6%">&nbsp;</td>
                  <td width="50%" colspan="5"></td>
                  <td width="44%" colspan="4"></td>
                </tr>
              </table>
              
              <table cellSpacing="0" cellPadding="1" border="0" width="100%" bordercolor="#808080">
      
			<% if vStatus ="P" or vStatus="A" or vStatus="R" or vStatus="" then%>
		    <p align=center>
			
			<%Response.Write "<br>" 
			Response.Write "<td colspan=""4"" align=""center"">"
  ''''''''''''''''''''''''''''''''''''''''''''''paging records start'''''''''''''''''''''''''''''''''''''''''''''''''
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "First</font><font=""2""> | </font>"
			end if
  
			iF current_page >= 2 then
				Response.Write "<a href=""app_approval.asp?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""app_approval.asp?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""app_approval.asp?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""app_approval.asp?page="
				Response.Write page_count
				Response.Write """><font face=""Verdana"" & size=""1"">Last >></font></a>" & vbCrlf
			end if 
  
			if current_page >= page_count then 
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Last</font>" & "</font>"
			end if
      ''''''''''''''''''''''''''''''''''''''''paging records end''''''''''''''''''''''''''''''''''''''''''''''''''              
			Response.Write "</center>"%>
		
			<font face=Verdana size=1><center>Page <%=current_page%> of <%=page_count%></center>
			<% end if %>
	
			
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
    <td width="100%" align="center"><img border="0" src="/eHres3/Image/dottedlinenav.gif" WIDTH="408" HEIGHT="4"></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<br>Copyright © 1997-2000 Software
      Factory Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>
<p>&nbsp;</p>
</html>