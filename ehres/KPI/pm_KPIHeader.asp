<!-- #include virtual ="/ehres/global/ConnectStrKPI.asp"-->
     
<html><title>eHRES</title>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">

<body bgColor="#ffffff" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<script language="vbscript">
<!--
function View()
	   dim rowcount 
	dim ssql
	dim maxrow
	dim action
	dim del
	maxrow = document.frmDelLeave.txtRowNo.value	
	action = ""
    if isnumeric(maxrow) then	
	   do until rowcount = cint(maxrow)
	      rowcount = rowcount + 1
	      ssql="if " + "document.frmDelLeave.A" + cstr(rowcount) + ".checked = false then" + chr(10) 
	      ssql= ssql + " document.frmDelLeave.D" + cstr(rowcount) + ".value=0" + chr(10)
	      ssql= ssql + " document.frmDelLeave.PM" + cstr(rowcount) + ".value=0" + chr(10) 
	      ssql=ssql + "end if"
	
  	      execute ssql
	   loop
	   document.frmDelLeave.submit()
	end if
end function
// -->
</script>


<div align="center">
  <center>
<table cellSpacing="0" cellPadding="0" border="0" width="100%" height="100">
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
    <td vAlign="top" align="center" width="27" height="109"></td>
    <td vAlign="top" align="center" width="907" height="109"><br><img alt="Main Menu" src="../Image/pmType.gif" border="0" width="696" height="89"><br>
      &nbsp;</td></tr>
  <tr><TD colspan = "2" align = "center" width = "100%">
           <FONT class=small>Employee ID&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT>
            <select name=cboempid style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px">  
       			<%  dim tmpEmpID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_kpi_selSubordinate '"+ trim(Session("EmpID")) + "'"
         	   
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpEmpID = ""
	  			txtdate1 = ""
	  			txtdate2 = ""
	  			tmpEmpID = session("sscboempid")
			 	
			 	if Request.form("cboempid") = "" or tmpEmpID = "1" then 
	              response.write "<option selected value='1'>"+session("EMPID")+ " - " + session("EmpName")+ "</option>"
	    
	            else  
	              response.write "<option value='1'>"+session("EMPID")+ " - " + session("EmpName")+ "</option>"
	            end if 
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("empid")) = Request.form("cboempid") ) or ( trim(webdbRecordset.Fields("empid")) = tmpEmpID ) or ( trim(webdbRecordset.Fields("empid")) = session("sscboempidlv"))then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
 				    end if
 				    
 				    if tmpEmpID = "" OR tmpEmpID <> "" then
					      tmpEmpID = trim(webdbRecordset.Fields("empid"))
					      
					end if   
			  	 
				   webdbRecordset.MoveNext  
		        loop       
				Response.Write(tmpEmpID)
			%></select>&nbsp;&nbsp;&nbsp; 
           </TD>
           </tr> 
           <tr><td colspan=2><p>&nbsp;</p></td></tr>
  <tr>
    <td vAlign="top" align="center" colspan="2" width="936" height="100%">
      <div align="center">
        <center>
         <form method="POST" action="pm_kpidetail.asp" name="frmDelLeave">
          <table border="0" width="100%" bordercolor="#808080" cellSpacing="0" cellPadding="0">
            <tr>
              <td height="20" align="center" width="7%"><font class="marineblack"> </font></td>            
              <td height="20" align="center" width="10%" bgcolor="#F3F3F3"><font class="marineblack">View</font></td>
              <td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">EmpID</font></td>
              <!--<td height="20" width="15%" bgcolor="#F3F3F3"><font class="marineblack">Performace No</font></td>//-->
              <td height="20" width="18%" bgcolor="#F3F3F3"><font class="marineblack">Current Position</font></td>
              <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Title</font></td>
              <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Status</font></td>              
            </tr>

            <%   
            	   Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		           Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           Set webdbCommand = Server.CreateObject("ADODB.Command")

		           ssql = "Exec sp_KPI_SelKPIHeader """ + Session("EmpID") + """"
					
				   
			        Set webdbCommand.ActiveConnection = webdb
			            webdbCommand.CommandText = ssql
			            webdbRecordset.Open webdbCommand,,1 , 3

					 colour = 0

			        Do Until webdbRecordset.EOF
					    rowno = rowno + 1			        
					 
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
				        
				        response.write "<tr>"
						response.write "<td> </td>" 				        
						response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' name=A"+cstr(rowno)+" value='View'"+cstr(rowno)+"'></font></td>" 
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("EmpID") + "<input type='hidden' name=D" + cstr(rowno) + " value= " + webdbRecordset.Fields("EmpID") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("Position") + "<input type='hidden' name=P" + cstr(rowno) + " value= " + webdbRecordset.Fields("Position") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("Subject") + "<input type='hidden' name=S" + cstr(rowno) + " value= " + webdbRecordset.Fields("Subject") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("status") + "</td></tr>"
				        response.write "<font class='small'><input type='hidden' name=PM" + cstr(rowno) + " value= " + webdbRecordset.Fields("PMno") + ">"
				        webdbRecordset.MoveNext  
				        count = abs(count - 1)   
			        loop
	  				 response.write "<input type=hidden name=txtRowNo value=" + cstr(rowno) + ">"

			        webdbRecordset.close
			        webdb.close      
			 %>

          </table>
          </td>
        </tr>
        <tr>
          <td height="19" colspan=4></td>
        </tr>
        <tr>
			<td width="6%" height="19"></td>
          <td width="94%" height="19"><input type="button" value="View" name="cmdView" onclick="View()" class="small"></td>
        </tr>
      </table>
    </form>
    <p>&nbsp;</p>

    </td>
  </tr>
</table>
<p>&nbsp;</p>
<table border="0" width="96%">
  <tr>
    <td width="100%" align="center"><img border="0" src="/eHres3/Image/dottedlinenav.gif" WIDTH="408" HEIGHT="4"></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<br>Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>

<p>&nbsp;</p>
</html>

          