<!-- #include virtual ="/ehres/global/ConnectStrKPI.asp"-->
    
<html><title>eHRES</title>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">

<body bgColor="#ffffff" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<script language="vbscript">
<!--
function View()
	   document.frmDelLeave.submit()
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
          response.write session("EmpID") + session("Num")
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
  <tr>
    <td vAlign="top" align="center" colspan="2" width="936" height="100%">
      <div align="center">
        <center>
         <form method="POST" action="pm_AceptAmend.asp" name="frmDelLeave">
          <table border="0" width="100%" bordercolor="#808080" cellSpacing="0" cellPadding="0">
           <tr></tr>
           <% 
						Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		           Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           Set webdbCommand = Server.CreateObject("ADODB.Command")
  		           
		           ssql = "Exec sp_kpi_SelAverage """ + request.Form("D1") + """,""" + request.Form("PM1") + """"
					 
			        Set webdbCommand.ActiveConnection = webdb
			            webdbCommand.CommandText = ssql
			            webdbRecordset.Open webdbCommand,,1 , 3
			           
						response.write "<tr>"
						response.write "<td colspan=4> </td>"
						response.write "<td align='left'" + colour + "><font class='marineblack'>Average Mark: " + webdbRecordset.Fields("Average") + "</td></tr>"
						  
						webdbRecordset.close
			        webdb.close
			%>
            <tr>
              <td height="20" align="center" width="3%"><font class="marineblack"> </font></td>            
              <td height="20" align="center" width="7%" bgcolor="#F3F3F3"><font class="marineblack">Accept</font></td>
              <td height="20" align="Left" width="7%" bgcolor="#F3F3F3"><font class="marineblack">Amend</font></td>
              <td height="20" width="30%" bgcolor="#F3F3F3"><font class="marineblack">Criteria</font></td>
              <td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Rating</font></td>
              <td height="20" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Enter By</font></td>              
            </tr>

            <%   
            	   Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open Session("ConnectStr")
  		           Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           Set webdbCommand = Server.CreateObject("ADODB.Command")
  		           
		           ssql = "Exec sp_KPI_SelKPIDetail """ + request.Form("D1") + """,""" + request.Form("PM1") + """"
					 
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
						response.write "<td align='center'" + colour + "><font class='small'><input type='CheckBox' name=A"+cstr(rowno)+" value='A'"+cstr(rowno)+"'></font></td>" 
						response.write "<td align='left'" + colour + "><font class='small'><input type='CheckBox' name=B"+cstr(rowno)+" value='B'"+cstr(rowno)+"'></font></td>" 
				        'response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("EmpID") + "<input type='hidden' name=D" + cstr(rowno) + " value= " + webdbRecordset.Fields("EmpID") + "></td>"
				        'response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("PMno") + "<input type='hidden' name=PM" + cstr(rowno) + " value= " + webdbRecordset.Fields("PMno") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("IDName") + "<input type='hidden' name=C" + cstr(rowno) + " value= " + webdbRecordset.Fields("IDName") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("RatingName") + "<input type='hidden' name=R" + cstr(rowno) + " value= " + webdbRecordset.Fields("RatingName") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("empname") + "</td></tr>"
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
          <td width="94%" height="19"><input type="submit" value="Update" name="cmdUpdate" onclick="View()" class="small"></td>
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
    <td width="100%" align="center"><img border="0" src="/eHres/Image/dottedlinenav.gif" WIDTH="408" HEIGHT="4"></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<br>Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>. </td></tr>
</table>

<p>&nbsp;</p>
</html>

          