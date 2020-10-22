<!-- #INCLUDE FILE = "../global/ConnectStr.asp"-->
<!-- #INCLUDE FILE = "../global/AdoVbs.asp"-->
<!-- #INCLUDE FILE = "../global/inputSession.asp"-->
<% dim row

row = request.Form("txtAction")

%>
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

<% dim connect_string
       connect_string =Session("ConnectStr")
%>
       
	<SCRIPT LANGUAGE="vbScript">
	function Back()
	document.frmEmpScore.action = "TD_GapAnalysis2.asp"
	document.frmEmpScore.submit()
	end function

	function Change()
		document.frmEmpScore.submit()
	End function
	
	function check()
		dim maxrow
		dim rowcount
		maxrow = document.frmEmpScore.txtrowno.value
		document.frmEmpscore.tolscore.value = 0
		if isnumeric(maxrow) then
			do until rowcount = cint(maxrow)
				rowcount = rowcount + 1
				ssql = " if isnumeric(document.frmEmpScore.S" + cstr(rowcount) + ".value) = true then" + chr(10)
				ssql = ssql + " if cint(document.frmEmpscore.S" + cstr(rowcount) + ".value) <= cint(document.frmEmpScore.Max" + cstr(rowcount) + ".value) then" + chr(10)
				ssql = ssql + " document.frmEmpscore.tolscore.value = cint(document.frmEmpScore.tolscore.value) + cint(document.frmEmpScore.S" + cstr(rowcount) + ".value)" + chr(10)
				ssql = ssql + " else " + chr(10)
				ssql = ssql + " document.frmEmpScore.S" + cstr(rowcount) + ".value = 0" + chr(10)
				ssql = ssql + " end if" + chr(10)
				ssql = ssql + " else " + chr(10)
				ssql = ssql + " document.frmEmpScore.S" + cstr(rowcount) + ".value = 0" + chr(10)
				ssql = ssql + " end if"
				
				execute ssql
				
			loop
		end if
		
	end function
	
	function validate()
		document.frmEmpScore.submit()
	end function
	
	</SCRIPT>


<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div align="center">
  <center>
<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%" height="392">
  <TBODY>
  <TR>
    <TD vAlign=top align=middle colspan="2" width="936" bgcolor="#0099cc" height="29">
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
    </TD></TR>
  <TR>
    <TD vAlign=top colspan="2" width="100%" height="21" class="small" align="middle">
      <p align="right"><A href="../main.asp"><font color="#000000">Home</font></A>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <A href="../signout.asp"><font color="#000000">Logout</font></A></p></TD></TR>
  <TR>
    <TD vAlign=top align=middle width="27" height="109"></TD>
    <TD vAlign=top align=middle width="907" height="109">
      <P><IMG height=84 src="../Image/enggapana.gif" width=683 border=0 ><BR><BR>
        <FORM name=frmEmpScore action=TD_UpdateEmpScore2.asp method=post>
      <TABLE cellSpacing=0 width="100%" border=0>
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <TD>
           <FONT class=small><b>Employee ID &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
            <select name=cboempid style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px">  
       			<%  dim tmpEmpID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WTD_SelEvalutionEmp '" + Session("Regisno") + "','" + Session("EmpID") + "'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpEmpID = ""
	  			txtdate1 = ""
	  			txtdate2 = ""
	  			tmpEmpID = Request.form("cboempid")
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("Empid")) = Request.form("cboempid") ) or ( trim(webdbRecordset.Fields("empid")) = tmpEmpID )then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("Empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("Empid")) + ">"  + " " + trim(webdbRecordset.Fields("empid")) + " " + "-" + " " + trim(webdbRecordset.Fields("empname")) + "</option>"
 				    end if
 				    
 				    if tmpEmpID = "" then
					      tmpEmpID = trim(webdbRecordset.Fields("Empid"))
					      
					end if 
				   webdbRecordset.MoveNext  
		        loop       
			%></select>&nbsp;&nbsp;&nbsp; 
           </TD>
        </TR>
        <tr><TD WIDTH=10%>&nbsp;</TD>
        <TD>
       			<%  dim tmpDeptID
       			    dim tmpSecID
       			    dim tmpJobID
       			    dim tmpcurrentID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_Wtd_selGapAnalysis '" + Session("Regisno") + "','" + tmpEmpID + "','','',0,'EMPDETAIL'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpDeptID = ""
	  			tmpSecID = ""
	  			tmpJobID = ""
	  			tmpcurrentID = ""
	  			tmpDeptID = Request.form("txtDept")
	  			tmpSecID = Request.form("txtSec")
	  			tmpJobID = Request.form("txtJob")
	  			tmpcurrentID = Request.form("txtCPos")

		        if webdbrecordset.eof then
					tmpDeptID = ""
	  				tmpSecID = ""
	  				tmpJobID = ""
	  				tmpcurrentID = ""
	  			else
	  				tmpDeptID = trim(webdbRecordset.Fields("dept"))
	  				tmpSecID = trim(webdbRecordset.Fields("sec"))
	  				tmpJobID = trim(webdbRecordset.Fields("job"))
	  				tmpcurrentID = trim(webdbRecordset.Fields("pos"))
	  			end if
			%><FONT class=small><b>Department &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
			  <input type=text name=txtDept readonly size=30 value="<%=tmpDeptID%>">&nbsp;&nbsp;&nbsp;
			   <FONT class=small><b>Section &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
			  <input type=text name=txtSec readonly size=30 value="<%=tmpSecID%>">&nbsp;&nbsp;&nbsp;
           </TD>
        </tr>
        <tr><td width=10%>&nbsp;</td>
        <td><FONT class=small><b>Job Title &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
			  <input type=text name=txtJob readonly size=30 value="<%=tmpJobID%>">&nbsp;&nbsp;&nbsp;
        </td>
        </tr>
        <tr>
        <td>&nbsp;</td>
        </tr>
        <tr><td width=10%></td>
        <td><FONT class=small><b>Target Position&nbsp;&nbsp;&nbsp;</b> </FONT>
        <input name=txtTPosDesc style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px" type=text readonly value="<%=trim(Left(request.Form("cboTPos"),200))%>">
        <input name=txtTPosID type=hidden value=<%=trim(Right(request.Form("CboTPos"),20))%>>
        &nbsp;&nbsp;&nbsp;
        </td>
        </tr>
        <tr>
        <td>&nbsp;</td>
        </tr>
        
  <TR>
    <TD vAlign=top align=middle colspan=2 width="936" height="193">
      <div align="center">
        <center>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="1">
        <TBODY>
    <tr>
      <td bgcolor="#ffffff" width="30">&nbsp;</td>
      <td bgcolor="#f3f3f3" width="500"><font class="marineblack">Competency Structure</font></td>
      <td bgcolor="#f3f3f3" width="150"><font class="marineblack">Max Score</font></td>
      <td bgcolor="#f3f3f3" width="150"><font class="marineblack">Employee Score</font></td>
    </tr>

            <%     page_size =10
			
					if request("page") = "" then
						current_page = 1
					else
						current_page = CInt(request("page"))
					end if
					
                    set myconn = server.CreateObject("ADODB.Connection")
			        set rs = server.CreateObject("ADODB.Recordset")
		                myconn.open connect_string
		                 
		            rs.cursorlocation = adUseClient
		            rs.pagesize = page_size
		            
                    if tmpCurrentID ="" then                
						 ssql = "Exec sp_WTD_SelCompEvaluteScore '" + Session("Regisno") + "','" + tmpEmpID + "','" + request.Form("HCP"+ cstr(row)) + "'"
					else	
					     ssql = "Exec sp_WTD_SelCompEvaluteScore '" + Session("Regisno") + "','" + tmpEmpID + "','" + request.Form("HCP"+ cstr(row)) + "'"
					end if
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
 
			        page_count = rs.pagecount
			        
			        if 1 > current_page then current_page = 1
			        if current_page > page_count then current_page = page_count
					
					if rs.RecordCount >0 then
			        rs.AbsolutePage = current_page
			        dim tolmaxscore
			        dim tolscore		       
					colour = 0
					rowno = 0 
			        do while rs.AbsolutePage = current_page and not rs.EOF
			        
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			            rowno = rowno + 1
			            tolmaxscore = tolmaxscore + cint(rs("maxscore"))
			            tolscore = tolscore + cint(rs("score"))
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>"+ rs("description") + "<input type=hidden name=EVa" + cstr(rowno) + " value='" + rs("EvaluationID") + "'></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + cstr(rs("maxscore")) + "<input type=hidden name=max" + cstr(rowno) + " value =" + cstr(rs("maxscore")) + "></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=text name=S" + cstr(rowno) +" value=" + cstr(rs("score")) + " onchange=check()></td>"
				        Response.Write "</tr>"
				       
				        rs.Movenext  
				        count = abs(count - 1)        
			        loop
			        response.write "<tr><font class='marineblack'><b>"
				    response.write "<td width='10%'>&nbsp;</td>"
				    response.Write "<td >Total Score</td>"
				    response.Write "<td>" + cstr(tolmaxscore) + "</td>"
				    response.Write "<td><input type=text name=tolscore readonly value=" + cstr(tolScore) + "></td>"
				    response.Write "</b></font></tr>"
				    
				    response.write "<tr>"
				    response.write "<td width='10%'>&nbsp;</td>"
				    response.Write "</tr>"
				    
			        response.write "<tr>"
				    response.write "<td width='10%'></td>"
			        response.Write "<td bgcolor='#ffffff' width='30'><input type=button name=cmdUpdate onclick=validate() value=Update size=50></td>"
			        response.write "<td><input name=cmdBack onclick=Back() value=Back type=button><input type=hidden name=txtRowNo value=" + cstr(rowno) + "><input type=hidden name=txtAction value=" + request.Form("HCP"+ cstr(row))+"></td>"
			        response.Write "</tr>"
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
			      else
			        response.write "<tr>"
				    response.write "<td width='10%'>&nbsp;</td>"
				    response.Write "</tr>"
				    
			        response.write "<tr>"
				    response.write "<td width='10%'></td>"
			        response.Write "<td bgcolor='#ffffff' width='30'><input type=button name=cmdUpdate onclick=validate() value=Update size=50></td>"
			        response.write "<td><input name=cmdBack onclick=Back() value=Back type=button><input type=hidden name=txtRowNo value=" + cstr(rowno) + "><input type=hidden name=txtAction value=" + request.Form("HCP"+ cstr(row))+"></td>"
			        response.Write "</tr>"
			      end if  
			       
			%> 
			      
			 <table cellSpacing="0" cellPadding="1" border="0" width="90%" bordercolor="#808080">

		    <p align=center>
			
			<%Response.Write "<br>" 
			Response.Write "<td colspan=""4"" align=""center"">"
  ''''''''''''''''''''''''''''''''''''''''''''''paging records start'''''''''''''''''''''''''''''''''''''''''''''''''
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "First</font><font=""2""> | </font>"
			end if
  
			iF current_page >= 2 then
				Response.Write "<a href=""TD_EmpScore2.asp?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""TD_EmpScore2.asp?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""TD_EmpScore2.asp?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""TD_EmpScore2.asp?page="
				Response.Write page_count
				Response.Write """><font face=""Verdana"" & size=""1"">Last >></font></a>" & vbCrlf
			end if 
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Last</font>" & "</font>"
			end if
      ''''''''''''''''''''''''''''''''''''''''paging records end''''''''''''''''''''''''''''''''''''''''''''''''''              
			Response.Write "</center>"%>
		
			<font face=Verdana size=1><center>Page <%=current_page%> of <%=page_count%></center>
			
			</table>     
		
            </TBODY>
            </TABLE></center>
      </div>
    </TD></TR>
    <center>  
    <TD align=center colspan=2 width="936" height="40" class="small"><br>
      <!--&nbsp;<br>
      &nbsp;<BR>--><font class ="small" >Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>.</font></TD></TR></TBODY></TABLE></center>
      </div>
</BODY>




