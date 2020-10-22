<!-- #INCLUDE FILE = "../global/ConnectStr.asp"-->
<!-- #INCLUDE FILE = "../global/AdoVbs.asp"-->
<!-- #INCLUDE FILE = "../global/inputSession.asp"-->
<%
if request.Form("txtAction") = "Upd" then
   Set webdb = Server.CreateObject("ADODB.Connection")
	   		webdb.Open Session("ConnectStr")
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	   Set webdbCommand = Server.CreateObject("ADODB.Command")
	   Set webdbCommand.ActiveConnection = webdb

	   dim maxrow
  	   dim rowcount
	   dim batch
	   
	   maxrow = Request.Form("txtRowNo")
	   batch = Request.Form("txtBatch")
	   
	   if batch = "" then
			batch = 0
	   end if
	   
	   if isnumeric(maxrow) then
			ssql = "Exec sp_Wtd_insUpdDelEmpTrainEvaluation '" & Session("Regisno") & "', '" & Request.form("CboEmpID") & "' , '" _
	      	            & Request.form("cboCourse") & "', '" & Request.form("cboPeriod") & "'," & batch & ", '" _
	      	            & Request.form("cboQGRP") & "',0,0,'','DEL'"
			 webdbCommand.CommandText = ssql
			 webdb.Execute webdbCommand.CommandText	
			 
          do until rowcount = cint(maxrow)
	         rowcount = rowcount + 1
			 
			 ssql = "Exec sp_Wtd_insUpdDelEmpTrainEvaluation '" & Session("Regisno") & "', '" & Request.form("CboEmpID") & "' , '" _
	      	            & Request.form("cboCourse") & "', '" & Request.form("cboPeriod") & "'," & batch & ", '" _
	      	            & Request.form("cboQGRP") & "'," & request.Form("txtQID" + cstr(rowcount)) & "," _
	      	            & request.Form("cboAnswer" + cstr(rowcount)) & ",'','Add'"
			 webdbCommand.CommandText = ssql
			 webdb.Execute webdbCommand.CommandText
		  loop
      end if
      response.redirect "TD_TranEvalution.asp"
end if
%>
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
	function Change()
		document.frmTranEvalution.txtaction.value = ""
		document.frmTranEvalution.submit()
	end function
	
	function validate()
	    document.frmTranEvalution.txtaction.value = "Upd"
		document.frmTranEvalution.submit()
	end function
	
	function chkUpd()
		dim maxrow
		dim rowcount
		dim blank
		dim update
		
		blank = ""
		update = false
		maxrow = document.frmTranEvalution.txtRowNo.value
		if isnumeric(maxrow) then
			do until rowcount = cint(maxrow) 
				rowcount = rowcount + 1
				ssql="if " + "document.frmTranEvalution.cboAnswer" + cstr(rowcount) + ".value = blank then" + chr(10) 
				ssql= ssql + " update = true" + chr(10) 
				ssql=ssql + "end if"
  				execute ssql
			loop
		end if
		
		if update = true then
			document.frmTranEvalution.cmdUpdate.disabled = true
		else
			document.frmTranEvalution.cmdUpdate.disabled = false
		end if
	end function
	
	// -->
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
      <P><IMG height=84 src="../Image/engtmsatt.gif" width=683 border=0 ><BR><BR>
      <FORM name=frmTranEvalution action=TD_TranEvalution.asp method=post>
      <TABLE cellSpacing=0 width="100%" border=0>
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <TD>
           <FONT class=small><b>Employee ID &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
            <select name=cboempid style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px"  onchange=Change()>  
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
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <TD>
           <FONT class=small><b>Course Name&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
            <select name=cboCourse style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px"  onchange=Change()>  
       			<%  dim tmpCourseID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WTD_SelTranTrain '" + Session("Regisno") + "','" + tmpEmpID + "','','',0,'Course'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpCourseID = ""
	  			tmpCourseID = Request.form("cboCourse")
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("courseid")) = Request.form("cboCourse") ) or ( trim(webdbRecordset.Fields("courseid")) = tmpCourseID )then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("courseid")) + ">"  + " " + trim(webdbRecordset.Fields("description")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("courseid")) + ">"  + " " + trim(webdbRecordset.Fields("description")) + "</option>"
 				    end if
 				    
 				    if tmpCourseID = "" then
					      tmpCourseID = trim(webdbRecordset.Fields("courseid"))
					      
					end if   
				   webdbRecordset.MoveNext  
		        loop       
			%></select> 
           </TD>
           <TD>
           <FONT class=small><b>Period &nbsp;&nbsp;&nbsp;</b> </FONT>
            <select name=cboPeriod style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 200px"  onchange=Change()>  
       			<%  dim tmpPeriodID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WTD_SelTranTrain '" + Session("Regisno") + "','" + tmpEmpID + "','" + tmpCourseID + "','',0,'Period'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpPeriodID = ""
	  			tmpPeriodID = Request.form("cboPeriod")
			 	
			 	if webdbrecordset.eof then
			 		tmpPeriodID = ""
			 	end if
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("dateStart")) + " - " + trim(webdbRecordset.Fields("dateend")) = Request.form("cboPeriod") ) or ( trim(webdbRecordset.Fields("dateStart")) + " - " + trim(webdbRecordset.Fields("dateend")) = tmpPeriodID )then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("dateStart")) + " - " + trim(webdbRecordset.Fields("dateend")) + ">"  + " " + trim(webdbRecordset.Fields("dateStart")) + " - " + trim(webdbRecordset.Fields("dateend")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("dateStart")) + " - " + trim(webdbRecordset.Fields("dateend")) + ">"  + " " + trim(webdbRecordset.Fields("dateStart")) + " - " + trim(webdbRecordset.Fields("dateend")) + "</option>"
 				    end if
 				    
 				    if tmpPeriodID = "" then
					      tmpPeriodID = trim(webdbRecordset.Fields("dateStart")) + " - " + trim(webdbRecordset.Fields("dateend"))
					      
					end if   
				   webdbRecordset.MoveNext  
		        loop       
			%></select>&nbsp;&nbsp;&nbsp; 
           </TD>
        </TR>
       <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <TD>
           <FONT class=small><b>Question Group &nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
            <select name=cboQGRP style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px"  onchange=Change()>  
       			<%  dim tmpQGRP

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WTD_SelTranTrain '" + Session("Regisno") + "','" + tmpEmpID + "','" + tmpCourseID + "','',0,'QuestionGrp'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpQGRP = ""
	  			tmpQGRP = Request.form("cboQGRP")
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("groupID")) = Request.form("cboQGRP") ) or ( trim(webdbRecordset.Fields("groupID")) = tmpQGRP )then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("groupID")) + ">"  + " " + trim(webdbRecordset.Fields("groupdesc")) +  "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("groupID")) + ">"  + " " + trim(webdbRecordset.Fields("groupdesc")) +  "</option>"
 				    end if
 				    
 				    if tmpQGRP = "" then
					      tmpQGRP = trim(webdbRecordset.Fields("groupID"))
					      
					end if   
				   webdbRecordset.MoveNext  
		        loop       
			%></select>&nbsp;&nbsp;&nbsp;
           </TD>
           <TD>
           <FONT class=small><b>Batch &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
       			<%  dim tmpBatchID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WTD_SelTranTrain '" + Session("Regisno") + "','" + tmpEmpID + "','" + tmpCourseID + "','" + trim(left(tmpPeriodID,10)) + "',0,'Batch'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpBatchID = "0"
	  			tmpBatchID = Request.form("CboBatch")
		        
		        if webdbrecordset.eof then
			%><input type=text name=txtbatch readonly value=>&nbsp;&nbsp;&nbsp; 
			<% else %>
			<input type=text name=txtbatch readonly value=<%=trim(webdbRecordset.Fields("batchid"))%>>&nbsp;&nbsp;&nbsp; 
			<%tmpBatchID = trim(webdbRecordset.Fields("batchid"))
			end if %>
           </TD>
        </TR> 
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
      <td bgcolor="#f3f3f3" width="800"><font class="marineblack">Question</font></td>
      <td bgcolor="#f3f3f3" width="200"><font class="marineblack"><center>Answer</center></font></td>
    </tr>

            <%     page_size =10
					dim answer
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
		            if tmpBatchID = "" then
						tmpBatchID = "0"
					end if
					
		            if tmpPeriodID <> "" then
					     ssql = "Exec sp_WTD_SelEmpTrainEvaluation '" + Session("Regisno") + "','" + tmpEmpID + "','" + tmpCourseID + "','" + trim(left(tmpPeriodID,10)) + "'," + trim(tmpBatchID) + ",'"+ tmpQGRP + "',0,'QUESTION'"
					else
						 ssql = "Exec sp_WTD_SelEmpTrainEvaluation '" + Session("Regisno") + "','" + tmpEmpID + "','" + tmpCourseID + "','" + trim(left(tmpPeriodID,10)) + "'," + trim(tmpBatchID) + ",'"+ tmpQGRP + "',0,'QUESTION2'"
					end if
					
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
					
			        page_count = rs.pagecount
			        rowno = 0
			        if 1 > current_page then current_page = 1
			        if current_page > page_count then current_page = page_count
					
					if rs.RecordCount >0 then
			        rs.AbsolutePage = current_page
			        		       
					colour = 0
					 
			        do while rs.AbsolutePage = current_page and not rs.EOF
			     
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if

			            rowno = rowno + 1
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20' WIDTH=800px " + colour + "><font class='small'> " + rs("description") + "<input type=hidden name=txtQID" + cstr(rowno)+ " value=" + cstr(rs("questionid")) + "></td>"
				        
				        Set webdb = Server.CreateObject("ADODB.Connection")
						webdb.Open Session("ConnectStr")
						Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
						Set webdbCommand = Server.CreateObject("ADODB.Command")
						ssql ="Exec sp_WTD_SelEmpTrainEvaluation '" + Session("Regisno") + "','" + tmpEmpID + "','" + tmpCourseID + "','" + trim(left(tmpPeriodID,10)) + "'," + trim(tmpBatchID) + ",'"+ tmpQGRP + "'," + cstr(rs("questionid")) + ",'ANSWER'"
			  			Set webdbCommand.ActiveConnection = webdb
			  			webdbCommand.CommandText = ssql
			  			webdbRecordset.Open webdbCommand,,1 , 3
			  			
			  			response.write "<td height='20'  " + colour + "><font class='small'>"
			  			
			  			response.Write "<select name=cboAnswer" + cstr(rowno) + " style='HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 200px' onchange=ChkUpd()>"
			  			Do Until webdbRecordset.EOF
							answer = rs("answerdescid")
 							if trim(webdbRecordset.Fields("answerdescid")) = rs("answerdescid") then
   			  					response.write "<option selected value=" + trim(webdbRecordset.Fields("answerdescid")) + ">"  + " " + trim(webdbRecordset.Fields("answerdesc")) + "</option>"
							else 
   			  					response.write "<option value=" + trim(webdbRecordset.Fields("answerdescid")) + ">"  + " " + trim(webdbRecordset.Fields("answerdesc")) + "</option>"
 							end if
		 				    
						webdbRecordset.MoveNext  
						loop 
				        Response.Write "</select></td></tr>"
				       
				        rs.Movenext  
				        count = abs(count - 1)        
			        loop
			        response.Write "<tr><td bgcolor='#ffffff' width='30'>&nbsp;</td>"
					response.Write "<td></td>"
					response.Write "</tr>"
					response.Write "<tr>"
					response.Write "<td bgcolor='#ffffff' width='30'>&nbsp;</td>"
					
					if answer <> "" then
					response.Write "<td bgcolor='#ffffff' width='30'><input type=button name=cmdUpdate onclick=validate() value=Update size=50></td>"
					else
					response.Write "<td bgcolor='#ffffff' width='30'><input type=button name=cmdUpdate onclick=validate() value=Update size=50 disabled></td>"
					end if
					response.write "<td><input type=hidden name=txtRowNo value=" + cstr(rowno) + "><input type=hidden name=txtAction value=></td>"
					response.Write "</tr>"
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
			      else
					response.Write "<tr><td bgcolor='#ffffff' width='30'>&nbsp;</td>"
					response.Write "<td></td>"
					response.Write "</tr>"
					response.Write "<tr>"
					response.Write "<td bgcolor='#ffffff' width='30'>&nbsp;</td>"
					response.Write "<td bgcolor='#ffffff' width='30'><input type=button name=cmdUpdate onclick=validate() value=Update size=50 disabled></td>"
					response.write "<td><input type=hidden name=txtRowNo value=" + cstr(rowno) + "><input type=hidden name=txtAction value=></td>"
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
				Response.Write "<a href=""TD_TranEvalution.asp?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""TD_TranEvalution.asp?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""TD_TranEvalution.asp?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""TD_TranEvalution.asp?page="
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
    <center>  </form>
    <TD align=center colspan=2 width="936" height="40" class="small"><br>
      <!--&nbsp;<br>
      &nbsp;<BR>--><font class ="small" >Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>.</font></TD></TR></TBODY></TABLE></center>
      </div>
</BODY>




