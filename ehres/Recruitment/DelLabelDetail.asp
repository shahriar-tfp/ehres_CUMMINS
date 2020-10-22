<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/AdoVbs.asp"-->
<!-- #include virtual ="/ehres/global/inputSession.asp"-->
<%
	if Request("txtAction")="DEL" then
	
	   Set webdb = Server.CreateObject("ADODB.Connection")
	   webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	   Set webdbCommand = Server.CreateObject("ADODB.Command")
	   Set webdbCommand.ActiveConnection = webdb

	   dim maxrow
  	   dim rowcount
	
	   maxrow = Request.Form("txtRowNo")
	   del ="false"
    
       if isnumeric(maxrow) then	
          do until rowcount = cint(maxrow)
	         rowcount = rowcount + 1
   	         if Request.Form("D" + cstr(rowcount)) <> "0" then
			    del = "true"
	      	    ssql = "Exec sp_WRQ_insDelRQLabel '" & request.Form("CboType") & "', '" & Request.Form("D"+ cstr(rowcount)) & "' , '" _
	      	            & Request.Form("L"+ cstr(rowcount)) & "', 'DEL'"
			    webdbCommand.CommandText = ssql
			    webdb.Execute webdbCommand.CommandText
	         end if	
         loop
      end if
      
      if del = "true" then
         response.redirect "DelLabelDetail.asp"
      end if

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
	function Back()
		document.frmDelLabel.action = "../Maintenance.asp"
		document.frmDelLabel.submit()
	end function
	
	function Change()
		document.frmDelLabel.txtaction.value = ""
		document.frmDelLabel.submit()
	end function
	
	function ADDNew()
		document.frmDelLabel.txtaction.value = ""
		document.frmDelLabel.action = "ADDLabelDetail.asp"
		document.frmDelLabel.submit()
	end function
	
	function ValidateDelData() 
	
	dim rowcount 
	dim ssql
	dim maxrow
	dim action
	dim del
	
	document.frmDelLabel.txtaction.value = "DEL"
	maxrow = document.frmDelLabel.txtRowNo.value	
	action = ""
    if isnumeric(maxrow) then	
	   do until rowcount = cint(maxrow)
	      rowcount = rowcount + 1
	      ssql="if " + "document.frmDelLabel.C" + cstr(rowcount) + ".checked = false then" + chr(10) 
	      ssql= ssql + " document.frmDelLabel.D" + cstr(rowcount) + ".value=0" + chr(10) 
	      ssql=ssql + "end if"
	
  	      execute ssql
	   loop
	   document.frmDelLabel.submit()
	end if
End function	
	
	// -->
	</SCRIPT>

<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div align="center">
  <center>
<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%" height="392">
  <TBODY>
  <TR>
    <TD vAlign=top align=middle width="27" height="109"></TD>
    <TD vAlign=top align=middle width="907" height="109">
      <P><IMG height=84 src="../Image/engRecruitment.gif" width=683 border=0 ><BR><BR>
      <FORM name=frmDelLabel action=DelLabelDetail.asp method=post>
      <TABLE cellSpacing=0 width="100%" border=0>
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <TD>
           <FONT class=small><b>Type &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
            <select name=cboType style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px"  onchange=Change()>  
       			<%  dim tmpType

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WRQ_selTypeInfo "
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpType = ""
	  			tmpType = Request.form("cboType")
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("TypeID")) = Request.form("cboType") ) or ( trim(webdbRecordset.Fields("TypeID")) = tmpType )then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("TypeID")) + ">"  + " " + trim(webdbRecordset.Fields("Description")) + " " + "-" + " " + trim(webdbRecordset.Fields("TypeID")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("TypeID")) + ">"  + " " + trim(webdbRecordset.Fields("Description")) + " " + "-" + " " + trim(webdbRecordset.Fields("TypeID")) + "</option>"
 				    end if
 				    
 				    if tmpType = "" then
					      tmpType = trim(webdbRecordset.Fields("TypeID"))
					      
					end if   
				   webdbRecordset.MoveNext  
		        loop       
			%></select>&nbsp;&nbsp;&nbsp; 
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
              <td height="20" align="center" width="7%"><font class="marineblack"> </font></td>            
              <td height="20" align="center" width="10%" bgcolor="#F3F3F3"><font class="marineblack">Delete</font></td>
              <td height="20" width="30%" bgcolor="#F3F3F3"><font class="marineblack">Label ID</font></td>
              <td height="20" width="53%" bgcolor="#F3F3F3"><font class="marineblack">Label Description</font></td>
            </tr>

            <%   
            	   Set webdb = Server.CreateObject("ADODB.Connection")
   		               webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
  		           Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  		           Set webdbCommand = Server.CreateObject("ADODB.Command")

		           ssql = "Exec sp_WRQ_SelrqLabel """ + tmpType + """"
					
				   'Response.Write ssql
			        Set webdbCommand.ActiveConnection = webdb
			            webdbCommand.CommandText = ssql
			            webdbRecordset.Open webdbCommand,,1 , 3

					 colour = 0
					rowno = 0
			        Do Until webdbRecordset.EOF
					    rowno = rowno + 1			        
					 
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			        
				        response.write "<tr>"
						response.write "<td> </td>" 				        
						response.write "<td align='center'" + colour + "><font class='small'><input type='checkbox' name=C" + cstr(rowno) + " value='ON' " + strcheck + "></font></td>" 
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("ID") + "<input type='hidden' name=D" + cstr(rowno) + " value= " + webdbRecordset.Fields("ID") + "></td>"
				        response.write "<td align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("Description") + "<input type='hidden' name=L" + cstr(rowno) + " value= " + webdbRecordset.Fields("Description") + "></td></tr>"
				        webdbRecordset.MoveNext  
				        count = abs(count - 1)   
			        loop
	  				 response.write "<input type=hidden name=txtRowNo value=" + cstr(rowno) + "><input type=hidden name=txtAction value=>"

			        webdbRecordset.close
			        webdb.close      
			 %>
            </TBODY>
            </TABLE></center>
      </div>
    </TD></TR>
    <tr>
          <td height="19"></td>
        </tr>
        <tr>
          <td width="6%" height="19"></td>
          <td width="94%" height="19"><input type=button value="ADD New" name=cmdADD onclick="ADDNew()" class="small" ID="Button1"><input type="submit" value="Delete" name="cmdDelete" onclick="ValidateDelData()" class="small"><input type="button" value="Back" name="cmdBack" onclick="Back()" class="small"></td>
        </tr>
    <center>  </form>
    <TD align=center colspan=2 width="936" height="40" class="small"><br>
      <!--&nbsp;<br>
      &nbsp;<BR>--><font class ="small" >Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>.</font></TD></TR></TBODY></TABLE></center>
      </div>
</BODY>




