<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/AdoVbs.asp"-->
<!-- #include virtual ="/ehres/global/inputSession.asp"-->
<%
	if Request("txtAction")="ADD" then
	
	   Set webdb = Server.CreateObject("ADODB.Connection")
	   webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	   Set webdbCommand = Server.CreateObject("ADODB.Command")
	   Set webdbCommand.ActiveConnection = webdb


	    ssql = "Exec sp_WRQ_insDelRQLabel '" & request.Form("CboType") & "', '" & Request.Form("ID") & "' , '" _
	      	    & Request.Form("Description") & "', 'ADD'"
		webdbCommand.CommandText = ssql
		webdb.Execute webdbCommand.CommandText
		
     response.redirect "ADDLabelDetail.asp"
	else 
		if Request("txtAction")="BACK" then
			response.redirect "DELLabelDetail.asp"
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
	function Change()
		document.frmADDLabel.txtaction.value = ""
		document.frmADDLabel.submit()
	end function
	
	function ADDNew()
		document.frmADDLabel.txtaction.value = "ADD"
		document.frmADDLabel.submit()
	end function
	
	function Back()
		document.frmADDLabel.txtaction.value = "BACK"
		document.frmADDLabel.submit()
	end function
	
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
      <FORM name=frmADDLabel action=ADDLabelDetail.asp method=post>
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
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="1">
        <TBODY>
    <tr>
              <td height="20" align="center" width="7%"><font class="marineblack"> </font></td>            
              <td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Label ID</font></td>
              <td height="20" width="73%" bgcolor="#FFFFFF"><input type=text name=ID size=60></td>
            </tr>
			<tr>
              <td height="20" align="center" width="7%"><font class="marineblack"> </font></td>            
              <td height="20" width="20%" bgcolor="#F3F3F3"><font class="marineblack">Label Description</font></td>
              <td height="20" width="73%" bgcolor="#FFFFFF"><input type=text name=description size=60><input type=hidden name=txtAction value= ></td>
            </tr>
            <tr>
            <td height="20" align="center" width="7%"><font class="marineblack"> </font></td> 
          <td width="20%" height="19"><input type=button value="Update" name=cmdUpdate onclick="ADDNew()" class="small"><input type="button" value="Back" name="cmdBack" onclick="Back()" class="small"></td>
          <td width="73%" height="19"></td>
        </tr>
            </TBODY>
            </TABLE>
    </TD></TR>
    <center>  </form>
    <TD align=center colspan=2 width="936" height="40" class="small"><br>
      <!--&nbsp;<br>
      &nbsp;<BR>--><font class ="small" >Copyright © 1997-2006 SoftFac
      Technology Sdn Bhd <i>All Rights Reserved</i>.</font></TD></TR></TBODY></TABLE></center>
      </div>
</BODY>




