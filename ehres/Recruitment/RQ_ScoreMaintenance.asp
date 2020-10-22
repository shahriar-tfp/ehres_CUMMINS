<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/AdoVbs.asp"-->
<!-- #include virtual ="/ehres/global/inputSession.asp"-->
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
	
	function Change()
		document.frmEmpScore.action = "RQ_ScoreMaintenance.asp"
		document.frmEmpScore.submit()
	End function
	
	function Finish()
		document.frmEmpScore.action = "../Maintenance.asp"
		document.frmEmpScore.submit()
	End function
	
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
    <TD vAlign=top align=middle width="27" height="109"></TD>
    <TD vAlign=top align=middle width="907" height="109">
      <P><IMG height=84 src="../Image/engrecruitment.gif" width=683 border=0 ><BR><BR>
        <FORM name=frmEmpScore action=UpdateQAAnswerScore.asp method=post>
      <TABLE cellSpacing=0 width="100%" border=0>
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <td>
           <FONT class=small><b>Recruitment ID &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
            <select name=cboRQID style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px"  onchange=Change() ID="Select1">  
       			<%  dim tmpRQID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="select RQID from rq_Profile group by RQID order by RQID"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpRQID = ""
	  			tmpRQID = Request.form("cboRQID")
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("RQID")) = Request.form("cboRQID") ) or ( trim(webdbRecordset.Fields("RQID")) = tmpRQID )then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("RQID")) + ">"  + " " + trim(webdbRecordset.Fields("RQID")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("RQID")) + ">"  + " " + trim(webdbRecordset.Fields("RQID")) + "</option>"
 				    end if
 				    
 				    if tmpRQID = "" then
					      tmpRQID = trim(webdbRecordset.Fields("RQID"))
					      
					end if   
				   webdbRecordset.MoveNext  
		        loop       
			%></select>&nbsp;&nbsp;&nbsp; 
           </TD>
        </TR><tr>
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
      <td bgcolor="#f3f3f3" width="500"><font class="marineblack">Question</font></td>
      <td bgcolor="#f3f3f3" width="150"><font class="marineblack">Max Score</font></td>
      <td bgcolor="#f3f3f3" width="150"><font class="marineblack">Answer</font></td>
      <td bgcolor="#f3f3f3" width="150"><font class="marineblack">Score</font></td>
    </tr>

            <%     
                    set myconn = server.CreateObject("ADODB.Connection")
			        set rs = server.CreateObject("ADODB.Recordset")
		                myconn.open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
		                 
					ssql = "select a.QID,a.description,a.maxscore,isnull(b.answer,'') as answer, isnull(b.score,0) as score from rq_QandA a left outer join rq_QAScore b on a.QID = b.QID and b.RQID = '" + tmpRQID + "' order by a.QID"
					rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
						
					colour = 0
					rowno = 0 
			        do while not rs.EOF
			        
				        if count = 1 then
				           colour = " bgcolor='#eeeeee'"
				        else
				           colour = ""
				        end if
			            rowno = rowno + 1
			           
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>"+ rs("description") + "<input type=hidden name=EVa" + cstr(rowno) + " value='" + rs("QID") + "'></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + cstr(rs("maxscore")) + "<input type=hidden name=max" + cstr(rowno) + " value =" + cstr(rs("maxscore")) + "></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=text name=S" + cstr(rowno) +" value='"+rs("Answer")+"'></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'><input type=text name=M" + cstr(rowno) +" value='"+cstr(rs("Score"))+"'></td>"
				        Response.Write "</tr>"
				       
				        rs.Movenext        
			        loop
			        
				    response.write "<tr>"
				    response.write "<td width='10%'>&nbsp;</td>"
				    response.Write "</tr>"
				    
			        response.write "<tr>"
				    response.write "<td width='10%'></td>"
			        response.Write "<td bgcolor='#ffffff'><input type=button name=cmdUpdate onclick=validate() value=Update><input type=button name=cmdFinish onclick=Finish() value=Back></td>"
			        response.write "<td><input type=hidden name=txtRowNo value=" + cstr(rowno) + "><input type=hidden name=txtAction value=></td>"
			        response.Write "</tr>"
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
			       
			%>        
			
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




