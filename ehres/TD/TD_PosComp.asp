<!-- #INCLUDE FILE = "../global/ConnectStr.asp"-->
<!-- #INCLUDE FILE = "../global/AdoVbs.asp"-->
<!-- #INCLUDE FILE = "../global/inputSession.asp"-->
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

<% dim connect_string
       connect_string =Session("ConnectStr")
%>
       
	<SCRIPT LANGUAGE="JavaScript">

	function Change()
	{	
		document.forms[0].submit();
	}
	
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
      <P><IMG height=84 src="../Image/engposcom.gif" width=683 border=0 ><BR><BR>
        <FORM name=frmGrpPosComp action=TD_PosComp.asp method=post>
      <TABLE cellSpacing=0 width="100%" border=0>
        <TR>
           <TD WIDTH=10%>&nbsp;</TD>
           <TD>
           <FONT class=small><b>Position&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b> </FONT>
            <select name=cboempid style="HEIGHT: 22px; LEFT:83px; TOP: 8px; WIDTH: 400px"  onchange=Change()>  
       			<%  dim tmpEmpID

       			Set webdb = Server.CreateObject("ADODB.Connection")
				webdb.Open Session("ConnectStr")
			    Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
			    Set webdbCommand = Server.CreateObject("ADODB.Command")
                ssql ="Exec sp_WTD_SelPosition '" + Session("Regisno") + "','" + session("EmpID") + "'"
			  	Set webdbCommand.ActiveConnection = webdb
			  	webdbCommand.CommandText = ssql
			  	webdbRecordset.Open webdbCommand,,1 , 3
	  			
	  			tmpEmpID = ""
	  			txtdate1 = ""
	  			txtdate2 = ""
	  			tmpEmpID = Request.form("cboempid")
			 	 
			  	Do Until webdbRecordset.EOF
                    
 					if ( trim(webdbRecordset.Fields("PositionID")) = Request.form("cboempid") ) or ( trim(webdbRecordset.Fields("PositionID")) = tmpEmpID )then
   			  	        response.write "<option selected value=" + trim(webdbRecordset.Fields("PositionID")) + ">"  + " " + trim(webdbRecordset.Fields("description")) + "</option>"
					else 
   			  	        response.write "<option value=" + trim(webdbRecordset.Fields("PositionID")) + ">"  + " " + trim(webdbRecordset.Fields("description")) + "</option>"
 				    end if
 				    
 				    if tmpEmpID = "" then
					      tmpEmpID = trim(webdbRecordset.Fields("PositionID"))
					      
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
    <TD vAlign=top align=middle colspan=2 width="936" height="193">
      <div align="center">
        <center>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="1">
        <TBODY>
    <tr>
      <td bgcolor="#ffffff" width="30">&nbsp;</td>
      <td bgcolor="#f3f3f3" width="80"><font class="marineblack">Date</font></td>
      <td bgcolor="#f3f3f3" width="240"><font class="marineblack">Competency Structure</font></td>
      <td bgcolor="#f3f3f3" width="50"><font class="marineblack">Score</font></td>
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
		            
                    if Request.Form("cboempid") ="ALL" or tmpEmpID ="ALL" then                
						 ssql = "Exec sp_WTD_SelGrpPosComp '" + Session("Regisno") + "','" + tmpEmpID + "','ALL'"		
					else	
					     ssql = "Exec sp_WTD_SelGrpPosComp '" + Session("Regisno") + "','" + tmpEmpID + "','One'"
					end if
					
                    rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText 
						 
			        page_count = rs.pagecount
			        
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
			        
				        response.write "<tr>"
				        response.write "<td></td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("pos") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("comp") + "</td>"
				        response.write "<td height='20'  " + colour + "><font class='small'>" + rs("Score") + "</td>"
				        Response.Write "</tr>"
				       
				        rs.Movenext  
				        count = abs(count - 1)        
			        loop
			        
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
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
				Response.Write "<a href=""TD_PosComp.asp?page=1"
				Response.Write """ ><font face=""Verdana"" & size=""1""><< First</font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page >= page_count then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">Next ></font></a>" & "<font=""2""> | </font>"
			end if
  
			if current_page < page_count then
				Response.Write "<a href=""TD_PosComp.asp?page="
				Response.Write current_page + 1
				Response.Write """ ><font face=""Verdana"" & size=""1"">Next ></font></a>" & "<font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page <> 1 then
				Response.Write "<a href=""TD_PosComp.asp?page="
				Response.Write current_page - 1
				Response.Write """><font face=""Verdana"" & size=""1"">< Previous </font></a><font size=""2""> |</font>" & vbCrlf
			end if
  
			if current_page = 1 then
				Response.Write"<font face=""Verdana"" & color =""silver"" & size=""1"">" & "< Previous </font><font size=""""> | </font>"
			end if				
 
			if current_page <> page_count then
				Response.Write "<a href=""TD_PosComp.asp?page="
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




