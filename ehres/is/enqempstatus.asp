<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<%
Response.Buffer = true
%>

<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<LINK href="../css/login.css" type=text/css rel=stylesheet>
<title>Employee Status</title>
</head>

<body bgcolor="#ffffff">

<table width="100%" border="0" style="BORDER-COLLAPSE: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%">
    <p align="center">
    <IMG src="../Image/engisstatus.gif" border=0></p></td>
  </tr>
</table>
<table border="0" width="100%" bordercolor="#111111" cellpadding="0" cellspacing="0">
  <tr>
    
      <td align="left" width="100%" colSpan="2" height="12%">
      <font class="marineblack" >
      <P>
      <BR><BR>
      <TABLE width="100%" border=0>
<%

      dim ssql 
        
		Set webdb = Server.CreateObject("ADODB.Connection")
			webdb.Open Session("ConnectStr")
		Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
		Set webdbCommand = Server.CreateObject("ADODB.Command")

		If Request("empid") <> "" Then
		   ssql = "Exec sp_wis_selconfimemployeedetail '" & Request("regisno") & "','" & Request("empid") &"','PERSONAL'"

		   Set webdbCommand.ActiveConnection = webdb
			   webdbCommand.CommandText = ssql
			   webdbRecordset.Open webdbCommand, , 1  , 3 

           Response.Write "<TR>"
           Response.Write "<TD class='marineblack' width='15%'>Employee ID</TD>"
           Response.Write "<TD class='marineblack'>" + webdbRecordset.fields(0) + "</TD>"
           Response.Write "<TD class='marineblack' width='15%'>Name</TD>"
           Response.Write "<TD class='marineblack'>" + webdbRecordset.fields(1) + "</TD></TR>"
           Response.Write "<TR>"
           Response.Write "<TD class='marineblack' width='15%'>Division</TD>"
           Response.Write "<TD class='marineblack'>" + webdbRecordset.fields(2) + "</TD>"
           Response.Write "<TD class='marineblack' width='15%'>Department</TD>"
           Response.Write "<TD class='marineblack'>" + webdbRecordset.fields(3) + "</TD></TR>"
           Response.Write "<TR>"
           Response.Write "<TD class='marineblack' width='15%'>Job Grade</TD>"
           Response.Write "<TD class='marineblack'>" + webdbRecordset.fields(4) + "</TD>"
           Response.Write "<TD class='marineblack' width='15%'>Job Title</TD>"
           Response.Write "<TD class='marineblack'>" + webdbRecordset.fields(5) + "</TD>"
           Response.Write "</TR>"		   
		   
		End if

%>      

      </TABLE>        
      </font></P>
      </td>
    
  </tr></table>
<table width="100%" border="0" style="BORDER-COLLAPSE: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0">
  <tr>
    <td width="95%">
    <table borderColor="#808080" cellSpacing="0" cellPadding="0" width="100%" border="0">

      <tr>
        <td width="15%" height="20">
        <font class="marineblack" size="1" >&nbsp;&nbsp;&nbsp;</font></td>
        <td width="5%"  height="20">
        <font class="marineblack" size="1" >&nbsp;&nbsp;&nbsp;</font></td>
      </tr>

      <tr>
        <td width="15%" bgColor="#f3f3f3" height="20">
        <font class="marineblack" size="1" ><b>Status</b></font></td>
        <td width="5%" bgColor="#f3f3f3" height="20">
        <font class="marineblack" size="1" ><b>Date</b></font></td>
      </tr>

    <% 

        dim count, colour
         
		Set webdb = Server.CreateObject("ADODB.Connection")
			webdb.Open Session("ConnectStr")
		Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
		Set webdbCommand = Server.CreateObject("ADODB.Command")

		If Request("empid") <> "" Then
           ssql = "Exec sp_wis_selconfimemployeedetail '" & Request("regisno") & "','" & Request("empid") &"','DETAIL'"

		   Set webdbCommand.ActiveConnection = webdb
			   webdbCommand.CommandText = ssql
			   webdbRecordset.Open webdbCommand, , 1  , 3 
			   
           colour = 0
	       Do Until webdbRecordset.EOF 
	 	      if count = 1 then
				 colour = " bgcolor='#eeeeee'"
			  else
			     colour = ""
			  end if

              Response.Write "<TR>"
              Response.Write "<TD height='25' " + colour + "><font class=small>" + webdbRecordset.fields(0) + "</Font></TD>"
              Response.Write "<TD height='25' " + colour + "><font class=small>" + webdbRecordset.fields(1) + "</Font></TD>"
              Response.Write "</TR>"		   

              webdbRecordset.MoveNext
              count = abs(count - 1)
              
           loop   
		   
		End if
       
    %>
    </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" height="100" style="BORDER-COLLAPSE: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0">
  <tr>
    <td align="middle" width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td class="small" align="middle" width="100%">
    <img src="http://10.10.10.4/eHres/Image/dottedlinenav.gif" border="0" width="408" height="4"><font face="Verdana" size="1">&nbsp;<br>
    &nbsp;</font><p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </td>
  </tr>
</table>

<p>&nbsp;</p>

</body>

</html>