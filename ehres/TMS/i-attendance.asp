<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

	<SCRIPT LANGUAGE="JavaScript">

	function Verify()
	{
		msg = "";
		m = true;
		n = true;

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
	
	</SCRIPT>


<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div align="center">
  <center>
    <TABLE cellSpacing=0 cellPadding=0 border=0 width="100%" height="392">
      <!--DWLayoutTable-->
      <TBODY>
        <TR> 
          <TD height="29" colspan="3" align=middle vAlign=top bgcolor="#0099cc"> 
            <div align="center"> 
              <center>
                <table border="0" width="100%">
                  <tr> 
                    <td width="3%"> </td>
                    <td width="23%"> <font class="marinewhite"> Employee ID : 
                      <%    
          response.write session("EmpID")
       %>
                      </font> </td>
                    <td width="74%"><font class="marinewhite"> Name : 
                      <%    
          response.write session("EmpName")
       %>
                      </font> </td>
                  </tr>
                </table>
              </center>
            </div></TD>
        </TR>
        <TR> 
          <TD height="21" colspan="3" align="middle" vAlign=top class="small"> 
            <p align="right"><A href="../main.asp"><font color="#000000">Home</font></A>&nbsp;&nbsp;&nbsp; 
              |&nbsp;&nbsp;&nbsp; <A href="../signout.asp"><font color="#000000">Logout</font></A></p>
            <form name="form1" method="post" action="attendance.asp">
              <table width="75%" border="1">
                <!--DWLayoutTable-->
                <tr> 
                  <td width="200" height="22" valign="top"><font size="2">Date</font></td>
                  <td width="127" valign="top"><font size="2">Date In</font></td>
                  <td width="103" valign="top"><font size="2">Time In</font></td>
                  <td width="142" valign="top"><font size="2">Date Out</font></td>
                  <td width="138" valign="top"><font size="2">Time Out</font></td>
                </tr>
                <tr> 
                  <td height="26" valign="top"><font size="2" face="Arial, Helvetica, sans-serif">Select 
                    date and Time </font></td>
                  <td valign="top"><div align="right"> 
                      <input type="text" name="txtdatein" size='15'>
                    </div></td>
                  <td valign="top"><div align="right"> 
                      <input type="text" name="txttimein" size='15'>
                    </div></td>
                  <td valign="top"><div align="right"> 
                      <input name="txtdateout" type="text" id="txtdateout" size='15'>
                    </div></td>
                  <td valign="top"><div align="right"> 
                      <input name="txttimeout" type="text" id="txttimeout" size='15'>
                    </div></td>
                </tr>
              </table>
              <br>
              <input type="submit" name="Submit" value="Submit">
            </form>
            <p align="left"><font color="#000000"></font></p></TD>
        </TR>
        <TR> 
          <TD width="18" height="33">&nbsp;</TD>
          <TD width="962">&nbsp;</TD>
          <TD width="12">&nbsp;</TD>
        </TR>
        <TR> 
          <TD height="249">&nbsp;</TD>
          <TD vAlign=top>
	
		  <table width="98%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
			      <td align="right" bgcolor="#F3F3F3" width="75" height="30"><div align="center"><font class="marineblack">Date 
                    In</font></div></td>
                <td align="right" bgcolor="#F3F3F3" width="75" height="30"><div align="center"><font class="marineblack">Date 
                    In</font></div></td>
                <td align="right" bgcolor="#F3F3F3" width="75" height="30"><div align="center"><font class="marineblack">Time 
                    In</font></div></td>
                <td align="right" bgcolor="#F3F3F3" width="75" height="30"><div align="center"><font class="marineblack">Date 
                    Out</font></div></td>
                <td align="right" bgcolor="#F3F3F3" width="75" height="30"><div align="center"><font class="marineblack">Time 
                    Out</font></div></td>
              </tr>
              <%    

dim colour
dim count
dim id



     Set webdb = Server.CreateObject("ADODB.Connection")
         webdb.Open Session("ConnectStr")
	 Set webdbRecordset = Server.CreateObject("ADODB.Recordset") ''" + Session("empid") + "'
	 Set webdbCommand = Server.CreateObject("ADODB.Command")

    ssql = "Exec sp_web_getsheetdata  '',""" + Session("EmpID") + """, '', '', '','', 'RETRIEVE'"
    'ssql = "select datein,datein, timein, dateout, timeout from  etimesheet "
    'Response.write ssql
   
	 Set webdbCommand.ActiveConnection = webdb
	     webdbCommand.CommandText = ssql
	     webdbRecordset.Open webdbCommand,,1 , 3
	  'Response.Write ssql
     colour = 0
     
     Do Until webdbRecordset.EOF

	 
          response.write "<TABLE cellSpacing=0 cellPadding=0 width=98% border=0 height=0>"             
         response.write "<tr>"
		 response.write "<td align='center' bgcolor='#F3F3F3' width=75 height='20'><font class='small'>"
	     response.write    webdbRecordset.Fields("datein")   
	     response.write "</font></td>"
         response.write "<td align='center' bgcolor='#F3F3F3' width=75 height='20'><font class='small'>"
	     response.write    webdbRecordset.Fields("datein")   
	     response.write "</font></td>"
	      response.write "<td align='center' bgcolor='#F3F3F3'  width=75 height='20'><font class='small'>"
		 response.write  FormatDateTime(webdbRecordset.Fields("timein"),3)
		 response.write "</font></td>"
		   response.write "<td align='center' bgcolor='#F3F3F3' width=75 height='20'><font class='small'>"
		 response.write    webdbRecordset.Fields("dateout")
		 response.write "</font></td>"
		  response.write "<td align='center' bgcolor='#F3F3F3'  width=75 height='20'><font class='small'>"
		 response.write  FormatDateTime(webdbRecordset.Fields("timeout"),3) 
		 response.write "</font></td>"
		 response.write "</tr>"
		 response.write "</table>"
        webdbRecordset.MoveNext  

		
     loop     
	   
%>
            </table></TD>
          <TD>&nbsp;</TD>
        </TR>
        <TR> 
          <TD height="52" colspan="3" align=middle valign="top" class="small"><br> 
            &nbsp;<br> &nbsp;<BR>
            Copyright © 1997-2005 SofFac Technology Sdn Bhd <i>All Rights Reserved</i>. 
          </TD>
        </TR>
        <TR>
          <TD height="97">&nbsp;</TD>
          <TD>&nbsp;</TD>
          <TD>&nbsp;</TD>
        </TR>
      </TBODY>
    </TABLE>
  </center>
</div>
</BODY>




