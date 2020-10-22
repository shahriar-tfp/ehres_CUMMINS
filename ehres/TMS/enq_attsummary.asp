<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 5.0" name=GENERATOR>

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
      <P><IMG height=83 src="../Image/engtmssum.gif" width="679" border=0 ><BR><BR>
      <TABLE cellSpacing=0 width="100%" border=0>
        
        <TR>
          <TD>
            <form method="POST" action="enq_attsummary.asp" name="frmTmsSummary">
              <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              <font class="small">Month</font> <select size="1" name="cboMonth" style="font-size: 8pt">
				<%
					dim vMonth
					
					vMonth = Request("cboMonth")
					
				      Set webdb = Server.CreateObject("ADODB.Connection")
				          webdb.Open Session("ConnectStr")
				      Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
				      Set webdbCommand = Server.CreateObject("ADODB.Command")

				      ssql = "Exec sp_Wtms_SelTmsSummary '', '', 0, 0, '', 'CUR_MONTH'"
				             
				      Set webdbCommand.ActiveConnection = webdb
				          webdbCommand.CommandText = ssql
				          webdbRecordset.Open webdbCommand,,1 , 3
				          
					   If webdbRecordset.Fields("Month") = "1" and vMonth = "" Then
			              response.write "<option Selected value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "2" and vMonth = "" Then
			              response.write "<optionvalue = '1' >January</option>"
			              response.write "<option Selected value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "3" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option Selected value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "4" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option Selected value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "5" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option Selected value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "6" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option Selected value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "7" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option Selected value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "8" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option Selected value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "9" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option Selected value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "10" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option Selected value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "11" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option Selected value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf webdbRecordset.Fields("Month") = "12" and vMonth = "" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option Selected value = '12' >December</option>"
					   ElseIf vMonth = "1" Then
			              response.write "<option Selected value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "2" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option selected value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "3" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option selected value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "4" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option selected value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "5" Then
			              response.write "<option Selected value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option selected value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "6" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option selected value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "7" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option selected value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "8" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option selected value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "9" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option selected value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "10" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option selected value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "11" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option Selected value = '11' >November</option>"
			              response.write "<option value = '12' >December</option>"
					   ElseIf vMonth = "12" Then
			              response.write "<option value = '1' >January</option>"
			              response.write "<option value = '2' >February</option>"
			              response.write "<option value = '3' >March</option>"
			              response.write "<option value = '4' >April</option>"
			              response.write "<option value = '5' >May</option>"
			              response.write "<option value = '6' >June</option>"
			              response.write "<option value = '7' >July</option>"
			              response.write "<option value = '8' >August</option>"
			              response.write "<option value = '9' >September</option>"
			              response.write "<option value = '10' >October</option>"
			              response.write "<option value = '11' >November</option>"
			              response.write "<option Selected value = '12' >December</option>"			              
                    End If
				      webdbRecordset.close
				      webdb.close
				%>
                 &nbsp; </select>&nbsp;&nbsp;&nbsp;
              
              <font class="small">Year</font>&nbsp;<select size="1" name="cboYear" style="font-size: 8pt">

				<%
				      Set webdb = Server.CreateObject("ADODB.Connection")
				          webdb.Open Session("ConnectStr")
				      Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
				      Set webdbCommand = Server.CreateObject("ADODB.Command")

				      ssql = "Exec sp_Wtms_SelTmsSummary '', '', 0, 0, '', 'YEAR'"
       
				      Set webdbCommand.ActiveConnection = webdb
				          webdbCommand.CommandText = ssql
				          webdbRecordset.Open webdbCommand,,1 , 3

				 	   i = 1
				 	   Do Until webdbRecordset.EOF
				 	      If i = 2 and Request("cboYear") = "" Then
					         response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("year")) + "'>" + cstr(webdbRecordset.Fields("year")) + "</OPTION> "
					      Elseif Request("cboYear") = cstr(webdbRecordset.Fields("year")) then
				 	         response.write "<OPTION Selected value='" + cstr(webdbRecordset.Fields("year")) + "'>" + cstr(webdbRecordset.Fields("year")) + "</OPTION> "
				 	      else   
					         response.write "<OPTION value='" + cstr(webdbRecordset.Fields("year")) + "'>" + cstr(webdbRecordset.Fields("year")) + "</OPTION> "
					      End If
					      i = i + 1
					      webdbRecordset.MoveNext
				      loop
				      webdbRecordset.close
				      webdb.close
				%>
              
              </select> 
&nbsp;&nbsp;&nbsp;
              
              <font class="small">Cut Off</font>&nbsp;<select size="1" name="cboCutOff" style="font-size: 8pt">

				<%
                  if Request("cboCutOff") = "1" or Request("cboYear") = "" then
 		  	         response.write "<OPTION Selected value='1'>1st Half Cut Off</OPTION> "
 		  	      else   
 		  	         response.write "<OPTION value='1'>1st Half Cut Off</OPTION> "
 		  	      end if   

                  if Request("cboCutOff") = "2" then
 		  	         response.write "<OPTION Selected value='2'>2nd Half Cut Off</OPTION> "
 		  	      else   
 		  	         response.write "<OPTION value='2'>2nd Half Cut Off</OPTION> "
 		  	      end if   

				%>
              
              </select>               
				<input type="submit" value="Refresh" name="cmdRefresh" style="font-size: 8pt"></p>
            </form>
         </TD>
       </TR>
     </TABLE>&nbsp;</P>
   </TD></TR>
  <TR>
    <TD vAlign=top align=middle colspan=2 width="100%" height="193">
      <div align="center">
        <center>
      <TABLE cellSpacing=0 cellPadding=0 width="90%" border=0>
        <TBODY>

    <tr><td></td><td><font class="bigmarineblue">Overtime Summary</font></td></tr>
    <tr>
      <td bgcolor="#f3f3f3" width="10">&nbsp;</td>
	  <td width="40%" bgcolor="#f3f3f3"><font class="marineblack"><b>Day Type</b></font></td>
	  <td width="20%" bgcolor="#f3f3f3"><font class="marineblack"><b>Rate</b></font></td>
	  <td width="18%" bgcolor="#f3f3f3"><font class="marineblack"><b>Original</b></font></td>
	  <td width="18%" bgcolor="#f3f3f3"><font class="marineblack"><b>Adjustment</b></font></td>
    </tr>

    <%
       Set webdb = Server.CreateObject("ADODB.Connection")
   		   webdb.Open Session("ConnectStr")
 	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  	   Set webdbCommand = Server.CreateObject("ADODB.Command")

	   if request("cboMonth") = "" and request("cboYear") = "" then
	      ssql = "Exec sp_Wtms_SelTmsSummary """ + Session("Regisno") + """, """ + Session("EmpID") + """, 0, 0 , '1', 'OT'"
	   else
	      ssql = "Exec sp_Wtms_SelTmsSummary """ + Session("Regisno") + """, """ + Session("EmpID") + """, " _
	       		 + request("cboMonth") + ", " + request("cboYear") + " , """ + request("cboCutOff") + """, 'OT'"
	   end if


	   Set webdbCommand.ActiveConnection = webdb
	       webdbCommand.CommandText = ssql
	       webdbRecordset.Open webdbCommand,,1 , 3

	  	   colour = 0

       Do Until webdbRecordset.EOF
         if count = 1 then
             colour = " bgcolor='#eeeeee'"
          else
             colour = ""
          end if
			        
          response.write "<tr>"			
          response.write "<td " + colour + "> </td>"
          response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("DayType") + "</td>"
          response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("Range") + "</td>"
          response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("Original") + "</td>"
          response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("Adjustment") + "</td></tr>"
          webdbRecordset.MoveNext  
          count = abs(count - 1)        
       loop

       webdbRecordset.close
       webdb.close      

    %>
    
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td></td><td><font class="bigmarineblue">Shift Summary</font></td></tr>
    
    <tr>
      <td bgcolor="#f3f3f3" width="10">&nbsp;</td>
      <td width="40%" bgcolor="#f3f3f3"><font class="marineblack"><b>Shift</b></font></td>
      <td width="20%" bgcolor="#f3f3f3"><font class="marineblack"><b>TMS ID</b></font></td>
      <td width="18%" bgcolor="#f3f3f3"><font class="marineblack"><b>Original</b></font></td>
      <td width="18%" bgcolor="#f3f3f3"><font class="marineblack"><b>Adjustment</b></font></td>
    </tr>
    
    <%   
      Set webdb = Server.CreateObject("ADODB.Connection")
   		  webdb.Open Session("ConnectStr")
  	  Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  	  Set webdbCommand = Server.CreateObject("ADODB.Command")

 	  if request("cboMonth") = "" and request("cboYear") = "" then
	     ssql = "Exec sp_Wtms_SelTmsSummary '" + Session("Regisno") + "', '" + Session("EmpID") + "',0,0 , '1', 'SHIFT'"
	  else
	     ssql = "Exec sp_Wtms_SelTmsSummary '" + Session("Regisno") + "', '" + Session("EmpID") + "', '" + Request("cboMonth") + "', '" + Request("cboYear") + "' , '" + request("cboCutOff") + "', 'SHIFT'"
	  end if   

      Set webdbCommand.ActiveConnection = webdb
	      webdbCommand.CommandText = ssql
	      webdbRecordset.Open webdbCommand,,1 , 3

	  colour = 0
          count = 0
	  Do Until webdbRecordset.EOF
	     if count = 1 then
	        colour = " bgcolor='#eeeeee'"
	     else
	        colour = ""
	     end if
			        
		 response.write "<tr>"	
                 response.write "<td height='20' " + colour + "> </td>"
		 response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("ShiftID") + "</td>"
  		 response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields("TmsID") + "</td>"
		 response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields(2) + "</td>"
		 response.write "<td height='20' align='left'" + colour + "><font class='small'>" + webdbRecordset.Fields(3) + "</td></tr>"
		 webdbRecordset.MoveNext  
		
		 count = abs(count - 1)        
	  loop
	
	  webdbRecordset.close
	  webdb.close      
	 %>	

    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td></td><td><font class="bigmarineblue">Late/ Leave Early</font></td></tr>

    <tr>
      <td bgcolor="#f3f3f3" width="10">&nbsp;</td>
	  <td width="25%" bgcolor="#f3f3f3"><font class="marineblack"><b>Type</b></font></td>
	  <td width="15%" bgcolor="#f3f3f3"><font class="marineblack"><b>Original (minute)</b></font></td>
	  <td width="15%" bgcolor="#f3f3f3"><font class="marineblack"><b>Adjustment (minute)</b></font></td>
      <td bgcolor="#f3f3f3" width="10">&nbsp;</td>
    </tr>
    
    <%   
       dim l1  
       dim l2  
       dim le1 
       dim le2 
       dim a1  
       dim a2  
       
    
       Set webdb = Server.CreateObject("ADODB.Connection")
   		   webdb.Open Session("ConnectStr")
  	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
  	   Set webdbCommand = Server.CreateObject("ADODB.Command")


	   if request("cboMonth") = "" and request("cboYear") = "" then
          ssql = "Exec sp_Wtms_SelTmsSummary """ + Session("Regisno") + """, """ + Session("EmpID") + """, 0,0, '1', 'LATE'"
	   else
 	      ssql = "Exec sp_Wtms_SelTmsSummary """ + Session("Regisno") + """, """ + Session("EmpID") + """, " + Request("cboMonth") + ", " + Request("cboYear") + " , """ + request("cboCutOff") + """, 'LATE'"
	   end if

	   Set webdbCommand.ActiveConnection = webdb
	       webdbCommand.CommandText = ssql
	       webdbRecordset.Open webdbCommand,,1 , 3

	   colour = 0
           count = 0
       
	   Do Until webdbRecordset.EOF
	      if count = 1 then
	         colour = " bgcolor='#eeeeee'"
	      else
	         colour = ""
	      end if
			        
	      l1  = webdbRecordset.Fields(0)
	      l2  = webdbRecordset.Fields(1)
	      le1 = webdbRecordset.Fields(2)
	      le2 = webdbRecordset.Fields(3)
              a1  = "0"'webdbRecordset.Fields(4)
              a2  = "0"'webdbRecordset.Fields(5)
	      
	      webdbRecordset.MoveNext  
	      count = abs(count - 1)        
	  loop
	
	  webdbRecordset.close
	  webdb.close      

	  response.write "<tr>"
	  response.write "<td></td>"
	  response.write "<td height='20' align='left'><font class='small'>" + "Late" + "</td>"
	  response.write "<td height='20' align='left'><font class='small'>" + l1 + "</td>"
	  response.write "<td height='20' align='left'><font class='small'>" + l2 + "</td>"
	  response.write "<td></td>"
	  response.write "</tr>"

	  response.write "<tr>"
	  response.write "<td bgcolor='#eeeeee'> </td>"
	  response.write "<td height='20' align='left' bgcolor='#eeeeee' ><font class='small'>" + "Leave Early" + "</td>"
	  response.write "<td height='20' align='left' bgcolor='#eeeeee' ><font class='small'>" + le1 + "</td>"
	  response.write "<td height='20' align='left' bgcolor='#eeeeee' ><font class='small'>" + le2 + "</td>"
	  response.write "<td bgcolor='#eeeeee'></td>"
	  response.write "</tr>"

	  response.write "<tr>"
	  response.write "<td></td>"
	  response.write "<td height='20' align='left' ><font class='small'>" + "Absent Hour" + "</td>"
	  response.write "<td height='20' align='left' ><font class='small'>" + a1 + "</td>"
	  response.write "<td height='20' align='left' ><font class='small'>" + a2 + "</td>"
	  response.write "<td></td>"
	  response.write "</tr>"

	    
    %> 	 

 		 
    </TBODY>
    </TABLE>
    </center>
    </div>
    </TD></TR>
  <TR>
    <TD align=middle colspan=2 width="100%" height="40" class="small"><br>
      &nbsp;<br>
      &nbsp;<BR>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All Rights Reserved</i>. </TD></TR></TBODY></TABLE></center>
</div>
</BODY>