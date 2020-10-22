<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<!-- #include virtual ="/ehres/global/ADOVBS.ASP" -->
<% dim connect_string 

connect_string = "Provider=SQLOLEDB.1;Persist Security Info=False;UID=HRISMGR;PWD=TIGER;Initial catalog=HRDB_CSEM;Data Source=DESKTOP-SQCF4E5\DEV2017;Connect Timeout=900000"
%>
 
 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>eHRES</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 5.0" name=GENERATOR>

<link rel="stylesheet" type="text/css" HREF="../css/login.css">
</HEAD>

<body bgColor="#ffffff" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<div align="center">
  <center>
<table cellSpacing="0" cellPadding="0" border="0" width="100%">
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
  <!--<tr>
    <td vAlign="top" colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></td></tr>
  <tr>-->
  <TR>
    <TD vAlign=top colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></TD></TR>
  <TR>
    <TD vAlign=top align=center width="1" height="109"></TD>
  </center>
    <TD vAlign=top align=center width="100%" height="109">
      <p align="center"><IMG alt='Main Menu' 
      src="../Image/empemergency.gif" 
     border=0 width="712" height="88"><br></br>
     	 
	 <%
        set myconn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.Recordset")    

	myconn.open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=HRISMGR;PWD=TIGER;Initial catalog=HRDB_CSEM;Data Source=DESKTOP-SQCF4E5\DEV2017;Connect Timeout=900000"
		       
        'sql = "Exec sp_is_empprofile '" + trim(Session("EmpID")) + "','" + trim(Session("Regisno")) + "','Retrieve'"
        'sql = "Exec sp_sa_selEmpAddress 'ms0036','id'"
        'Response.Write sql       
        'sql = "Exec sp_is_selFamilyInfoweb '" + trim(Session("EmpID")) + "','eng'"
        sql = "Exec sp_is_selContact '" + trim(Session("EmpID")) + "','eng'"                'sql = "Exec sp_is_empprofile 'ms0036','185612-k','Retrieve'"   rs("status")
        'Response.Write sql       
        rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText 
      	
      	do until rs.EOF
			'if count = 1 then   '#eeeeee' #0099CC
				 colour = " bgcolor='#0099CC'"
				 colour1 = " bgcolor='#eeeeee'"
				 colour2 = " bgcolor='#000080'"
		    'else
				 'colour1 = "#000080"
		    'end if	
		    		  response.Write "<table width= '90%' border='0'>"
		    		  Response.write "<tr>"
		    		  Response.Write "<td width='7%' height='19' align='left'" + colour + "><font class='marinewhite'>Name</td>"	
		    		  Response.Write "<td width='25%' height='19' align='left'" + colour + "><font class='marinewhite'>" + rs("name") + "</td>"
		    		  Response.Write "</tr>"
		    		  Response.Write "<tr>"
		    		  Response.write "<td width='7%' height='19' align='left'" + colour1 + "><font class='small'>Address 1</td>"
		    		  Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("address1") + "</td>"	
		    		  'Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>Working Status</td>"
		    		  'Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("workingstatus") + "</td>"		
		    		  Response.Write "</tr>"
		    		  Response.Write "<tr>"
		    		  Response.write "<td width='7%' height='19' align='left'" + colour1 + "><font class='small'>Address 2</td>"
		    		  Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("address2") + "</td>"
		    		  'Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>Occupation</td>"
		    		  'Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("occupation") + "</td>"		
		    		  Response.Write "</tr>"
		    		  Response.Write "<tr>"
		    		  Response.write "<td width='7%' height='19' align='left'" + colour1 + "><font class='small'>Relationship</td>"
		    		  Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("description") + "</td>"
		    		  'Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>Company Name</td>"
		    		  'Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("company") + "</td>"		
		    		  Response.Write "</tr>"
		    		  Response.Write "<tr>"
		    		  Response.write "<td width='7%' height='19' align='left'" + colour1 + "><font class='small'>House Tel</td>"
		    		  Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("housetel") + "</td>"
		    		  'Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>Status</td>"
		    		  'Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("status") + "</td>"		
		    		  Response.Write "</tr>"
		    		  Response.Write "<tr>"
		    		  Response.write "<td width='7%' height='19' align='left'" + colour1 + "><font class='small'>Office Tel</td>"
		    		  Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("officetel") + "</td>"		  
		    		  Response.Write "</tr>"
		    		  Response.Write "<tr>"
		    		  Response.Write "<td width='7%' height='19' align='left'" + colour1 + "><font class='small'>Handphone No</td>"
		    		  Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("handphone") + "</td>"		
		    		  Response.Write "</tr>"
		    		  Response.Write "<tr>"
		    		  Response.Write "<td width='7%' height='19' align='left'" + colour1 + "><font class='small'>Email</td>"
		    		  Response.Write "<td width='25%' height='19' align='left'" + colour1 + "><font class='small'>" + rs("email") + "</td>"		
		    		  Response.Write "</tr>"
		    		  Response.Write "</table>"
		    		  Response.Write "<br>"		
		    
				        rs.MoveNext  
				        count = abs(count - 1)        
			        loop
		            'end if	
			        rs.close
			        set rs = nothing
			        myconn.close
			        set myconn = nothing
     %>			
      
      <!--</table>-->
      
     <!--</div>-->
      
      <p>&nbsp;</p>
      <table border="0" width="96%">
  <tr>
    <td width="100%" align="middle"><IMG border=0 height=4 src="/ehres/Image/dottedlinenav.gif" width=408></td>
  </tr>
  <tr>
    <td align="middle" colspan="2" width="936" height="40" class="small">
      &nbsp;<br>&nbsp;<p>Copyright © 1997-2005 SoftFac Technology Sdn Bhd <i>All 
      Rights Reserved</i>. </td></tr>
   </table>
      </div>
      <div align="center">      
  </table>     
</BODY>