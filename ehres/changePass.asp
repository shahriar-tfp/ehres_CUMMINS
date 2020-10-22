<!-- #INCLUDE VIRTUAL = "/ehres/global/ConnectStr.asp"-->
<!-- #INCLUDE VIRTUAL = "/ehres/global/AdoVbs.asp"-->
<% dim PASSWORD_MIN_LEN
   dim connect_string
   connect_string =Session("ConnectStr")

%>

<HTML><HEAD><TITLE>Employee Sign In</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<LINK href="css/login.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>


<BODY vLink=#000099 aLink=#000099 link=#000099 bgColor=#ffffff>
<script language ="JavaScript">
	function change(){
		document.frmchangepass.txtAction.value ="change"
		}
</script>
<script LANGUAGE="VBScript" ></script>

<% 
   PASSWORD_MIN_LEN =6
   function msgBox(strMessage)
   dim strHTML
	strHTML = "<script language=""JavaScript"">"
	strHTML = strHTML & "alert('"& strMessage &"');"
	strHTML = strHTML & "history.go(-1);"
	strHTML = strHTML & "</script>"
	Response.Write strHTML
   end function
   
   password = Request("txtExisPass")
   newpass1 = Request("txtNewPass")
   newpass2 = Request("txtConfPass")

'if Request.ServerVariables("Content_Length") > 0 then
if Request.Form("txtAction")="change" then

     if password = "" then
        call MsgBox("You must enter your current password, try again.")
     elseif newpass1 = "" or newpass2 = "" then
       call MsgBox("You must select a new password and enter it twice, try again")
     elseif newpass1 <> newpass2 then
       call MsgBox("New password fields did not match, try again.")
     elseif Len(newpass1) < PASSWORD_MIN_LEN then
       call MsgBox("Passwords must be at least " & PASSWORD_MIN_LEN & " characters long, try again.")
     else

       'Verify current password.

       'call OpenDB()
       set myconn = server.CreateObject("ADODB.Connection")
       set rs = server.CreateObject("ADODB.Recordset")
           myconn.open connect_string
       
       'sql = "Exec sp_wls_selwebpass '"+ Session("EmpID")+"','"+ password +"'"
       sql = "Exec sp_wls_selwebpass '"+ Session("EmpID")+"'"
	'Response.Write sql
	'response.End
	   rs.Open sql, myconn, adopenstatic, adLockReadOnly, adCmdText
       'sql = "select * from Users" _
        '  & " where Username = '" & Session("PoolUsername") & "'"
       'set rs = DbConn.Execute(sql)
       
       
       
       if (rs.BOF and rs.EOF) then
           'call ErrorMessage("User '" & username & "' not found, try again.")
         elseif rs.Fields("password") <> password then
         'elseif Hash(rs.Fields("password").Value & password) <> rs.Fields("Password").Value then
           call MsgBox("Incorrect password, try again.")
         else
            
           'Update record with new seed and password.
			
			set myconn = server.CreateObject("ADODB.Connection")
            set rs = server.CreateObject("ADODB.Recordset")
                myconn.open connect_string
	   
	       ssql ="Exec sp_sa_changepassword1 '"+ Session("EmpID")+"','"+ password +"','"+ newpass1 +"','"+ newpass2 +"','CHANGE'"
	       'response.Write ssql
	       'response.End
	       rs.Open ssql, myconn, adopenstatic, adLockReadOnly, adCmdText
		   
		   'temp = rs("RanNo")
		         
           'if temp = "" then 
				'call MsgBox("Your password has been changed.")
				'response.Redirect "../main.asp"
				call MsgBox("Your password has been changed.")
		   'else
			'	call MsgBox("Invalid user ID.")
		   'end if			
         end if
     end if

   end if %>

<table cellSpacing="0" cellPadding="0" border="0" width="100%">
  <tr>
    <td vAlign="top" colspan="2" width="100%" height="21" class="small" align="center">
      <p align="right"><a href="../main.asp"><font color="#000000">Home</font></a>&nbsp;&nbsp;&nbsp;
      |&nbsp;&nbsp;&nbsp; <a href="../signout.asp"><font color="#000000">Logout</font></a></td></tr>
  
</table>
<BR>
<TABLE borderColor=#000066 cellSpacing=0 cellPadding=5 align=center border=1>
  <TBODY>
  <TR bgColor=#000066>
    <TD class=white>
      <DIV align=center><font color="#ffffff">Please enter Existing Password and New Password.</font></DIV></TD></TR>
  <TR>
    <TD>
    <form action="changepass.asp" method="post" name="frmchangepass">
     
      <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
        <TBODY>
        <TR>
          <TD class=normal>Existing Password</TD>
          <TD><INPUT class=small name="txtExisPass" type=password maxlength="30" tabindex="1" ><input name=txtAction type=hidden></TD></TR>
        <TR>
          <TD class=normal>New Password</TD>
          <TD><INPUT class=small type=password name="txtNewPass" maxlength="30" tabindex="2" > </TD></TR>
        <TR>
          <TD class=normal>Confirmed Password</TD>
          <TD><INPUT class=small type=password name="txtConfPass" maxlength="30" tabindex="2" > </TD></TR>
        <TR>
          <TD>&nbsp;</TD>
          <TD>
              <INPUT class=small type=submit value=Submit name=Submit tabindex="3" onclick =change()>&nbsp;&nbsp;&nbsp;&nbsp;
        </TD></TR></TBODY></TABLE></FORM>
      <P></P></TD></TR></TBODY></TABLE><BR>
<TABLE cellSpacing=0 cellPadding=0 align=center border=0>
  <TBODY>
  <TR>
    <TD align=middle width="100%">
      <HR>

    </TD></TR>
  <TR>
    <TD align=middle width="100%"><BR>
      <P class=small>Copyright © 1997-2007 SoftFac Technology Sdn Bhd <I>All Rights Reserved</I>. 
  </P></TD></TR></TBODY></TABLE><IMG 
height=5 src="image/count.gif" width=5 border=0> 
</BODY></HTML>
