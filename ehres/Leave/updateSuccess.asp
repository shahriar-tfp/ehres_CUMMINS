<html>
<link rel="stylesheet" type="text/css" HREF="../css/login.css">
<title>Redirect to Main Page</title>
<script langauage="JavaScript">
function Redirect()
{
	location.href= "/eHRES/Leave/app_approval.asp"
}
function RedirectWithDelay()
{
	window.setTimeout("Redirect();", 1000);
}
</script>
<body bgcolor="#ffffff" onload="RedirectWithDelay();">

<div align="center">
  <center>
  <table border="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td width="100%">
        <p align="center"><font class="bigmarineblue">You Have UPDATED Leave Application !</font></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>



