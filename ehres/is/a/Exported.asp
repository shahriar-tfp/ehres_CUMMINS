<%
Response.Write("<!--#include virtual='/x-objects-new/content/restricted.inc'-->")
Response.Write("<!--#include virtual='/x-objects-new/content/all.inc'-->")
%>
<html>
<head>
<title>Seagate Crystal Reports Web Samples &amp; Utilities Active Server Pages
Samples</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="SUMMARY" content="$Header: /community/_sharedtemplates/pages/templates/content.asp 9     5/17/99 7:19p Eglass $">
<meta name="LOCATION" content="$Archive: /community/_sharedtemplates/pages/templates/content.asp $">
<meta name="REVISION" content="$Revision: 9 $">
<meta name="MODIFIED" content="$Date: 5/17/99 7:19p $">
<meta name="AUTHOR" content="$Author: Eglass $">
<meta name="PUBLISHER" content="Seagate Software">
<link rel="stylesheet" href="../images/main.css" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth = "0" marginheight = "0" bgcolor="#ffffff" text="#000000" link="#0000ff" vlink="#800080" alink="#ff0000" marginspacing = 0>
<table border="0" cellpadding="0" cellspacing="0" width="760">
  <tr align="left" valign="top"> 
    <td align="left" valign="top" width="130" height="1" bgcolor="#000000" class="xMASTHEADLOGO"><img src="../images/shim.gif" width="130" height="1" border="0"></td>
    <td align="left" valign="top" width="10" height="1" bgcolor="#000000" class="xMASTHEADLOGO"><img src="../images/shim.gif" width="10" height="1" border="0"></td>
    <td align="left" valign="top" width="400" height="1" bgcolor="#000000" class="xMASTHEADLOGO"><img src="../images/shim.gif" width="400" height="1" border="0"></td>
    <td align="left" valign="top" width="60" height="1" bgcolor="#000000" class="xMASTHEADLOGO"><img src="../images/shim.gif" width="60" height="1" border="0"></td>
    <td align="left" valign="top" width="10" height="1" bgcolor="#000000" class="xMASTHEADLOGO"><img src="../images/shim.gif" width="10" height="1" border="0"></td>
    <td align="left" valign="top" width="150" height="1" bgcolor="#000000" class="xMASTHEADLOGO"><img src="../images/shim.gif" width="150" height="1" border="0"></td>
  </tr>
  <tr align="left" valign="top"> 
    <td align="left" valign="center" height="28" colspan="6" bgcolor="#000000" class="xMASTHEADLOGO"><a href="/homepage/"><img src="../images/logo_seagate_black_small.gif" alt="Seagate Software" border="0" width="140" height="28"></a>
           </td>
  </tr>
  <tr align="left" valign="top"> 
    <td align="left" valign="top" height="22" colspan="6" bgcolor="#6699cc" class="xMASTHEADLIGHT">
            <img border="0" src="../images/titleheader.gif" width="400" height="23" align="middle"></td>
  </tr>
  <tr align="left" valign="top"> 
    <td align="left" valign="top" bgcolor="#6699cc" class="xMASTHEADLIGHT">&nbsp;</td>
    <td align="left" valign="top"><img src="../images/nub_topleft_navtomasthead.gif" width="10" height="10"></td>
    <td align="left" valign="top" colspan="4">
            </td>
  </tr>
  <tr align="left" valign="top"> 
    <td align="left" valign="top">
           <img name="Nnavbarmain_01_01" src="../images/navbar-main_01_01.gif" border="0" width="130" height="16"><a href="../default.htm"><br>
      <img border="0" src="../images/navbar-main_02_01.gif" width="130" height="16"><br></a><a href="Default.htm"><img border="0" src="../images/navbar-main_03_01.gif" width="130" height="16"></a><br>
      <a href="../Active%20Server%20Pages/default.htm"><img border="0" src="../images/navbar-main_04_01.gif" width="130" height="16"></a><br>
      <a href="../Java/Default.htm"><img border="0" src="../images/navbar-main_05_01.gif" width="130" height="16"></a><br>
      <a href="http://community.seagatesoftware.com"><img border="0" src="../images/navbar-main_06_01.gif" width="130" height="16"></a><br>
      <a href="http://www.seagatesoftware.com"><img border="0" src="../images/navbar-main_07_01.gif" width="130" height="16"></a><br>
      <a href="http://webacd.seagatesoftware.com"><img border="0" src="../images/navbar-main_08_01.gif" width="130" height="16"></a><br>
      <img border="0" src="../images/navbar-main_09_01.gif" width="130" height="16"></td>
    <td align="left" valign="top" rowspan="3"></td>
    <td align="left" valign="top" colspan="2" rowspan="2"><img border="0" src="../images/title-ActiveServerPageSamples.gif" width="440" height="30">
      <p>&nbsp;</p><FONT size=3></FONT>
<HR>
<H2>Exported File Type Cannot Be Viewed Natively by the Browser</H2>
<BR>
<a href="<%=session("ExportVirtualDirectory") & session("filename")%>">Click Here to Download File</a>
<HR>
<BR>
<form>
<center><input type=BUTTON Value="OK" onClick="CallDefaultScreen()"></CENTER>
</form>
<Script language="javascript">
function CallDefaultScreen()
{
location.href= "ReportExport.htm"
}
</script>
    <td align="left" valign="top" rowspan="3"></td>
    <td align="left" valign="top" rowspan="2"><!--webbot bot="PurpleText" PREVIEW="delete this; edit for sidebars" --></td>
  </tr>
  <tr align="left" valign="top"> 
    <td align="left" valign="top">
    		<!--webbot bot="PurpleText" PREVIEW="delete this; edit for callouts" -->
    </td>
  </tr>
  <tr align="left" valign="top"> 
    <td align="left" valign="top"><!-- do not edit --></td>
    <td align="left" valign="top" colspan="2">
</td>
    <td align="left" valign="top"><!-- do not edit --></td>
  </tr>
</table>
<!--webbot bot="PurpleText" PREVIEW="Only edit the marked areas in the table above. Do not edit outside the table." -->
</body>
</html>
