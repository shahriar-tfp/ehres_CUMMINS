<!-- #INCLUDE FILE = "medauthfunc.asp"-->
<%

if Session("EmpId") = "" then
	errorMsg = "Session has expired.  Please login again."
else
	mafid = Request.QueryString("id")
	if mafid = "" then
		call create_medauthform()
		if mafid <> "" then
			Response.Redirect ("medauthform.asp?id=" & mafid)
		end if
	end if

	if mafid <> "" then
		call open_medauthform(mafid)
	end if
end if
%>
<HTML>
<HEAD>
<TITLE>Medical Authorization Form - 
<%
	if mafid <> "" then
		Response.Write(mafid)
	else
		Response.Write("Invalid")
	end if
%>
</TITLE>
<style type='text/css'>
body {
font-family:arial;
font-size:11pt;
}
#mafpage {
width:18cm;
height:25cm;
}
#mafpagehead {
width:100%; 
border-bottom:1px solid;
padding:5px 0;
}
#mafcologo{
margin-left:0.2cm;
padding:10px;
float:left;
}
#mafheaderinfo{
margin-right:2cm;
padding:10px; 
float:left;
}
#mafbody {
clear:both;
margin:10px 50px;
}
#mafaddress{
font-size:10pt;
text-align:center;
}
#maftitle{
margin-top:16px;
font-size:14pt;
font-weight:bold;
text-align:center;
text-transform:uppercase;
}
#mafinfo {
float:right;
}
td.box {
border:1px solid black;
width:120px;
}
</style>
</HEAD>
<BODY>
<%if errorMsg = "" and mafid <> "" then%>
<div id="mafpage">
	<div id="mafpagehead">
		<div id="mafcologo"><img src="../image/cologo.gif" height="68" widht="190"/></div>
		<div id="mafheaderinfo">
			<div style="position:relative; font-size:12pt;font-weight:bold; width:100%; text-align:center;">Cummins Sales and Service Sdn Bhd (Company No.1005452-M)</div>
			<div style="position:relative; font-size:9pt; width:100%; text-align:center;margin-top:6px">(Formerly known as Cummins Scott & English Malaysia Sdn. Bhd.)</div>
		</div>
		<div style="clear:both" ></div>
	</div>
	
	<div id="mafbody">
	<div id='mafaddress'>
No. 12, Jalan Pemaju U1/15, Seksyen U1, Hicom Glenmarie Industrial Park,<br />
40150 Shah Alam, Selangor Darul Ehsan, Malaysia.<br />
P.O.Box 7684, 40724 Shah Alam, Selangor Darul Ehsan, Malaysia.<br />
Tel: 603-50228888   Fax: 603-50228822   http://www.cummins.com <br />
	</div>
	<div id="maftitle">Medical Authorisation Form</div>
	<div id="mafinfo">
	<table class="maf" cellpadding=3 cellspacing=10>
		<tr><td>No</td><td class="box" style="font-weight:bold;">
<%
	if mafid = "" then
		Response.Write("Invalid")
	else
		Response.Write(mafid)
	end if
%>
		</td></tr>
		<tr><td>Date</td><td class="box">
<%
	if dategenmaf = "" then
		Response.Write("&nbsp;")
	else
		Response.Write(Day(dategenmaf) & "-" & MonthName(Month(dategenmaf), true) & "-" & Year(dategenmaf))
	end if
%>
		</td></tr>
	<tr><td>Date of Expiry</td><td class="box">
<%
	if dategenmaf = "" then
		Response.Write("&nbsp;")
	else
		Response.Write(Day(dateexpmaf) & "-" & MonthName(Month(dateexpmaf), true) & "-" & Year(dateexpmaf))
	end if
%>
		</td></tr>
		<tr><td>Dept</td><td class="box">
<%
	if empdeptmaf = "" then
		Response.Write("&nbsp;")
	else
		Response.Write(empdeptmaf)
	end if
%>		
		&nbsp;</td></tr>
	</table>
	</div>
	<div style="clear:both;	padding:20px 0;">
Dear Sirs,<br /><br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Kindly examine Mr./Mrs./Miss<span style="text-decoration:underline">
<%
	Response.Write "&nbsp;&nbsp;&nbsp;"
	Response.Write empnamemaf & "&nbsp;&nbsp;&nbsp;"
%></span>who is a member of our staff and report thereon.
	</div>
	<div id="mafinfo" style="font-size:10pt; width:300px; text-align:center; margin-top:60px;"></div>
	<div style="clear:both;" />
	<div id="mafinfo" style="border-top:1px solid black; width:300px; text-align:center; margin-bottom:60px; font-weight: bold; text-transform: uppercase; color: red; font-family: Arial; font-variant: small-caps;">
		Authorized signature
        <br />
        <span style="font-size: 8pt">Authorized Name :                                  </span></div>
	<div style="clear:both; width:100%; border:1px solid black;">
		<div style="text-align:center; font-size:12pt; border-bottom:1px solid black; padding:3px 0;">For doctor's use only</div>
		<div style="padding:20px;">
			<div style="height:17px;margin-top:10px;">
				<div style="float:left;">Time in:</div>
				<div style="border-bottom:1px solid black;width:208px;float:left; height:100%;"></div>
				<div style="float:left;">Time out:</div>
				<div style="border-bottom:1px solid black;width:208px;float:left; height:100%;"></div>
				<div style="clear:both;"></div>
			</div>
			<div style="height:17px;margin-top:20px;">
				<div style="float:left;">Fit/Unfit for duty:</div>
				<div style="border-bottom:1px solid black;width:425px;float:left; height:100%;"></div>
				<div style="clear:both;"></div>
			</div>
			<div style="height:17px;margin-top:20px;">
				<div style="float:left;">No. of days medical leave:</div>
				<div style="border-bottom:1px solid black;width:360px;float:left; height:100%;"></div>
				<div style="clear:both;"></div>
			</div>
			<div style="margin-top:20px;">
				<div style="float:left; text-align:center; width:45%; border-top:1px solid black; margin-top:60px;">Date examined</div>
				<div style="float:right; text-align:center; width:45%; border-top:1px solid black; margin-top:60px;">Doctor's signature and stamp</div>
				<div style="clear:both;"></div>
			</div>
		</div>
	</div>
	<span style="font-size:8pt;">PSN-3</span>
	</div>
</div>
<%else%>
<span style="font-weight:bold;color:red;">Error:&nbsp;<%=errorMsg%></span>
<%end if%>
</BODY>
</HTML>
