<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head><title>eHRES</title>

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<link rel="stylesheet" type="text/css" HREF="css/login.css">
</head>
<body bgColor="#ffffff" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<table cellSpacing="0" cellPadding="0" border="0" height="698" width="100%">
  <tbody>
  <tr>
    <td height="1" width="100%">
      <table cellSpacing="0" cellPadding="0" border="0" width="100%" height="9">
        <tr>
          <td vAlign="top" align="middle" colspan="2" width="936" bgcolor="#0099cc" height="1">
            <div align="center">
              <center>
              <table border="0" width="100%">
                <tr>
                  <td width="3%"><a name="top"></a></td>
                  <td width="23%"><font class="marinewhite">Employee ID : <%    
          response.write session("EmpID")
                    %>
                    </font></td>
                  <td width="74%"><font class="marinewhite">Name : <%   '   changePass.asp
          response.write session("EmpName")
                    %>
                    </font></td>
                </tr>
              </table>
              </center>
            </div>
          </td>
        </tr>
      </table>
    </td></tr>
  <tr>
    <td vAlign="top" height="41" width="100%">
      <p align="right"><img height="1" src="image/clear.gif" width="60"><a href="changePass.asp"><font class="marinered">ChangePassword</font></a> | <a href="signout.asp"><font class="marinered">Logout</font></a><!--<font class="marinered">
      | <font class="marinered"><a href="changePass.asp">ChangePassword</a></font>--></p>
    </td></tr>
  <tr>
    <td vAlign="top" align="left" height="73" width="100%">
      <p align="center"><img alt="business" src="Image/title-3.gif" align="middle" border="0" WIDTH="713" HEIGHT="71"></p>
    </td></tr>
  <form name="cust_matrix" action method="post">
  <tr>
    <td vAlign="top" align="left" height="597" width="100%">
      <table cellSpacing="0" cellPadding="0" width="749" border="0">
        <tbody>
        <tr>
          <td width="25" rowSpan="100"><img height="1" src="image/clear.gif" width="25"> 
        </td></tr>
        <tr>
          <td><img height="33" src="Image/matrix_top1.jpg" width="97"></td>
          <td vAlign="center" background="image/matrix_top2.gif" colSpan="6">
            <nobr>
            </nobr></td>
          <td width="189" colSpan="2"><img height="33" src="Image/matrix_top3.gif" width="189"></td></tr>
          
          <tr>
          <td vAlign="top" width="97" rowSpan="5"><img height="94" alt="Employee Personal Profile" src="Image/is_profile.gif" width="97"> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3"><font class="white">Employee Personal Profile</font> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3">&nbsp;</td>
          <td vAlign="top" width="189" colSpan="2"><img height="23" src="Image/left.gif" width="4" border="0"><img height="23" src="Image/blank2.gif" width="63" border="0"><img height="23" src="Image/divider_blank.gif" width="4" border="0"><a href="#top"><img height="23" src="Image/top.gif" width="29" border="0"></a><img height="23" src="Image/divider_blank.gif" width="4" border="0"><img height="23" src="Image/blank.gif" width="79" border="0"><img height="23" src="Image/right.gif" width="4" border="0"></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Profile Viewing</div></td>
          <td width="140" colSpan="2">
            <!--<div id="matrixCSEQ" class="marineblack">Enquiry</div>--></td>
          <td width="140" colSpan="2">
            <!--<div id="matrixCSEQ" class="marineblack">Approval</div>--></td>
          <td width="189">
            <!--<div id="matrixCSEQ" class="marineblack">Reports</div>--></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; <a href="empprofile/emp_profile.asp"><font color="#000000">Personal Profile</font></a>&nbsp;<br>
              &#149; <a href="empprofile/emp_address.asp"><font color="#000000">Address</font></a><br>&#149; <a href="empprofile/emp_typeEmploy2.asp"><font color="#000000">Type of Employment</font></a>&nbsp;<br>&#149; <a href="empprofile/emp_familyprofile.asp"><font color="#000000">Family Profile</font></a></div></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; <a href="empprofile/emp_emergencycontact.asp"><font color="#000000">Emergency Contact</font></a><br>
              <!--&#149; <a href="Leave/enq_leaveapp.asp"><font color="#000000"> Leave Application&nbsp;</font></a></div></td>-->
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <!--<div id="matrixText" class="small">&#149; <a href="empprofile/emp_typeEmploy.asp"><font color="#000000">Type of Employment</font></a> </div></td>-->
          <td vAlign="top" align="right"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="189">
            <div id="matrixText" class="small">None 
          </div></td></tr>
               
        <tr>
          <td vAlign="top" width="97" rowSpan="5"><img height="94" alt="Leave System" src="Image/mleave.gif" width="97"> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3"><font class="white">Leave
            System</font> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3">&nbsp;</td>
          <td vAlign="top" width="189" colSpan="2"><img height="23" src="Image/left.gif" width="4" border="0"><img height="23" src="Image/blank2.gif" width="63" border="0"><img height="23" src="Image/divider_blank.gif" width="4" border="0"><a href="#top"><img height="23" src="Image/top.gif" width="29" border="0"></a><img height="23" src="Image/divider_blank.gif" width="4" border="0"><img height="23" src="Image/blank.gif" width="79" border="0"><img height="23" src="Image/right.gif" width="4" border="0"></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Application</div></td>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Enquiry</div></td>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Approval</div></td>
          <td width="189">
            <div id="matrixCSEQ" class="marineblack">Reports</div></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; <a href="Leave/app_applyleave.asp"><font color="#000000"> Leave Application</font></a>&nbsp;<br>
              &#149; <a href="Leave/app_leavecancel.asp"><font color="#000000"> Leave
              Cancellation</font></a>&nbsp;</div></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; <a href="Leave/enq_balance.asp"><font color="#000000"> Leave Balance</font></a><br>
              &#149; <a href="Leave/enq_leaveapp.asp"><font color="#000000"> Leave Application&nbsp;</font></a></div></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; <a href="Leave/app_approval.asp"><font color="#000000"> Leave Approval</font></a> </div></td>
          <td vAlign="top" align="right"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="189">
            <div id="matrixText" class="small">None 
</div></td></tr>
        <tr>
          <td vAlign="top" width="97" rowSpan="5"><img height="94" alt="Time Management System" src="Image/mtms.gif" width="97"> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3"><font class="white">Time
            Management System</font> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3">&nbsp;</td>
          <td vAlign="top" width="189" colSpan="2"><img height="23" src="Image/left.gif" width="4" border="0"><img height="23" src="Image/blank2.gif" width="63" border="0"><img height="23" src="Image/divider_blank.gif" width="4" border="0"><a href="#top"><img height="23" src="Image/top.gif" width="29" border="0"></a><img height="23" src="Image/divider_blank.gif" width="4" border="0"><img height="23" src="Image/blank.gif" width="79" border="0"><img height="23" src="Image/right.gif" width="4" border="0"></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Application</div></td>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Enquiry</div></td>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Approval</div></td>
          <td width="189">
            <div id="matrixCSEQ" class="marineblack">Reports</div></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; Overtime Claim</div></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; <a href="TMS/enq_attendance.asp"><font color="#000000">Attendance</font></a><br>
              &#149; <a href="TMS/enq_attsummary.asp"><font color="#000000">TMS, OT &amp; Shift Summary</font></a><br>
              &#149; <a href="TMS/enq_atterror.asp"><font color="#000000">Attendance Error</font></a></br> 
              <!--&#149; <a href="TMS/enq_lateleaveEarly.asp"><font color="#000000">Late Only / Leave Early</font></a>--></div></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; <a href="TMS/tmsapproval.asp"><font color="#000000">Overtime Approval</font></div>
            <!--<div id="matrixText" class="small">&#149;<a href="TMS/enq_attendance1.asp"><font color="#000000">Staff Attendance</div></font></a></div>--></td>
          <td vAlign="top" align="right"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          
          <tr>
          <td vAlign="top" width="97" rowSpan="5"><img height="94" alt="Patroll System" src="Image/payroll.gif" width="97"> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3"><font class="white">Payroll System</font> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3">&nbsp;</td>
          <td vAlign="top" width="189" colSpan="2"><img height="23" src="Image/left.gif" width="4" border="0"><img height="23" src="Image/blank2.gif" width="63" border="0"><img height="23" src="Image/divider_blank.gif" width="4" border="0"><a href="#top"><img height="23" src="Image/top.gif" width="29" border="0"></a><img height="23" src="Image/divider_blank.gif" width="4" border="0"><img height="23" src="Image/blank.gif" width="79" border="0"><img height="23" src="Image/right.gif" width="4" border="0"></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Monthly Salary</div></td>
          <td width="140" colSpan="2">
            <!--<div id="matrixCSEQ" class="marineblack">Enquiry</div>--></td>
          <td width="140" colSpan="2">
            <!--<div id="matrixCSEQ" class="marineblack">Approval</div>--></td>
          <td width="189">
            <!--<div id="matrixCSEQ" class="marineblack">Reports</div>--></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td vAlign="top" width="140">
           <div id="matrixText" class="small">&#149; <a href="Payroll/emp_payslip.asp"><font color="#000000">Payslip Viewing</font></a><br>
              &#149; <a href="payroll/emp_salary.asp"><font color="#000000">Payroll Accounts Profile&nbsp;</font></a></div></td>
           <!--<div id="matrixText" class="small"><a href="TMS/enq_attendance.asp">&#149; Payslip Viewing</div></a></td>-->
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
         <td vAlign="top" width="140">
            <!--<div id="matrixText" class="small">&#149; <a href="TMS/enq_attendance.asp"><font color="#000000">Attendance</font></a><br>
              &#149; <a href="TMS/enq_attsummary.asp"><font color="#000000">TMS, OT &amp; Shift Summary</font></a><br>
              &#149; <a href="TMS/enq_atterror.asp"><font color="#000000">Attendance Error</font></a></div>--></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <!--<div id="matrixText" class="small">&#149;<a href="TMS/tmsapproval.asp"><font color="#000000">Overtime Approval</font></div>
            <div id="matrixText" class="small">&#149;<a href="TMS/enq_attendance1.asp"><font color="#000000">Staff Attendance</div></font></a></div>--></td>
          <td vAlign="top" align="right"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          </tr>
          <!-- new KPI -->
          <tr>
          <td vAlign="top" width="97" rowSpan="5"><img height="94" alt="Patroll System" src="Image/mconfirm.gif" width="97"></td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3"><font class="white">Performance Evaluation</font> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3">&nbsp;</td>
          <td vAlign="top" width="189" colSpan="2"><img height="23" src="Image/left.gif" width="4" border="0"><img height="23" src="Image/blank2.gif" width="63" border="0"><img height="23" src="Image/divider_blank.gif" width="4" border="0"><a href="#top"><img height="23" src="Image/top.gif" width="29" border="0"></a><img height="23" src="Image/divider_blank.gif" width="4" border="0"><img height="23" src="Image/blank.gif" width="79" border="0"><img height="23" src="Image/right.gif" width="4" border="0"></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td width="140" colSpan="2">
            <div id="Div1" class="marineblack">Application</div></td>
          <td width="140" colSpan="2">
            <!--<div id="matrixCSEQ" class="marineblack">Enquiry</div>--></td>
          <td width="140" colSpan="2">
            <!--<div id="matrixCSEQ" class="marineblack">Approval</div>--></td>
          <td width="189">
            <!--<div id="matrixCSEQ" class="marineblack">Reports</div>--></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td vAlign="top" width="140">
           <div id="Div2" class="small"><!--&#149; <a href="KPI/emp_KPIprofile.asp"><font color="#000000">KPI Profile</font></a><br>//-->
              &#149; <a href="TD/TD_PosComp.asp"><font color="#000000">Position Competency&nbsp;</font></a><br>
              &#149; <a href="TD/TD_GapAnalysis.asp"><font color="#000000">Gap Analysis&nbsp;</font></a><br>
           &#149; <a href="TD/TD_TranEvalution.asp"><font color="#000000">Tranning Evalution</font></a></div></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
         <td vAlign="top" width="140">
            <!--<div id="matrixText" class="small">&#149; <a href="TMS/enq_attendance.asp"><font color="#000000">Attendance</font></a><br>
              &#149; <a href="TMS/enq_attsummary.asp"><font color="#000000">TMS, OT &amp; Shift Summary</font></a><br>
              &#149; <a href="TMS/enq_atterror.asp"><font color="#000000">Attendance Error</font></a></div>--></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <!--<div id="matrixText" class="small">&#149;<a href="TMS/tmsapproval.asp"><font color="#000000">Overtime Approval</font></div>
            <div id="matrixText" class="small">&#149;<a href="TMS/enq_attendance1.asp"><font color="#000000">Staff Attendance</div></font></a></div>--></td>
          <td vAlign="top" align="right"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <!--<td vAlign="top" width="189">
            <div id="matrixText" class="small">None </div></td></tr>
  
          <tr>
          <td vAlign="top" width="97" rowSpan="5"><img height="94" alt="Confirmation of employment" src="Image/mconfirmation.gif" width="97"> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3"><font class="white">Confirmation of Employment</font> </td>
          <td vAlign="top" background="image/spacer2.jpg" colSpan="3">&nbsp;</td>
          <td vAlign="top" width="189" colSpan="2"><img height="23" src="Image/left.gif" width="4" border="0"><img height="23" src="Image/blank2.gif" width="63" border="0"><img height="23" src="Image/divider_blank.gif" width="4" border="0"><a href="#top"><img height="23" src="Image/top.gif" width="29" border="0"></a><img height="23" src="Image/divider_blank.gif" width="4" border="0"><img height="23" src="Image/blank.gif" width="79" border="0"><img height="23" src="Image/right.gif" width="4" border="0"></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Application</div></td>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Enquiry</div></td>
          <td width="140" colSpan="2">
            <div id="matrixCSEQ" class="marineblack">Approval</div></td>
          <td width="189">
            <div id="matrixCSEQ" class="marineblack">Reports / Letters</div></td></tr>
        <tr>
          <td colSpan="7" height="5"></td></tr>
        <tr>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; None</div></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; <a href="TMS/enq_attendance.asp"><font color="#000000">Staff Confirmation</font></a></div></td>
          <td vAlign="top"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="140">
            <div id="matrixText" class="small">&#149; <a href="is/app_confirm.asp"><font color="#000000">Staff Confirmation</font></a></div></td>
          <td vAlign="top" align="right"><img height="109" src="Image/vertical_dashed_line2.gif" width="11"> 
          </td>
          <td vAlign="top" width="189">
            <div id="matrixText" class="small">
            &#149; Confirmed List <br>
            &#149; Extension List <br>
            &#149; Extension Letter <br>
            &#149; Termination Letter <br>
            </div></td>--></tr>          
  </tbody></table></td></tr></form>
  <tr>
    <td align="middle" height="40" width="100%"><br> </td></tr></tbody></table></body></html>
