<!-- #include virtual ="/ehres/global/ConnectStr.asp"-->
<% 
dim mafid
dim empidmaf
dim empnamemaf
dim dategenmaf
dim dateexpmaf
dim empdeptmaf
dim empsalutationmaf
dim errorMsg
sub create_medauthform()
	if Session("Regisno") <> "" and Session("EmpId") <> "" then
		set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open  Session("ConnectStr")
		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.maxRecords = 1
		oRs.Open "EXEC usp_mafCreate @regisno='" & Session("Regisno") & "', @empid='" & Session("EmpId") & "'", oConn
		if Not oRs.BOF and Not oRs.EOF then
			if oRs.Fields("Result") = 0 then
				mafid = oRs.Fields("Output")
			else
				errorMsg = oRs.Fields("Output")
			end if
		else
			errorMsg = "Fail to generate Medical Authorisation Form."
		end if
		oRs.Close
		set oRs = nothing
		oConn.Close
		set oConn = nothing
	end if
end sub

sub open_medauthform(vMafid)
	if vMafid <> "" then
		set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open  Session("ConnectStr")
		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open "EXEC usp_mafOpen @regisno='" & Session("Regisno") & "', @mafid='" & vMafid & "'", oConn, 0, 1
		
		if Not oRs.BOF and Not oRs.EOF then
			mafid = oRs.Fields("mafid")
			empidmaf = oRs.Fields("empid")
			empnamemaf = oRs.Fields("empname")
			dategenmaf = oRs.Fields("dategenerated")
			dateexpmaf = oRs.Fields("dateexpiry")
			empdeptmaf = oRs.Fields("dept")
			empsalutationmaf = oRs.Fields("title")
		else
			mafid = ""
			empidmaf = ""
			empnamemaf = ""
			dategenmaf = ""
			empdeptmaf = ""
			empsalutationmaf = ""
		end if
		oRs.Close
		set oRs = nothing
		oConn.Close
		set oConn = nothing
	end if
end sub
%>
