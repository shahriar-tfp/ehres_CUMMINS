<%

Response.Buffer = true

	   Set webdb = Server.CreateObject("ADODB.Connection")
	   		webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=SA;PWD=;Initial catalog=HRDB_SNE;Server=(local);Connect Timeout=900000"
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	   Set webdbCommand = Server.CreateObject("ADODB.Command")
	   Set webdbCommand.ActiveConnection = webdb
	   dim RQID
	   RQID = request.Form("cboRQID")
	   ssql = "exec sp_WRQ_GetHire '" & RQID & "','" & request.Form("EMPID") & "'"
	   webdbCommand.CommandText = ssql
	   webdb.Execute webdbCommand.CommandText
		
	   response.Redirect "PersonalProfileVw.asp"

%>