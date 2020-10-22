<%

Response.Buffer = true

	   Set webdb = Server.CreateObject("ADODB.Connection")
	   		webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=sa;PWD=;Initial catalog=HRDB_SNE;Data Source=HRDBSERVER\HRDB;Connect Timeout=900000"
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	   Set webdbCommand = Server.CreateObject("ADODB.Command")
	   Set webdbCommand.ActiveConnection = webdb
	   
	   dim rowcount
	   dim maxrow
	   
	   maxrow = Request.Form("txtRowNo")
	   
	   if isnumeric(maxrow) then
		do until rowcount = cint(maxrow)
	      rowcount = rowcount + 1
			
			ssql = "Exec sp_WRQ_insQA '" & Request.form("RQID") & "' , '" _
	      	            & request.Form("Eva" + cstr(rowcount)) & "','" _
	      	            & request.Form("S" + cstr(rowcount)) & "',0,'DEL'"
			webdbCommand.CommandText = ssql
			webdb.Execute webdbCommand.CommandText
			 
			ssql = "Exec sp_WRQ_insQA '" & Request.form("RQID") & "' , '" _
	      	            & request.Form("Eva" + cstr(rowcount)) & "','" _
	      	            & request.Form("S" + cstr(rowcount)) & "',0,'ADD'"
			webdbCommand.CommandText = ssql
			webdb.Execute webdbCommand.CommandText
		
		loop
      end if
      
      response.redirect "RQ_question.asp?RQID=" + Request.form("RQID")

%>