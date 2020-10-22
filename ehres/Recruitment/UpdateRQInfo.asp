<%

Response.Buffer = true

	   Set webdb = Server.CreateObject("ADODB.Connection")
	   		webdb.Open "Provider=SQLOLEDB.1;Persist Security Info=False;UID=SA;PWD=;Initial catalog=HRDB_SNE;Server=(local);Connect Timeout=900000"
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	   Set webdbCommand = Server.CreateObject("ADODB.Command")
	   Set webdbCommand.ActiveConnection = webdb
	   
	   dim rowcount
	   dim RQID
	   dim maxrow1
	   dim maxrow2
	   dim maxrow3
	   dim maxrow4
	   
	   maxrow1 = Request.Form("txtRowNo1")
	   maxrow2 = Request.Form("txtRowNo2")
	   maxrow3 = Request.Form("txtRowNo3")
	   maxrow4 = Request.Form("txtRowNo4")
	   
	   ssql = "Exec sp_WRQ_GrenerateID"
	   webdbCommand.CommandText = ssql
	   webdbRecordset.Open webdbCommand,,1 , 3
	   
	   if not webdbrecordset.eof then
			RQID = trim(webdbRecordset.Fields("ID"))
	   else
			RQID = "RQ99999"
	   end if
	   
	   if isnumeric(maxrow1) then
		do until rowcount = cint(maxrow1)
	      rowcount = rowcount + 1
			 
			ssql = "exec sp_WRQ_InsRQProfile '" & RQID & "','1','" & request.Form("PH" + cstr(rowcount)) & "','" & request.Form("P" + cstr(rowcount)) & "'"
			webdbCommand.CommandText = ssql
			webdb.Execute webdbCommand.CommandText
		
		loop
      end if
      
      rowcount = 0
      if isnumeric(maxrow2) then
		do until rowcount = cint(maxrow2)
	      rowcount = rowcount + 1
			 
			ssql = "exec sp_WRQ_InsRQProfile '" & RQID & "','2','" & request.Form("EH" + cstr(rowcount)) & "','" & request.Form("E" + cstr(rowcount)) & "'"
			webdbCommand.CommandText = ssql
			webdb.Execute webdbCommand.CommandText
		
		loop
      end if
      
      rowcount = 0
      if isnumeric(maxrow3) then
		do until rowcount = cint(maxrow3)
	      rowcount = rowcount + 1
			 
			ssql = "exec sp_WRQ_InsRQProfile '" & RQID & "','3','" & request.Form("EXH" + cstr(rowcount)) & "','" & request.Form("EX" + cstr(rowcount)) & "'"
			webdbCommand.CommandText = ssql
			webdb.Execute webdbCommand.CommandText
		
		loop
      end if
      
      rowcount = 0
      if isnumeric(maxrow4) then
		do until rowcount = cint(maxrow4)
	      rowcount = rowcount + 1
			 
			ssql = "exec sp_WRQ_InsRQProfile '" & RQID & "','4','" & request.Form("AH" + cstr(rowcount)) & "','" & request.Form("A" + cstr(rowcount)) & "'"
			webdbCommand.CommandText = ssql
			webdb.Execute webdbCommand.CommandText
		
		loop
      end if
      
      response.redirect "RQ_question.asp?RQID=" + RQID

%>