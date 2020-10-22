<%

Response.Buffer = true

	   Set webdb = Server.CreateObject("ADODB.Connection")
	   		webdb.Open Session("ConnectStr")
	   Set webdbRecordset = Server.CreateObject("ADODB.Recordset")
	   Set webdbCommand = Server.CreateObject("ADODB.Command")
	   Set webdbCommand.ActiveConnection = webdb
	   
	   dim rowcount
	   dim maxrow
	   
	   maxrow = Request.Form("txtRowNo")
	   
	   if isnumeric(maxrow) then
		do until rowcount = cint(maxrow)
	      rowcount = rowcount + 1
			 
			ssql = "Exec sp_Wtd_insUpdEmpevaluation '" & Session("Regisno") & "', '" & Request.form("CboEmpID") & "' , '" _
	      	            & Request.form("txtAction") & "','" & request.Form("Eva" + cstr(rowcount)) & "','" _
	      	            & request.Form("posID") & "'," & request.Form("S" + cstr(rowcount)) & ",'TARGET'"
			webdbCommand.CommandText = ssql
			webdb.Execute webdbCommand.CommandText
		
		loop
      end if
      
      ssql = "Exec sp_WTD_insUpdDelGapAnalysis '" & session("regisno") & "','" & request.Form("CboEmpid") & "','" _
					& request.Form("txtTPosID") & "','" & request.Form("txtAction") & "'," & request.Form("tolscore") & ",'Del','TARGET'"
      webdbCommand.CommandText = ssql
	  webdb.Execute webdbCommand.CommandText
      
      ssql = "Exec sp_WTD_insUpdDelGapAnalysis '" & session("regisno") & "','" & request.Form("CboEmpid") & "','" _
					& request.Form("txtTPosID") & "','" & request.Form("txtAction") & "'," & request.Form("tolscore") & ",'ADD','TARGET'"
	  webdbCommand.CommandText = ssql
	  webdb.Execute webdbCommand.CommandText
	  
      response.redirect "TD_GapAnalysis2.asp"

%>