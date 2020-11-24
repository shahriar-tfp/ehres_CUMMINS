
<% 

sub inputsession(vtempleavetypeid,vtempdate1,vtempdate2,vtempstatus)
      
    if vtempdate1 ="" and vtempdate2 ="" then 
       vtempdate1 ="01/01/1900"
       vtempdate2 ="01/01/1900"
    end if 
    if vtempStatus = "" then 
       vtempStatus = "P"
    else
       vtempStatus=vtempstatus   
    end if     
	Session("ssLeavetypeid") =vtempleavetypeid
	Session("ssDate1") =vtempdate1
	Session("ssDate2") =vtempdate2
	Session("ssStatus")=vtempstatus
  end sub
  
  sub inputsessionLeave(vtempstatus,vtempdate1,vtempdate2,vtempsearchby,vtempid)
    
   if vtempdate1 ="" and vtempdate2 ="" then 
       vtempdate1 ="01/01/1900"
       vtempdate2 ="01/01/1900"
   end if 
   if vtempStatus = "" then 
       vtempStatus ="P"
   else
       vtempStatus=vtempstatus   
   end if      
   session("ssStatus")= vtempstatus
   session("ssDate1") = vtempdate1
   session("ssDate2") = vtempdate2
   session("ssSearchBy")= vtempsearchby
   session("ssID")= vtempid
   end sub
   
   sub inputsessionApp(vtempstatus,vtempdate1,vtempdate2,vtempID,vtempSearchBy)
   
   if vtempStatus = "" then 
       vtempStatus ="P"
   else
       vtempStatus=vtempstatus   
   end if      
   if vtempdate1 ="" and vtempdate2 ="" then 
       vtempdate1 ="01/01/1900"
       vtempdate2 ="01/01/1900"
   end if 
      session("ssStatus") = vtempstatus
      session("ssDate1") = vtempdate1
      session("ssdate2") = Vtempdate2
      session("ssID") = vtempID
      session("ssSearchBy") = vtempSearchBy
    
   end sub    
   
   sub inputleavebal(vtempdate1,vtempdate2,vtempcboempid)
   
   if vtempdate1 ="" and vtempdate2 ="" then 
       vtempdate1 =""
       vtempdate2 =""
   end if 
	  session("ssDate1lv") = vtempdate1
	  session("ssDate2lv") = vtempdate2
	  session("sscboempidlv") = vtempcboempid
	  
   end sub
   
   sub inputleavebal1(vtempdate1,vtempdate2,vtempcboempid)
   
   if vtempdate1 ="" and vtempdate2 ="" then 
       vtempdate1 =""
       vtempdate2 =""
   end if 
	  session("ssDate1lv1") = vtempdate1
	  session("ssDate2lv1") = vtempdate2
	  session("sscboempidlv1") = vtempcboempid
	  
   end sub
   sub inputtmserror(vtempempid,vtemptypeerror,vtempdate1,vtempdate2)
     
      if vtemptypeerror = "" then
         vtemptypeerror = "ALL" 
      end if   
	  session("ssempiderror") = vtempempid
	  session("sserrorerror") = vtemptypeerror
	  session("ssdate1error") = vtempdate1
	  session("ssdate2error") = vtempdate2
	  'session("sstemp") = vtemp
	  
  end sub	
  
  sub ssleaveapp2(vtempyear, vtempstatus)
	session("ssyear") = vtempyear
	session("sssel") = vtempstatus
	
  end sub 	  	   	
%>	

