<%

function FormatDate (vDate,formatString)
dim userDate
dim sDate
dim msg
dim validChars
dim i
dim aDay
dim sDay
dim aMonth
dim sMonth
dim aYear
dim sYear
dim sMonths
dim arrMonths
dim dayProcessed
dim monthProcessed
dim yearProcessed
sMonths = "JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC"
arrMonths = split(sMonths, ",")
validChars = "mdy/, "
do

      userDate = vDate
      if len(userDate) = 0 then
         msg = "You must enter a date."
         exit do
      end if
      if not isDate(userDate) then
         msg = "The value you entered is not a valid date."
         exit do
      end if
      if len(formatString) = 0 then 
         msg = "You must enter a format string"
         exit do
      end if
      formatString = lcase(formatstring)
      for i = 1 to len(formatString)
         if instr(1, validChars, mid(formatString, i, 1), vbTextCompare) = 0 then
            msg = "Your format string contains invalid characters. You may only use the characters 'd' for Day, 'm' for Month, and 'y' for year."
            exit do
         end if
      next

   aDay = day(userDate)
   aMonth = month(userDate)
   aYear = year(userDate)
   if instr(formatString, "dd") then
      ' include the day in two-digit, leading zero form
      if aDay < 10 then
         sDay = "0" & cstr(aDay)
      else
         sDay = cstr(aDay)
      end if
   elseif instr(formatString, "d") then
      ' include the day without leading zeroes
      sDay = cstr(aDay)
   end if
   if instr(formatString, "mmm") then
      ' include the month as a three-character abbreviation
      ' find the month abbreviation
      sMonth = arrMonths(aMonth - 1)
   elseif instr(formatString, "mm") then
      ' format the month in two-digit, leading zero form
      if aMonth < 10 then
         sMonth = "0" & cstr(aMonth)
      else
         sMonth = cstr(aMonth)
      end if
   elseif instr(formatString, "m") then
      ' include the month without leading zeroes
      sMonth = cstr(aMonth)
   end if
   if instr(formatString, "yyyy") then
      sYear = cstr(aYear)
   elseif instr(formatString, "yy") then
      sYear = mid(aYear, 3)
   end if
   
   ' now put the elements in position, with separators if included
   for i = 1 to len(formatstring)
      select case mid(formatString, i, 1)
      case "d"
         if not dayProcessed then
            sDate = sDate & sDay
            dayProcessed=true
         end if
      case "m"
         if not monthProcessed then
            sDate = sDate & sMonth
            monthProcessed = true
         end if       
      case "y"
         if not yearProcessed then
            sDate = sDate & sYear
            yearProcessed=true
         end if
      case else
         sDate = sDate & mid(formatString, i, 1)
      end select
   next
   FormatDate = sDate
   msg=""
loop while False

end function
%>