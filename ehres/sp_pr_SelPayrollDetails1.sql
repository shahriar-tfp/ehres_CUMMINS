if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_pr_SelPayrollDetails1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_pr_SelPayrollDetails1]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE proc sp_pr_SelPayrollDetails1    
@OrganID varchar(20),    
@IDName varchar(80),    
@year int,    
@month int,    
@Type varchar(10)    
AS    
declare @tempempdbit varchar(20),    
        @tempempcredit varchar(20),    
        @COUNT AS INT    
    
    
if @Type = 'NAME'    
  select empid,empname    
  from is_emppersonal    
  where empid = @IDName     
    
if @Type = 'ALLOWANCE'    
    
  select @count =count(*)    
  from pr_payrolldetails ,sa_allowanceType --,is_emppersonal    
  where pr_payrolldetails.regisno = @OrganID    
  and   pr_payrolldetails.regisno = sa_allowanceType.regisno    
  and   pr_payrolldetails.empid                     = @IDName    
  and   year                      = @year    
  and   month                     = @month    
  and   processid                 = allowanceid    
  and   companycredit             = 0     
    
if @count >0
begin    
  --SELECT @COUNT = COUNT(*)    
  select description,--is_emppersonal.empid,is_emppersonal.empname,     
         sum(empdebit) as 'empdebit',    
         sum(empcredit) as 'empcredit',    
         grouping,    
         ''
  from pr_payrolldetails ,sa_allowanceType --,is_emppersonal    
  where pr_payrolldetails.regisno = @OrganID    
  and   pr_payrolldetails.regisno = sa_allowanceType.regisno    
  and   pr_payrolldetails.empid                     = @IDName    
  and   year                      = @year    
  and   month                     = @month    
  and   processid                 = allowanceid    
  and   companycredit             = 0    
  --and is_emppersonal.empid = @IDName
 group by description,grouping      
  union   
  select description,--is_emppersonal.empid,is_emppersonal.empname,      
         sum(empdebit) as 'empdebit',    
         sum(empcredit) as 'empcredit', 
         grouping,    
         '' 
  from pr_payrolldetails, sa_DeductionType --,--is_emppersonal    
  where pr_payrolldetails.regisno =  @OrganID    
  and   pr_payrolldetails.regisno =  sa_DeductionType.regisno    
  and   processid                 =  deductionid    
  and   pr_payrolldetails.empid                     = @IDName    
  and   year                      = @year    
  and   month                     = @month    
  and   deductionid               <> '~~EPFER'     
  and   deductionid               <> '~~SOCSOER'    
  and   companycredit             = 0    
  --and is_emppersonal.empid = @IDName 
  group by description,grouping   
  order by grouping
  
end    
else    
  select description = '',--is_emppersonal.empid,is_emppersonal.empname,      
         empdebit = convert(decimal, 0),            
         empcredit = convert(decimal,0),    
         grouping = ''    
      
/*else    
if @Type = 'DEDUCTION'    
  select description, empcredit    
  from pr_payrolldetails, sa_DeductionType    
  where pr_payrolldetails.regisno =  @OrganID    
  and   pr_payrolldetails.regisno =  sa_DeductionType.regisno    
  and   processid                 =  deductionid    
  and   empid                     = @IDName    
  and   year                      = @year    
  and   month                     = @month    
  and   deductionid               <> "~~EPFER"      
  and   deductionid               <> "~~SOCSOER"    
  order by empid, deductionid    
else    
if @Type = 'ID'    
  select pr_payrollHeader.empid, empname, description, convert(varchar,hrprate)    
  from pr_payrollHeader, is_emppersonal, sa_reference    
  where regisno     = @OrganID    
  and   pr_payrollHeader.empid       = @IDName    
  and   pr_payrollHeader.empid = is_emppersonal.empid    
  and   year        = @year    
  and   month       = @month    
  and   paymode     = referenceid    
  and   type        = 'PAYMODE'    
else    
     
if @Type = 'SRCID'    
   Select empid    
     From is_empstatus    
     where regisno     =   @OrganID    
     And   empID like @IDName    
     Order By empID    
else    
IF @Type = 'EMPNAME'    
   Select empname    
     From is_emppersonal, is_empstatus    
     where is_empstatus.regisno     =   @OrganID    
     And   is_emppersonal.empID = @IDName    
     And   is_emppersonal.empID = is_empstatus.empID    
     Order By is_emppersonal.empID */    
    
    
    
    
  


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

