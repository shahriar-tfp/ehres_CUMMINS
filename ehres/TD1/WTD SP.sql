if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_WTD_insUpdDelGapAnalysis]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_WTD_insUpdDelGapAnalysis]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create proc sp_WTD_insUpdDelGapAnalysis
@regisno    	varchar(10),
@empid      	varchar(10),
@positionid 	varchar(10),
@comsubid   	varchar(10),
@tolscore   	int,
@Action     	varchar(5),
@Type	    	varchar(20)

as

declare @training bit,
@Performancemgt bit

select @Training = training, @performancemgt = performancemgt
From td_compscore
     Where regisno     = @regisno
     And   positionid  = @positionid
     And   comsubid    = @comsubid

If @Type = 'CURRENT' And @Action = 'ADD'
   Insert Into td_currentgapanalysis values 
	(@regisno, @empid, @positionid, @comsubid, @tolscore, @training, @performancemgt)

Else
If @Type = 'CURRENT' And @Action = 'DEL'
   Delete From td_currentgapanalysis 
   Where regisno    = @regisno
   And   empid      = @empid
   And   positionid = @positionid
   And   Comsubid   = @Comsubid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_Wtd_insUpdEmpEvalution]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_Wtd_insUpdEmpEvalution]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create proc sp_Wtd_insUpdEmpEvalution
@Regisno varchar(20),
@Empid varchar(10),
@Comp varchar(20),
@Evaulate varchar(20),
@score int

as

if exists(select * 
from td_EmpEvaluation
where regisno = @Regisno
and empid = @Empid
and competencyID = @Comp
and EvaluationID = @Evaulate)
delete from td_EmpEvaluation
where regisno = @Regisno
and empid = @Empid
and competencyID = @Comp
and EvaluationID = @Evaulate

insert into td_EmpEvaluation
select @Regisno, @Empid, @Comp, @Evaulate, @Score


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



