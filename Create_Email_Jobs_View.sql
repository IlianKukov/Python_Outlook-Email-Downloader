SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER VIEW [dbo].[v_sys_email_download]
AS
SELECT[mbox]
      ,[subj]
      ,[save_path]
      ,[job_name]
      ,[args]
	  ,[id]
  FROM [dbo].[sys_email_download]
  WHERE [on]=1

GO
