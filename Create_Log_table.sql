SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[sys_email_download_log](
	[job_id] [int] NOT NULL,
	[job_time] [datetime2](7) NOT NULL,
	[attm_count] [int] NOT NULL,
	[id] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
