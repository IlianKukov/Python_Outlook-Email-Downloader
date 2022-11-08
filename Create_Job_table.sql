SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[sys_email_download](
	[mbox] [nvarchar](250) NULL,
	[subj] [nvarchar](250) NULL,
	[save_path] [nvarchar](max) NULL,
	[job_name] [nvarchar](250) NULL,
	[args] [nvarchar](max) NULL,
	[client_name] [nvarchar](250) NULL,
	[on] [int] NULL,
	[id] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

