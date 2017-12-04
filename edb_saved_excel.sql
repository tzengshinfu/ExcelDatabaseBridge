/****** Object:  Table [dbo].[edb_saved_excel]    Script Date: 2017/12/1 ¤W¤È 10:21:26 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[edb_saved_excel]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[edb_saved_excel](
	[host_name] [nvarchar](20) NOT NULL,
	[file_path] [nvarchar](255) NOT NULL,
	[book_name] [nvarchar](255) NOT NULL,
	[sheet_name] [nvarchar](50) NOT NULL,
	[user_account] [nvarchar](50) NOT NULL,
	[modified_datetime] [datetime] NOT NULL,
	[row] [int] NOT NULL,
	[col] [int] NOT NULL,
	[value] [nvarchar](4000) NULL,
	[format] [varchar](50) NULL,
	[formula] [nvarchar](1000) NULL,
 CONSTRAINT [PK_edb_saved_excel] PRIMARY KEY CLUSTERED 
(
	[modified_datetime] ASC,
	[row] ASC,
	[col] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
END
GO

SET ANSI_PADDING OFF
GO

ALTER AUTHORIZATION ON [dbo].[edb_saved_excel] TO  SCHEMA OWNER 
GO

SET ANSI_PADDING ON

GO

/****** Object:  Index [IX_edb_saved_excel]    Script Date: 2017/12/1 ¤W¤È 10:21:26 ******/
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[edb_saved_excel]') AND name = N'IX_edb_saved_excel')
CREATE UNIQUE NONCLUSTERED INDEX [IX_edb_saved_excel] ON [dbo].[edb_saved_excel]
(
	[modified_datetime] ASC,
	[host_name] ASC,
	[file_path] ASC,
	[book_name] ASC,
	[sheet_name] ASC,
	[row] ASC,
	[col] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO


