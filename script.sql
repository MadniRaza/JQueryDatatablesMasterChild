USE [cwdw]
GO
/****** Object:  Table [dbo].[zVendor_Attributes]    Script Date: 2/24/2019 10:46:33 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[zVendor_Attributes](
	[RECNUM] [int] IDENTITY(1,1) NOT NULL,
	[Company] [varchar](8) NOT NULL,
	[Plant] [varchar](8) NOT NULL,
	[PartNum] [nvarchar](50) NOT NULL,
	[VendorID] [nvarchar](8) NOT NULL,
	[VenOrder] [int] NULL,
	[Alloc] [decimal](5, 2) NULL,
	[UnitPrice] [numeric](17, 5) NULL,
	[VendorLT] [int] NULL,
	[Sourcing] [char](1) NOT NULL,
	[StdDevLT] [int] NULL,
	[RecLT] [int] NULL,
	[Status] [char](1) NOT NULL,
	[InsertedBy] [varchar](50) NULL,
	[InsertedOn] [datetime] NOT NULL,
	[UpdatedBy] [varchar](50) NULL,
	[UpdatedOn] [datetime] NULL,
	[RoleCode] [varchar](100) NULL,
 CONSTRAINT [PK_zVendor_Attributes] PRIMARY KEY CLUSTERED 
(
	[Company] ASC,
	[PartNum] ASC,
	[VendorID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[zVendor_Attributes] ADD  CONSTRAINT [DF_zVendor_Attributes_Sourcing]  DEFAULT ('N') FOR [Sourcing]
GO
ALTER TABLE [dbo].[zVendor_Attributes] ADD  CONSTRAINT [DF_zVendor_Attributes_Status]  DEFAULT ('N') FOR [Status]
GO
ALTER TABLE [dbo].[zVendor_Attributes] ADD  CONSTRAINT [DF_Table_1_InsertOn]  DEFAULT (getdate()) FOR [InsertedOn]
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Purchase Quantity Per' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'zVendor_Attributes', @level2type=N'COLUMN',@level2name=N'Alloc'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Usage Quantity Per' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'zVendor_Attributes', @level2type=N'COLUMN',@level2name=N'UnitPrice'
GO
