USE [CIB]
GO
/****** Object:  Table [dbo].[Com_i]    Script Date: 11/4/2018 5:49:35 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Com_i](
	[com_id] [int] IDENTITY(1,1) NOT NULL,
	[cib_s_c] [varchar](150) NULL,
	[name_of_owner] [varchar](150) NULL,
	[role] [varchar](150) NULL,
	[fi] [varchar](150) NULL,
	[legal] [varchar](150) NULL,
	[cib_bb_id] [int] NULL,
	[stay_order] [varchar](150) NULL,
 CONSTRAINT [PK_companylist] PRIMARY KEY CLUSTERED 
(
	[com_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[D_Contract_History]    Script Date: 11/4/2018 5:49:35 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[D_Contract_History](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[Date] [varchar](150) NULL,
	[Outstanding] [varchar](150) NULL,
	[Overdue] [varchar](150) NULL,
	[NPI] [varchar](150) NULL,
	[Status] [varchar](150) NULL,
	[Defa] [varchar](150) NULL,
	[D_id] [int] NULL,
 CONSTRAINT [PK_I_D_Contract_History] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[d_Other_sub_linked]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[d_Other_sub_linked](
	[t_id] [int] IDENTITY(1,1) NOT NULL,
	[CIB_s_c] [varchar](50) NULL,
	[Role] [varchar](50) NULL,
	[Name] [varchar](50) NULL,
	[D_id] [int] NULL,
 CONSTRAINT [PK_I_d_Other_sub_linked] PRIMARY KEY CLUSTERED 
(
	[t_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DETAILS_OF_INSTALL_Faca]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DETAILS_OF_INSTALL_Faca](
	[D_id] [int] IDENTITY(1,1) NOT NULL,
	[Ref] [varchar](50) NULL,
	[FI_code] [varchar](150) NULL,
	[Branch_code] [varchar](150) NULL,
	[CIB_contract_code] [varchar](150) NULL,
	[FI_contract_code] [varchar](150) NULL,
	[Role] [varchar](150) NULL,
	[Phase] [varchar](150) NULL,
	[Facility] [varchar](150) NULL,
	[Starting_date] [varchar](50) NULL,
	[End_date_of_contract] [varchar](150) NULL,
	[Sanction_Limit] [varchar](150) NULL,
	[Total_Disbursement_Amount] [money] NULL,
	[Total_number_of_installments] [varchar](150) NULL,
	[Installment_Amount] [money] NULL,
	[Remaining_installments_Number] [varchar](150) NULL,
	[Security_Amount] [money] NULL,
	[Third_Party_guarantee_Amount] [varchar](150) NULL,
	[Security_Type] [varchar](150) NULL,
	[Date_of_Last_Update] [varchar](150) NULL,
	[Date_of_Law_suit] [varchar](150) NULL,
	[Date_of_Last_payment] [varchar](50) NULL,
	[Date_of_classification] [varchar](50) NULL,
	[Date_of_last_rescheduling] [varchar](50) NULL,
	[Method_of_payment] [varchar](150) NULL,
	[Payments_periodicity] [varchar](150) NULL,
	[Number_of_time_rescheduled] [varchar](150) NULL,
	[Remaining_installments_Amount] [varchar](150) NULL,
	[Reorganized_credit] [varchar](150) NULL,
	[Basis_for_classification_qualitative_judgment] [varchar](150) NULL,
	[Remarks] [varchar](500) NULL,
	[cib_bb_id] [int] NULL,
 CONSTRAINT [PK_I_DETAILS_OF_INSTALL_Faca] PRIMARY KEY CLUSTERED 
(
	[D_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Flags]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Flags](
	[flag_id] [int] NOT NULL,
	[table_name] [varchar](50) NULL,
	[frist_name] [varchar](50) NULL,
	[frist_flag] [int] NULL,
	[second_name] [varchar](50) NULL,
	[second_flag] [int] NULL,
	[third_name] [varchar](50) NULL,
	[third_flag] [int] NULL,
 CONSTRAINT [PK_Flags] PRIMARY KEY CLUSTERED 
(
	[flag_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[I_ADDRESS]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[I_ADDRESS](
	[Address_id] [int] IDENTITY(1,1) NOT NULL,
	[Address_Type] [varchar](150) NULL,
	[Address] [varchar](150) NULL,
	[Postal_code] [varchar](150) NULL,
	[District] [varchar](150) NULL,
	[Country] [varchar](150) NULL,
	[flag] [int] NULL,
	[cib_bb_id] [int] NOT NULL,
 CONSTRAINT [PK_I_ADDRESS] PRIMARY KEY CLUSTERED 
(
	[Address_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[I_INQUIRED]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[I_INQUIRED](
	[Inque_id] [int] IDENTITY(1,1) NOT NULL,
	[Trade_name] [varchar](150) NULL,
	[Proprietorship_District] [varchar](150) NULL,
	[Proprietorship_Address] [varchar](150) NULL,
	[Owner_Name] [varchar](150) NULL,
	[Father_name] [varchar](150) NULL,
	[Mother_name] [varchar](150) NULL,
	[DOB] [varchar](50) NULL,
	[Proprietorship_Postalcode] [varchar](150) NULL,
	[NID] [varchar](150) NULL,
	[Owner_Address] [varchar](150) NULL,
	[Postcode] [varchar](150) NULL,
	[District] [varchar](150) NULL,
	[TIN] [varchar](150) NULL,
	[cib_bb_id] [int] NULL,
 CONSTRAINT [PK_I_INQUIRED] PRIMARY KEY CLUSTERED 
(
	[Inque_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IMaster]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IMaster](
	[cib_bb_id] [int] IDENTITY(1,1) NOT NULL,
	[CIB_subject_code] [varchar](50) NULL,
	[Date_of_Inquiry] [datetime] NULL,
	[User_ID] [varchar](50) NULL,
	[FI_Code] [varchar](50) NULL,
	[Branch_Code] [varchar](50) NULL,
	[FI_Name] [varchar](50) NULL,
	[file_location] [varchar](150) NULL,
	[Upload_date] [varchar](50) NULL,
 CONSTRAINT [PK__IMaster__7C8D7D293D0CDB8F] PRIMARY KEY CLUSTERED 
(
	[cib_bb_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[owner_list]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[owner_list](
	[ow_id] [int] IDENTITY(1,1) NOT NULL,
	[cib_sub] [varchar](150) NULL,
	[Name_owner] [varchar](150) NULL,
	[Role] [varchar](150) NULL,
	[stay_order] [varchar](150) NULL,
	[cib_bb_id] [int] NULL,
 CONSTRAINT [PK_owner_list] PRIMARY KEY CLUSTERED 
(
	[ow_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PROP_CONCERN]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PROP_CONCERN](
	[linked_id] [int] IDENTITY(1,1) NOT NULL,
	[CIb_sub_Code] [varchar](50) NULL,
	[Sector_type] [varchar](50) NULL,
	[Sector_code] [varchar](50) NULL,
	[Ref_number] [varchar](50) NULL,
	[Trade_Name] [varchar](50) NULL,
	[Tele_number] [varchar](50) NULL,
	[cib_bb_id] [int] NULL,
 CONSTRAINT [PK_I_PROP_CONCERN] PRIMARY KEY CLUSTERED 
(
	[linked_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[REQUESTED_CONTRACT]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[REQUESTED_CONTRACT](
	[Req_id] [int] IDENTITY(1,1) NOT NULL,
	[SL] [int] NULL,
	[Type_of_Contract] [varchar](50) NULL,
	[Facility] [varchar](50) NULL,
	[Phase] [varchar](50) NULL,
	[Role] [varchar](50) NULL,
	[FI_Code] [varchar](50) NULL,
	[Branch_Code] [varchar](50) NULL,
	[Request_date] [varchar](50) NULL,
	[Total_Requested_Amount] [money] NULL,
	[CIB_subject_code] [varchar](50) NULL,
	[CIB_contract_code] [varchar](50) NULL,
	[FI_0contract_codede] [varchar](50) NULL,
	[cib_bb_id] [int] NULL,
 CONSTRAINT [PK_I_REQUESTED_CONTRACT] PRIMARY KEY CLUSTERED 
(
	[Req_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Sub_ INFO]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sub_ INFO](
	[Sub_id] [int] IDENTITY(1,1) NOT NULL,
	[CIB_subject_code] [varchar](50) NULL,
	[Title_Name] [varchar](500) NULL,
	[Fathername] [varchar](500) NULL,
	[SpouseName] [varchar](500) NULL,
	[Mothername] [varchar](500) NULL,
	[Dob] [date] NULL,
	[Gender] [varchar](500) NULL,
	[District_Country] [varchar](500) NULL,
	[NID] [varchar](500) NULL,
	[TIN] [varchar](500) NULL,
	[Type_of_sub] [varchar](500) NULL,
	[Ref_number] [varchar](500) NULL,
	[Sector_type] [varchar](500) NULL,
	[ID_type] [varchar](500) NULL,
	[ID_number] [varchar](500) NULL,
	[ID_issue_date] [varchar](500) NULL,
	[ID_issue_country] [varchar](500) NULL,
	[Telephone] [varchar](500) NULL,
	[Remarks] [varchar](500) NULL,
	[cib_bb_id] [int] NULL,
	[trade_name] [varchar](500) NULL,
	[sector_code] [varchar](500) NULL,
	[legal_form] [varchar](500) NULL,
	[reg_num] [varchar](500) NULL,
	[reg_date] [varchar](500) NULL,
 CONSTRAINT [PK_I_Sub_ INFO] PRIMARY KEY CLUSTERED 
(
	[Sub_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SUM_OF_FACILITY_S_AS_BOR]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SUM_OF_FACILITY_S_AS_BOR](
	[Bor_id] [int] IDENTITY(1,1) NOT NULL,
	[No_of_reporting_Institutes] [int] NULL,
	[No_of_Living_Contracts] [int] NULL,
	[Total_Outstanding_Amount] [money] NULL,
	[Total_Overdue_Amount] [money] NULL,
	[No_of_Stay_order_contracts] [int] NULL,
	[Total_Outstanding_amount_for_Stay] [money] NULL,
	[cib_bb_id] [int] NULL,
	[flag] [int] NULL,
 CONSTRAINT [PK_I_SUM_OF_FACILITY_S_AS_BOR] PRIMARY KEY CLUSTERED 
(
	[Bor_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SUM_OF_FUNDED_FACILI_AS_BOR]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SUM_OF_FUNDED_FACILI_AS_BOR](
	[co_id] [int] IDENTITY(1,1) NOT NULL,
	[Contract_Category] [varchar](50) NOT NULL,
	[UC_NO] [int] NULL,
	[SMA_NO] [int] NULL,
	[SS_NO] [int] NULL,
	[DF_NO] [int] NULL,
	[B_NO] [int] NULL,
	[BLW_NO] [int] NULL,
	[Terminated_NO] [int] NULL,
	[Requested_NO] [int] NULL,
	[Stay_Order_NO] [int] NULL,
	[UC_Amount] [money] NULL,
	[SMA_Amount] [money] NULL,
	[SS_Amount] [money] NULL,
	[DF_Amount] [money] NULL,
	[BL_Amount] [money] NULL,
	[BLW_Amount] [money] NULL,
	[Terminated_Amount] [money] NULL,
	[Requested_Amount] [money] NULL,
	[Stay_Order_Amount] [money] NULL,
	[cib_bb_id] [int] NOT NULL,
	[flag] [int] NULL,
 CONSTRAINT [PK_I_1a_SUM_OF_FUNDED_FACILI_AS_BOR] PRIMARY KEY CLUSTERED 
(
	[co_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SUM_OF_NON_FUNDED_FACILI_AS_BOR]    Script Date: 11/4/2018 5:49:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SUM_OF_NON_FUNDED_FACILI_AS_BOR](
	[Bor_id] [int] IDENTITY(1,1) NOT NULL,
	[Type_of_Financing] [varchar](50) NULL,
	[Living_NO] [int] NULL,
	[Terminated_NO] [int] NULL,
	[Requested_NO] [int] NULL,
	[Stay_Order_NO] [int] NULL,
	[Living_Amount] [money] NULL,
	[Terminated_Amount] [money] NULL,
	[Requested_Amount] [money] NULL,
	[Stay_Order_Amount] [money] NULL,
	[cib_bb_id] [int] NOT NULL,
	[flag] [int] NULL,
 CONSTRAINT [PK_I_ 1b_SUM_OF_NON_FUNDED_FACILI_AS_BOR] PRIMARY KEY CLUSTERED 
(
	[Bor_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Sub_ INFO]  WITH CHECK ADD  CONSTRAINT [FK_I_Sub_ INFO_I_Sub_ INFO] FOREIGN KEY([Sub_id])
REFERENCES [dbo].[Sub_ INFO] ([Sub_id])
GO
ALTER TABLE [dbo].[Sub_ INFO] CHECK CONSTRAINT [FK_I_Sub_ INFO_I_Sub_ INFO]
GO
