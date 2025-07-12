/****** Object:  StoredProcedure [dbo].[SBO_SP_TransactionNotification]    Script Date: 4/29/2020 8:32:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[SBO_SP_TransactionNotification] 

@object_type nvarchar(30), 				-- SBO Object Type
@transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
@num_of_cols_in_key int,
@list_of_key_cols_tab_del nvarchar(255),
@list_of_cols_val_tab_del nvarchar(255)

AS

begin

-- Return values
declare @error  int				-- Result (0 for no error)
declare @error_message nvarchar (200) 		-- Error string to be displayed
select @error = 0
select @error_message = N'Ok'

--------------------------------------------------------------------------------------------------------------------------------

declare @islocked int
declare @count int

-----------P4S Code block----------------------------------------------------------------------------------------------------
select @islocked = 0
if @object_type = '4' or @object_type = '17' or @object_type = '22' or @object_type = '1250000001' or @object_type = '202' or @object_type = '13'
begin
	select @islocked = U_Lock from [@P4S_Export] where U_ObjType = @object_type and U_DocId = @list_of_cols_val_tab_del
	select @count = count(*) from [@P4S_Export] where U_ObjType = @object_type and U_DocId = @list_of_cols_val_tab_del
	if(@count = 0)
	begin
		insert into [@P4S_Export]
			(Code, Name, U_DocId, U_BackOrder, U_ExportFlag, U_Lock, U_ObjType, U_Error, U_DateCreated, U_DateModified)
		values
			(concat(@object_type, @list_of_cols_val_tab_del), concat(@object_type, @list_of_cols_val_tab_del),
			@list_of_cols_val_tab_del,'0', '0', '0', @object_type, '', getdate(), getdate())
		select @islocked = 0
	end

	if (@islocked = 0)
 	begin
 		update [@P4S_Export] set U_ExportFlag = 0
		where U_ObjType = @object_type and U_DocId = @list_of_cols_val_tab_del		
	end
end

if @object_type = '17' and @transaction_type in ('A', 'U')
begin
	if ((select top 1 'TRUE' from rdr1 x inner join oitm y on x.itemcode = y.ItemCode where x.docentry = @list_of_cols_val_tab_del and (y.U_P4S_IsDecimal <> '1' or y.U_P4S_IsDecimal is null) and (x.Quantity-ROUND(x.quantity,0,1)) <>0)) = 'TRUE'
	begin 
		set @error = '444001'
		set @error_message = 'Decimal controlled incorrect! Verify Order ' + @list_of_cols_val_tab_del + ' Item  ' + (select top 1  w.itemcode from rdr1 w inner join oitm z on w.ItemCode = z.ItemCode where w.docentry = @list_of_cols_val_tab_del and (z.U_P4S_IsDecimal in ('0', '') or isnull(z.U_P4S_IsDecimal,9) = 9) and (w.Quantity-ROUND(w.quantity,0,1)) <> 0)
	end
end

if (@islocked <> 0)
begin
	select @error = 1
	select @error_message = 'Object is locked'
end

--------------------------------------------------------------------------------------------------------------------------------
-- Select the return values
select @error, @error_message
end
