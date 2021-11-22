<%

'JM: 2018-10-24: Creación del archivo.

Function getContacts()

	Dim strSQLSelect

	strSQLSelect = "SELECT * FROM TB_JM_Employees"
	
	'strSQLSelect = strSQLSelect & " ORDER BY contactName DESC "

	Set getContacts = getADORecordSet(strSQLSelect)

End Function


'Nombre: replaceNullValues
'Descripcion: Funcion utilizada para cambiar los campos vacios por 0.
'Usada: En todo el sistema.
Function replaceNullValues(strValue)

	If  Isnull(strValue) Or IsEmpty(strValue) Or strValue = "" Or strValue = " " Or Len(Trim(strValue)) = 0 Or LCase(Trim(strValue)) = "undefined" Or LCase(Trim(strValue)) = "nan" Then 
		replaceNullValues = 0
	Else
		replaceNullValues = strValue
	End If

End Function


Function updateSaleStatus(doAction, b2cSaleID, saleStatusID, voucherDate, voucherChr, notes, invoiceTypeID, invoiceNumberChr, invoiceBillDate, transportID, shippingGuideNumber, shippingDate)

		Dim objConn 'As ADODB.Connection
		Dim objCmd 'As ADODB.Command
		Dim objRs 'As ADODB.Recordset

		Set objConn = Server.CreateObject("ADODB.Connection")
		Set objCmd = Server.CreateObject("ADODB.Command")
		Set objRs = Server.CreateObject("ADODB.Recordset")
		
		objConn.Open sSCN_CnxString
		objRs.CursorLocation = 3

		With objCmd
		Set .ActiveConnection = objConn
			.CommandText = "dbo.SP_B2C_Sale_Status_Edit"
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			
			'input parameters
			.Parameters.Append .CreateParameter("RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@doAction", 129, 1, 1, doAction)
			.Parameters.Append .CreateParameter("@b2cSaleID", 3, 1, 1, b2cSaleID)
			.Parameters.Append .CreateParameter("@IL_CustomerID", 3, 1, 1, IL_CustomerID)  
			.Parameters.Append .CreateParameter("@userID", 3, 1, 1, userID)          
			.Parameters.Append .CreateParameter("@saleStatusID", 3, 1, 1, saleStatusID)
			.Parameters.Append .CreateParameter("@voucherDate", 129, 1, 10, voucherDate)
			.Parameters.Append .CreateParameter("@voucherChr", 129, 1, 50, voucherChr)
			.Parameters.Append .CreateParameter("@notes", 129, 1, 2000, notes)
			.Parameters.Append .CreateParameter("@invoiceTypeID", 3, 1, 1, invoiceTypeID)
			.Parameters.Append .CreateParameter("@invoiceNumberChr", 129, 1, 50, invoiceNumberChr)  
			.Parameters.Append .CreateParameter("@invoiceBillDate", 129, 1, 10, invoiceBillDate)  	
			.Parameters.Append .CreateParameter("@transportID", 3, 1, 1, transportID)
			.Parameters.Append .CreateParameter("@shippingGuideNumber", 129, 1, 50, shippingGuideNumber)  
			.Parameters.Append .CreateParameter("@shippingDate", 129, 1, 10, shippingDate) 	                        	
					
		End With

		debugParamAttribs2(objCmd)
		
		objCmd.Execute
		
		int_SP_RV = objCmd.Parameters("RETURN_VALUE")
		
		updateSaleStatus = int_SP_RV

	    If int_SP_RV < 1 And Not (int_SP_RV < -100 And int_SP_RV > -200) Then
	    	iterateParamAttribs int_SP_RV,objCmd
	    End If		
		
		Set objCmd = Nothing
		
		objConn.Close
		Set objConn = Nothing

End Function


Function getProvidersSearchByPage(PageIndex, PageSize, providerID, providerName, companyName, providerCode, providerTypeID, subProviderTypeID, totalPagesQty, totalRecordsQty)

		Dim objConn 'As ADODB.Connection
		Dim objCmd 'As ADODB.Command
		Dim objRs 'As ADODB.Recordset

		Set objConn = Server.CreateObject("ADODB.Connection")
		Set objCmd = Server.CreateObject("ADODB.Command")
		Set objRs = Server.CreateObject("ADODB.Recordset")
		
		objConn.Open sSCN_CnxString
		objRs.CursorLocation = 3

		With objCmd
		Set .ActiveConnection = objConn
			.CommandText = "dbo.SP_Admin_Provider_Search"
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true

			.Parameters.Append .CreateParameter("RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@il_CustomerID", 3, 1, 1, IL_CustomerID)
			.Parameters.Append .CreateParameter("@companyID", 3, 1, 1, IL_companyID)
			.Parameters.Append .CreateParameter("@accountID", 3, 1, 1, accountID)
			.Parameters.Append .CreateParameter("@PageIndex", 3, 1, 1, PageIndex)
			.Parameters.Append .CreateParameter("@PageSize", 3, 1, 1, PageSize)
			.Parameters.Append .CreateParameter("@providerID", 3, 1, 1, providerID)
			.Parameters.Append .CreateParameter("@providerName", 129, 1, 100, providerName)
			.Parameters.Append .CreateParameter("@companyName", 129, 1, 100, companyName)
			.Parameters.Append .CreateParameter("@providerCode", 129, 1, 10, Left(Trim(providerCode),10)) 'MP + SC:: 13/05/2015::HAcemos un Left para limitar el tamaÃ±o de codigo de provedor, ya que se podia ingresar mas de 10
			.Parameters.Append .CreateParameter("@providerTypeID", 3, 1, 1, providerTypeID)
            'SDC::05/10/2015:: Agrego fltro por sub-tipo de proveedor.
			.Parameters.Append .CreateParameter("@subProviderTypeID", 3, 1, 1, SanitizeIntForSP(subProviderTypeID))
			'CA::2017-07-24:: Agrego userID.
        	.Parameters.Append .CreateParameter("@userID", 3, 1, 1, userID)

			.Parameters.Append .CreateParameter("@totalPagesQty", 3, 2, 1)
			.Parameters.Append .CreateParameter("@totalRecordsQty", 3, 2, 1)

		End With
		
		
		objRs.Open objCmd, ,  3, 1, 4
		
		int_SP_RV       = objCmd.Parameters("RETURN_VALUE")
		totalPagesQty   = objCmd.Parameters("@totalPagesQty")
		totalRecordsQty = objCmd.Parameters("@totalRecordsQty")
		
		Set getProvidersSearchByPage = objRs
		
		Set objCmd = Nothing

End Function


%>