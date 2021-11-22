<!--#include file="configuration.asp" -->
<!--#include file="functions.asp" -->
<%
' ***************************************************************************************************************************
' ***************************************************************************************************************************
' ***************************************************************************************************************************

Function getADORecordSet(strSQLSelect)
	
    Dim objConn, objRs

	Set objConn = Server.CreateObject("ADODB.Connection")
		
	Set objRs = Server.CreateObject("ADODB.Recordset")
    
        objConn.Open sSCN_CnxString
	    Set objRs.ActiveConnection = objConn
    
	    objRs.CursorLocation = adUseClient
	    objRs.CursorType = adOpenDynamic
        objRs.CacheSize = 50
    
	    objRs.Open strSQLSelect, , adOpenStatic, adLockReadOnly, adCmdText

	    Set getADORecordSet = objRs
	
	    Set objRs.ActiveConnection = Nothing      
       
	    objConn.Close
	    Set objConn = Nothing
 
End Function

' ***************************************************************************************************************************
' ***************************************************************************************************************************
' ***************************************************************************************************************************

Function getSortedRecordSet(PageIndex, PageSize, str_TableName, str_FilterConditions, str_SortField, str_SortDirection, totalPagesQty, totalRecordsQty)

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
		.CommandText = "dbo.SP_Core_Generic_Search"
		.CommandType = 4
		.CommandTimeout = 0
		.Prepared = true

        .Parameters.Append .CreateParameter("RETURN_VALUE", 3, 4)
        .Parameters.Append .CreateParameter("@IL_CustomerID", 3, 1, 1, IL_CustomerID)
		.Parameters.Append .CreateParameter("@PageIndex", 3, 1, 1, PageIndex)
		.Parameters.Append .CreateParameter("@PageSize", 3, 1, 1, PageSize)
		.Parameters.Append .CreateParameter("@str_TableName", 129, 1, 200, str_TableName)
		.Parameters.Append .CreateParameter("@str_FilterConditions", 129, 1, 2000, str_FilterConditions)
		.Parameters.Append .CreateParameter("@str_SortField", 129, 1, 200, str_SortField)
		.Parameters.Append .CreateParameter("@str_SortDirection", 129, 1, 4, str_SortDirection)

	    .Parameters.Append .CreateParameter("@totalPagesQty", 3, 2, 1)
	    .Parameters.Append .CreateParameter("@totalRecordsQty", 3, 2, 1)
		'debugParamAttribs2(objCmd)
    End With
	
	objRs.Open objCmd, ,  3, 1, 4
    objRs.CacheSize = 50
	
	int_SP_RV  = objCmd.Parameters("RETURN_VALUE")
    totalPagesQty   = objCmd.Parameters("@totalPagesQty")
    totalRecordsQty = objCmd.Parameters("@totalRecordsQty")
	
	Set getSortedRecordSet = objRs
	
	Set objCmd = Nothing

End Function


' ***************************************************************************************************************************
' ***************************************************************************************************************************
' ***************************************************************************************************************************


Function DisplayParamAtribs(objCmd)

  Dim blnTR1
  Dim param
    	
  Response.Write "<table>"
  Response.Write "<th>param.Name</th>"
  Response.Write "<th>param.Direction</th>"
  Response.Write "<th>param.Type</th>"
  Response.Write "<th>param.Precision</th>"
  Response.Write "<th>param.Size</th>"
  Response.Write "<th>param.Value</th>"
  
  for each param in objCmd.Parameters
        If blnTR1 Then 
            Response.Write "<TR style=""background-color:silver;"">"
        Else
            Response.Write "<TR>"
        End If
        blnTR1 = NOT blnTR1
		
        Response.Write  "" & _
            "<TD align=""left""> " & param.Name & " </TD>" & _
            "<TD align=""center""> " & param.Direction & " - " & GetParameterDirectionEnum(param.Direction) & "</TD>"  & _
        "<TD align=""center""> " & param.Type & " - " & GetDataTypeEnum(param.Type) & "</TD>" & _
        "<TD align=""center""> " & param.Precision & "</TD>" & _
            "<TD align=""center""> " & param.Size & "</TD>" & _
            "<TD align=""center""> " & param.Value & "</TD>" & _
        "</TR>"
  next
  Response.Write "</table>"
End Function

'Funci�n para debuguear SPs
Function debugParamAttribs(objCmd)
	
	spName = Replace(objCmd.CommandText,"{ ? = call dbo.","")
	spName = Replace(objCmd.CommandText,"{ ? = call","")
	spName = Replace(spName,"?","")
	spName = Replace(spName,", ","")
	spName = Replace(spName,"() }","")

	sParams = "<br /> -- il_helptext '" & spName & "'<br /><br />"
    sParams = sParams & "DECLARE @RC INT = '', @errorDesc NVARCHAR(500) = ''<br /> "
    sParams = sParams & "<br/>EXEC @RC = " & spName & " <br /> "
	iParamsCount = 0
	totalParamsCount = objCmd.Parameters.Count

    printString = ""
	For Each objParameter In objCmd.Parameters
		iParamsCount = iParamsCount + 1
        If objParameter.Name <> "RETURN_VALUE" And objParameter.Name <> "@RETURN_VALUE" Then
            paramValue = objParameter.Value
            Select Case objParameter.Type
                Case 11
                    if CBool(paramValue) Then
                        paramValue = 1
                    Else
                        paramValue = 0
                    End If
                Case 6
                    paramValue = Replace(paramValue,",",".")
                Case 129,202,200,72
                    paramValue = "'" & paramValue & "'"
            End Select
            If objParameter.direction <> 4 And objParameter.direction <> 2 Then
				IF totalParamsCount = iParamsCount Then
					sParams = sParams & paramValue & " -- " & objParameter.Name & " | " & objParameter.Type & " <br />" 
				Else
					sParams = sParams & paramValue & ", -- " & objParameter.Name & " | " & objParameter.Type & " <br />" 
				End If
            Else
			If totalParamsCount = iParamsCount  Then
				sParams = sParams & objParameter.Name & " OUT -- " & objParameter.Type & " <br />" 
			Else
				sParams = sParams & objParameter.Name & " OUT , -- " & objParameter.Type & " <br />" 
			End If
            printString = printString & "<br/>PRINT " & objParameter.Name
            End If
        End If
    Next
	sParams = sParams & "<br /> PRINT @RC " & printString
	Response.Write(sParams)
	Response.End()

End Function

Function debugParamAttribs2(objCmd)
	
	spName = Replace(objCmd.CommandText,"{ ? = call dbo.","")
	spName = Replace(objCmd.CommandText,"{ ? = call","")
	spName = Replace(spName,"?","")
	spName = Replace(spName,", ","")
	spName = Replace(spName,"() }","")

	headerDeclare = "<br /> -- il_helptext '" & spName & "'<br /><br /> "
    variableDeclare = " DECLARE @RC INT = '' "
    sParams = "<br/>EXEC @RC = " & spName & " <br /> "
	iParamsCount = 0
	totalParamsCount = objCmd.Parameters.Count

    printString = ""
	For Each objParameter In objCmd.Parameters
		iParamsCount = iParamsCount + 1
        If objParameter.Name <> "RETURN_VALUE" And objParameter.Name <> "@RETURN_VALUE" Then
            paramValue = objParameter.Value
            Select Case objParameter.Type
                Case 11
                    if CBool(paramValue) Then
                        paramValue = 1
                    Else
                        paramValue = 0
                    End If
                    paramDataType = "BIT = 0"
                Case 6
                    paramValue = Replace(paramValue,",",".")
                    paramDataType = "MONEY = 0"
                Case 129,202,200,72
                    paramValue = "'" & paramValue & "'"
                    paramDataType = "VARCHAR(MAX) = ''"
                Case 3
                    paramDataType = "INT = 0"
                    If paramValue = "" Then
                        paramValue = 0
                    End If 
            End Select
            If objParameter.direction <> 4 And objParameter.direction <> 2 Then
				IF totalParamsCount = iParamsCount Then
					sParams = sParams & paramValue & " -- " & objParameter.Name & " | " & objParameter.Type & " <br />" 
				Else
					sParams = sParams & paramValue & ", -- " & objParameter.Name & " | " & objParameter.Type & " <br />" 
				End If
            Else
			    If totalParamsCount = iParamsCount  Then
				    sParams = sParams & objParameter.Name & " OUT -- " & objParameter.Type & " <br />" 
    
			    Else
				    sParams = sParams & objParameter.Name & " OUT , -- " & objParameter.Type & " <br />" 
			    End If

                variableDeclare = variableDeclare & ", " & objParameter.Name &  " " & paramDataType & " "
                printString = printString & "<br/>PRINT " & objParameter.Name
            End If
        End If
    Next
	sParams = headerDeclare & variableDeclare & " <br /> " & sParams & "<br /> PRINT @RC " & printString
	Response.Write(sParams)
	Response.End()

End Function

Function iterateParamAttribs(int_SP_RV,objCmd)

    int_SP_RV = CLng(ReplaceNullValues(int_SP_RV))
    If int_SP_RV <> 1 Then
          spName = Replace(objCmd.CommandText,"{ ? = call dbo.","")
          spName = Replace(objCmd.CommandText,"{ ? = call","")
          spName = Replace(spName,"?","")
          spName = Replace(spName,", ","")
          spName = Replace(spName,"() }","")
        sParams = "<br /> -- il_helptext '" & spName & "'<br />"
        sParams = sParams & "-- RC: " & int_SP_RV & "<br /> DECLARE @RC INT = 0, @errorDesc NVARCHAR(500) = ''<br /> "
        sParams = sParams & "EXEC @RC = " & spName & " <br /> "
        
		iParamsCount = 0
		totalParamsCount = objCmd.Parameters.Count
		For Each objParameter In objCmd.Parameters
		  iParamsCount = iParamsCount + 1
          If objParameter.Name <> "RETURN_VALUE" Then
            If objParameter.Type = 129 Then ' 129: VARCHAR
              If objParameter.direction <> 4 And objParameter.direction <> 2 Then ' 4: OUT
                IF totalParamsCount = iParamsCount Then
					sParams = sParams & "'" & ReplaceNullValues(objParameter.Value) & "' -- " & objParameter.Name & " | " & objParameter.Type  & " | " & objParameter.Size & " | " & Len(ReplaceNullValues(objParameter.Value)) & " <br />" 
				Else
					sParams = sParams & "'" & ReplaceNullValues(objParameter.Value) & "', -- " & objParameter.Name & " | " & objParameter.Type  & " | " & objParameter.Size & " | " & Len(ReplaceNullValues(objParameter.Value)) & " <br />" 
				End If
              Else
				IF totalParamsCount = iParamsCount Then
					sParams = sParams & objParameter.Name & " OUT -- " & objParameter.Type  & " | " & objParameter.Size & " | " & Len(ReplaceNullValues(objParameter.Value)) & " Devolvi&oacute;: " & objParameter.Value & " <br />" 
				Else
					sParams = sParams & objParameter.Name & " OUT, -- " & objParameter.Type  & " | " & objParameter.Size & " | " & Len(ReplaceNullValues(objParameter.Value)) & " Devolvi&oacute;: " & objParameter.Value & " <br />" 
				End If
              End If
            Else 
              If objParameter.direction <> 4 And objParameter.direction <> 2 Then ' 4: OUT
                If objParameter.Type = 6 Then
					IF totalParamsCount = iParamsCount Then
						sParams = sParams & ReplaceNullValues(objParameter.Value) & " -- " & objParameter.Name & " | " & objParameter.Type & " <br />" 
					Else
						sParams = sParams & ReplaceNullValues(objParameter.Value) & ", -- " & objParameter.Name & " | " & objParameter.Type & " <br />" 
					End If
                  
                Else
					IF totalParamsCount = iParamsCount Then
						sParams = sParams & Replace(ReplaceNullValues(objParameter.Value),",",".") & " -- " & objParameter.Name & " | " & objParameter.Type & " <br />" 
					Else
						sParams = sParams & Replace(ReplaceNullValues(objParameter.Value),",",".") & ", -- " & objParameter.Name & " | " & objParameter.Type & " <br />" 
					End If
                End If
              Else
				IF totalParamsCount = iParamsCount  Then
					sParams = sParams & objParameter.Name & " OUT --  " & objParameter.Type & " Devolvi&oacute;: " & objParameter.Value & " <br />" 
				Else
					sParams = sParams & objParameter.Name & " OUT, -- " & objParameter.Type & " Devolvi&oacute;: " & objParameter.Value & " <br />" 
				End If
              End If
            End If
          End If
        Next
		sParams = sParams & "<br /> PRINT @RC "
        sParams = sParams & "<br /> PRINT @errorDesc "
        sSubject = "SP:" & spName
        EmailError 3, sParams
    End If

End Function


Function ExecProc(m_strProcName, strInputValues, strOutputValues, int_SP_RV, errDesc)

    Dim objConn 'As ADODB.Connection
	Dim objCmd 'As ADODB.Command

	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objCmd = Server.CreateObject("ADODB.Command")
    if m_strProcName = "dbo.SP_StatLog_Add" Then
        objConn.Open sStats_CnxString
    Else
	    objConn.Open sSCN_CnxString
    End If

	Set objCmd.ActiveConnection = objConn
	objCmd.CommandText = m_strProcName
	objCmd.CommandType = 4
	objCmd.CommandTimeout = 0
	objCmd.Prepared = true
		
    'Call refresh to retrieve the values	
    objCmd.Parameters.Refresh
    
    '************************************************************************************
    
    Dim blnFirstParameter	'Is this is the first parameter
	Dim strCommandParameters	'Parameters for the command
	Dim strOutputParameters	'Retrieving of output parameters
    Dim intParamIndex 'Indice donde estoy en el Command
    Dim arrInputParams
        
    intInputParamIndex = 0
    intOutputParamIndex = 0
    intParamIndex = 0
	blnFirstParameter = True
	arrInputParams = Split(strInputValues, "|")
	
	intParamCount = objCmd.Parameters.Count 'Variable que contiene la cantidad de param del Command
	
	If intParamCount-1 = UBound(arrInputParams)+1 Then 'El primer valor del command es el return value y el Ubound devuelve un elemento menos.
	
	    For Each param In objCmd.Parameters		
		    If NOT param.Direction = 4 Then		'adParamReturnValue	
			    'Tengo que setear los valores del SP, 
                
                
			    param.Value = arrInputParams(intInputParamIndex)    			
		        intInputParamIndex = intInputParamIndex + 1

                intParamIndex = intParamIndex + 1
		    Else
		        intOutputParamCount = intOutputParamIndex + 1
		    End If 	
            
	    Next

	    
	    objCmd.Execute

		int_SP_RV = CLng(ReplaceNullValues(objCmd.Parameters(0).Value))
		
		'If intOutputParamCount > 0 Then
		    Execute("Dim arrOutputParams(" & objCmd.Parameters.Count & ")")		
		    For Each param in objCmd.Parameters		
              
		        'If param.Direction = 4 Then		'adParamReturnValue
			        arrOutputParams(intOutputParamIndex) = param.Value
		            intOutputParamIndex = intOutputParamIndex + 1
		        'End If	

	        Next

            strOutputValues = join(arrOutputParams, ",")
	    'End If
    Else
        errDesc = "Parameters definition error. SP " & m_strProcName & " is expecting " & intParamCount-1 & " params. Array qty: " & (UBound(arrInputParams)+1)
    End If
    
    Set objCmd = Nothing
	objConn.Close
	Set objConn = Nothing
	    
End Function

' RME: 2015.05.22: Agrego funciones para nuevo framework de programaci�n. [iPN Framework]. Documentaci�n disponible en https://docs.google.com/a/infiniteloop.com.ar/document/d/16RLgqNck6nUNgQH3KtrQeeSM9UclhLrPpZ0T-Ee0rlo/edit?usp=sharing
Function ExecProcWithArray(m_strProcName, strInputValues, strOutputValues, int_SP_RV, errDesc)

    Dim objConn 'As ADODB.Connection
    Dim objCmd 'As ADODB.Command

    Set objConn = Server.CreateObject("ADODB.Connection")
    Set objCmd = Server.CreateObject("ADODB.Command")
    if m_strProcName = "dbo.SP_StatLog_Add" Then
        objConn.Open sStats_CnxString
    Else
        objConn.Open sSCN_CnxString
    End If

    Set objCmd.ActiveConnection = objConn
    objCmd.CommandText = m_strProcName
    objCmd.CommandType = 4
    objCmd.CommandTimeout = 0
    objCmd.Prepared = true
        
    'Call refresh to retrieve the values    
    objCmd.Parameters.Refresh
    
    '************************************************************************************
    
    Dim blnFirstParameter   'Is this is the first parameter
    Dim strCommandParameters    'Parameters for the command
    Dim strOutputParameters 'Retrieving of output parameters
    Dim intParamIndex 'Indice donde estoy en el Command
    Dim arrInputParams
        
    intInputParamIndex = 0
    intOutputParamIndex = 0
    intParamIndex = 0
    blnFirstParameter = True
    'RME: Cambio para que use un array en lugar de pasar un string
    arrInputParams = strInputValues 'Split(strInputValues, "|")
    
    intParamCount = objCmd.Parameters.Count 'Variable que contiene la cantidad de param del Command
    
    If intParamCount-1 = UBound(arrInputParams)+1 Then 'El primer valor del command es el return value y el Ubound devuelve un elemento menos.
        if bDebug Then
            Response.Write("===============================================================================================")
            Response.Write("<br/>")
            Response.Write("Values de los params de entrada")
            Response.Write("<br/>")
            Response.Write("===============================================================================================")
            Response.Write("<br/>")
        End If

        if bDebug Then
            response.write("<table style='border:solid 1px #ccc' cellpadding='0' cellspacing='0'>")
            response.write("<tr><th  style='border:solid 1px #ccc'>Nombre</th><th  style='border:solid 1px #ccc'>Tipo</th><th style='border:solid 1px #ccc'>Direcci&oacute;n</th><th style='border:solid 1px #ccc'>Valor</th></tr>")
        end if

        For Each param In objCmd.Parameters     
            if bDebug Then
                Response.Write("<tr>")
                Response.Write("<td style='border:solid 1px #ccc'>" & param.Name & "</td>")
                Response.Write("<td style='border:solid 1px #ccc'>" & param.Type & "</td>")
                Response.Write("<td style='border:solid 1px #ccc'>" & param.Direction & "</td>")
                Response.Write("<td style='border:solid 1px #ccc'>" & arrInputParams(intInputParamIndex) & "</td>")
                Response.Write("</tr>")
            End If
            If NOT param.Direction = 4 And NOT param.Direction = 3 Then     'adParamReturnValue 
                'Tengo que setear los valores del SP, 
                paramValue = arrInputParams(intInputParamIndex)
               
                param.Value = paramValue
                
                intInputParamIndex = intInputParamIndex + 1

                intParamIndex = intParamIndex + 1
            Else
                intOutputParamCount = intOutputParamIndex + 1
            End If  
        Next
         if bDebug Then
            response.write("</table>")
        end if

        'debugParamAttribs(objCmd)
        
        objCmd.Execute

        int_SP_RV = CLng(ReplaceNullValues(objCmd.Parameters(0).Value))
        'response.Write(int_SP_RV)
        'response.end()
         If int_SP_RV = -2 Or int_SP_RV = -3 Then
          iterateParamAttribs int_SP_RV,objCmd
        End If

        'If intOutputParamCount > 0 Then
            'Execute("Dim arrOutputParams(" & objCmd.Parameters.Count & ")")     
            Set arrOutputParams = CreateObject("Scripting.Dictionary")

            For Each param in objCmd.Parameters     
                'If param.Direction = 4 Then        'adParamReturnValue

                    'arrOutputParams(intOutputParamIndex) = param.Value
                    'intOutputParamIndex = intOutputParamIndex + 1
                    arrOutputParams.Add param.Name, param.Value
                'End If 
            Next

            If bDebug Then
                Response.Write("===============================================================================================")
                f_errorCode = arrOutputParams("@errorCode")
                f_errorDesc = arrOutputParams("@errorDesc")
                Response.Write("<br/>")
                Response.Write("Resultado de la ejecuci&oacute;n")
                Response.Write("<br/>")
                Response.Write("===============================================================================================")
                Response.Write("<br/>errorCode => " & f_errorCode)
                Response.Write("<br/>f_errorDesc => " & f_errorDesc)
                Response.Write("<br/>int_SP_RV => " & int_SP_RV)
                Response.Write("<br/>")
                Response.Write("===============================================================================================")
                Response.Write("<br/>")
                Response.Write("Values de los params de salida")
                Response.Write("<br/>")
                Response.Write("===============================================================================================")
                Response.Write("<br/>")
                For Each x In arrOutputParams
                    response.write(x & " => " & arrOutputParams(x) & "<br />")
                Next
                'Response.End()
            End If
            'strOutputValues = join(arrOutputParams, ",")
            Set strOutputValues = arrOutputParams
        'End If
    Else
        'TODO: Enviar esto por email. Junto con los params del SP
        errDesc = "Parameters definition error. SP " & m_strProcName & " is expecting " & intParamCount-1 & " params. Array qty: " & (UBound(arrInputParams)+1)
        if bDebug Then
            response.Write(errDesc)
            Response.End()
        End If
    End If
    
    Set objCmd = Nothing
    objConn.Close
    Set objConn = Nothing
        
End Function

' RME: 2015.05.22: Agrego funciones para nuevo framework de programaci�n. [iPN Framework]. Documentaci�n disponible en https://docs.google.com/a/infiniteloop.com.ar/document/d/16RLgqNck6nUNgQH3KtrQeeSM9UclhLrPpZ0T-Ee0rlo/edit?usp=sharing
Function GetRecordset(m_strProcName, strInputValues, int_SP_RV)
'Funci�n para poder traer un RS s�lo pas�ndole el nombre del SP a ejecutar y un array con los params
    Dim objConn 'As ADODB.Connection
    Dim objCmd 'As ADODB.Command
    DIm objRs

    Set objConn = Server.CreateObject("ADODB.Connection")
    Set objCmd = Server.CreateObject("ADODB.Command")
    Set objRs = Server.CreateObject("ADODB.Recordset")

    if m_strProcName = "dbo.SP_StatLog_Add" Then
        objConn.Open sStats_CnxString
    Else
        objConn.Open sSCN_CnxString
    End If

    Set objCmd.ActiveConnection = objConn
    objCmd.CommandText = m_strProcName
    objCmd.CommandType = 4
    objCmd.CommandTimeout = 0
    objCmd.Prepared = true


    objRs.CursorLocation = 3


    'Call refresh to retrieve the values    
    objCmd.Parameters.Refresh
    
    '************************************************************************************
    
    Dim blnFirstParameter   'Is this is the first parameter
    Dim strCommandParameters    'Parameters for the command
    Dim strOutputParameters 'Retrieving of output parameters
    Dim intParamIndex 'Indice donde estoy en el Command
    Dim arrInputParams
        
    intInputParamIndex = 0
    intOutputParamIndex = 0
    intParamIndex = 0
    blnFirstParameter = True
    'RME: Cambio para que use un array en lugar de pasar un string
    arrInputParams = strInputValues 'Split(strInputValues, "|")
    
    intParamCount = objCmd.Parameters.Count 'Variable que contiene la cantidad de param del Command
    
    If intParamCount-1 = UBound(arrInputParams)+1 Then 'El primer valor del command es el return value y el Ubound devuelve un elemento menos.
        if bDebug Then
            Response.Write("===============================================================================================")
            Response.Write("<br/>")
            Response.Write("Values de los params de entrada")
            Response.Write("<br/>")
            Response.Write("===============================================================================================")
            Response.Write("<br/>")
        End If

        if bDebug Then
            response.write("<table style='border:solid 1px #ccc' cellpadding='0' cellspacing='0'>")
            response.write("<tr><th  style='border:solid 1px #ccc'>Nombre</th><th  style='border:solid 1px #ccc'>Tipo</th><th style='border:solid 1px #ccc'>Direcci&oacute;n</th><th style='border:solid 1px #ccc'>Valor</th></tr>")
        end if

        For Each param In objCmd.Parameters     
            if bDebug Then
                Response.Write("<tr>")
                Response.Write("<td style='border:solid 1px #ccc'>" & param.Name & "</td>")
                Response.Write("<td style='border:solid 1px #ccc'>" & param.Type & "</td>")
                Response.Write("<td style='border:solid 1px #ccc'>" & param.Direction & "</td>")
                Response.Write("<td style='border:solid 1px #ccc'>" & arrInputParams(intInputParamIndex) & "</td>")
                Response.Write("</tr>")
            End If
            If NOT param.Direction = 4 And NOT param.Direction = 3 Then     'adParamReturnValue 
                'Tengo que setear los valores del SP, 
                paramValue = arrInputParams(intInputParamIndex)
               
                param.Value = paramValue
                
                intInputParamIndex = intInputParamIndex + 1

                intParamIndex = intParamIndex + 1
            Else
                intOutputParamCount = intOutputParamIndex + 1
            End If  
        Next
        
        if bDebug Then
            response.write("</table>")
        end if

       ' if bDebugParamAttribs Then
        '    bDebugParamAttribs = false
         ''   debugParamAttribs(objCmd)
        'end if
        
        objRs.Open objCmd, ,  3, 1, 4    
        Set GetRecordset = objRs
       
        int_SP_RV = CLng(ReplaceNullValues(objCmd.Parameters(0).Value))
        
         If int_SP_RV = -2 Or int_SP_RV = -3 Then
          iterateParamAttribs int_SP_RV,objCmd
        End If

        Set objCmd = Nothing
        
    Else
        'TODO: Enviar esto por email. Junto con los params del SP
        errDesc = "Parameters definition error. SP " & m_strProcName & " is expecting " & intParamCount-1 & " params. Array qty: " & (UBound(arrInputParams)+1)
        if bDebug Then
            response.Write(errDesc)
            response.Write("<br/>")
            response.Write("<br/>")
            response.write("<table style='border:solid 1px #ccc' cellpadding='0' cellspacing='0'>")
            response.write("<tr><th  style='border:solid 1px #ccc'>Nombre</th><th  style='border:solid 1px #ccc'>Tipo</th><th style='border:solid 1px #ccc'>Direcci&oacute;n</th><th style='border:solid 1px #ccc'>Valor</th></tr>")
            For Each param In objCmd.Parameters     
                    Response.Write("<tr>")
                    Response.Write("<td style='border:solid 1px #ccc'>" & param.Name & "</td>")
                    Response.Write("<td style='border:solid 1px #ccc'>" & param.Type & "</td>")
                    Response.Write("<td style='border:solid 1px #ccc'>" & param.Direction & "</td>")
                    Response.Write("</tr>")
            Next
            Response.End()
        End If
    End If

End Function


Function GetParameterDirectionEnum(lngDirection)
    Select Case lngDirection
        Case 0  'adParamUnknown
            GetParameterDirectionEnum = "adParamUnknown"
        Case 1  'adParamInput
            GetParameterDirectionEnum = "adParamInput"
        Case 2  'adParamOutput
            GetParameterDirectionEnum = "adParamOutput"
        Case 3  'adParamInputOutput
            GetParameterDirectionEnum = "adParamInputOutput"
        Case 4  'adParamReturnValue
            GetParameterDirectionEnum = "adParamReturnValue"
        Case Else
			GetParameterDirectionEnum = "<B>Direction Not Found</B>"
    End Select	
End Function


Function GetDataTypeEnum(lngType)    
    Select Case lngType					
        Case 0
            GetDataTypeEnum = "adEmpty"
        Case 2
            GetDataTypeEnum = "adSmallInt"
        Case 3
            GetDataTypeEnum = "adInteger"
        Case 4
            GetDataTypeEnum = "adSingle"
        Case 5
            GetDataTypeEnum = "adDouble"
        Case 6
            GetDataTypeEnum = "adCurrency"
        Case 7
            GetDataTypeEnum = "adDate"
        Case 8
            GetDataTypeEnum = "adBSTR"
        Case 9
            GetDataTypeEnum = "adIDispatch"
        Case 10
            GetDataTypeEnum = "adError"
        Case 11
            GetDataTypeEnum = "adBoolean"
        Case 12
            GetDataTypeEnum = "adVariant"
        Case 13
            GetDataTypeEnum = "adIUnknown"
        Case 14
            GetDataTypeEnum = "adDecimal"
        Case 16
            GetDataTypeEnum = "adTinyInt"
        Case 17
            GetDataTypeEnum = "adUnsignedTinyInt"
        Case 18
            GetDataTypeEnum = "adUnsignedSmallInt"
        Case 19
            GetDataTypeEnum = "adUnsignedInt"
        Case 20
            GetDataTypeEnum = "adBigInt"
        Case 21
            GetDataTypeEnum = "adUnsignedBigInt"
        Case 64
            GetDataTypeEnum = "adFileTime"
        Case 72
            GetDataTypeEnum = "adGUID"
        Case 128
            GetDataTypeEnum = "adBinary"
        Case 129
            GetDataTypeEnum = "adChar"
        Case 130
            GetDataTypeEnum = "adWChar"
        Case 131
            GetDataTypeEnum = "adNumeric"
        Case 132
            GetDataTypeEnum = "adUserDefined"
        Case 133
            GetDataTypeEnum = "adDBDate"
        Case 134
            GetDataTypeEnum = "adDBTime"
        Case 135
            GetDataTypeEnum = "adDBTimeStamp"
        Case 136
            GetDataTypeEnum = "adChapter"
        Case 138
            GetDataTypeEnum = "adPropVariant"
        Case 139
            GetDataTypeEnum = "adVarNumeric"
        Case 200
            GetDataTypeEnum = "adVarChar"
        Case 201
            GetDataTypeEnum = "adLongVarChar"
        Case 202
            GetDataTypeEnum = "adVarWChar"
        Case 203
            GetDataTypeEnum = "adLongVarWChar"
        Case 204
            GetDataTypeEnum = "adVarBinary"
        Case 205
            GetDataTypeEnum = "adLongVarBinary"
        Case 8192
            GetDataTypeEnum = "adArray"
        Case Else
            GetDataTypeEnum = "<B>Type Constant Not Found</B>"
    End Select
End Function


Function FieldExistsInRs(rs, sFieldName)
    
    If IsObject(rs) Then
        If Not rs.EOF Then
            Dim fld
            For Each fld In rs.Fields
                If UCase(fld.Name) = UCase(sFieldName) Then
                    FieldExistsInRs = True
                    Exit Function
                End If
            Next
        End If
    End If
End Function

%>
