<!--#include file="./data/configuration.asp" -->
<!--#include file="./data/adovbs.inc" -->
<%
Function AgregarEmpleado(Nombre,Apellido,DNI,Fecha,Cargo,Telefono,Email,Direccion)
        
        'AgregarEmpleado=1
        AgregarEmpleado=1
End Function

Function EditarEmpleado(IDEmp,Nombre,Apellido,DNI,Fecha,Cargo,Telefono,Email,Direccion)
End Function

Function BorrarEmpleado(IDEmp)
        
        'BorrarEmpleado=1
        BorrarEmpleado=-1
End Function

Function BusquedaEmpleado(Filter)
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
			.CommandText = "dbo.SP_AQ_Employees_Search"
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true

                        .Parameters.Append .CreateParameter("RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@Filter", 129, 1, Len(Trim(filter)) +1, Trim(filter))

		End With
		
		
		objRs.Open objCmd, ,  3, 1, 4
		Set BusquedaEmpleado = objRs
		
		Set objCmd = Nothing

End Function

'GET
Function DetallesEmpleado(IDEmp)
                if IDEmp  then
                StrSQL = "SELECT * FROM VW_AQ_Employees  WHERE EmployeeID ="
                StrSQL= StrSQL & IDEmp
                Set DetallesEmpleado = getADORecordSet(StrSQL)
                end if

End Function


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
%>