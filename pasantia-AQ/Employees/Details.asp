<!--#include file="../System/DataFunctions.asp"-->
<%
  dim Nombre
  dim Apellido
  dim Fecha 
  dim DNI
  dim Cargo
  dim Telefono 
  dim Direccion
  dim Email
  dim IDEmp
  dim DoAction
  dim Resultado
  dim ObjRs

  IDEmp = CInt(Request.QueryString("IDEmp"))
  Resultado = Request.QueryString("Result")
  DoAction = Request.Form("DoAction")

  if DoAction = 1 Then
    response.redirect("Edit.asp?IDEmp="+IDEmp)
    response.end()
  end if

  if IDEmp < 0 then
    response.redirect("index.asp?Result=-1")
    response.end()
  else
    set ObjRs = DetallesEmpleado(IDEmp)
    if ObjRs.EOF then
      response.redirect("index.asp?Result=-1")
      response.end()
    end if
  end if

%>
<!DOCTYPE html>
<html>

  <head>
    <link rel="stylesheet" type="text/css" href="../System/Style.css">
    <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-3.3.1.min.js"></script>
    <script src="../System/jQuery.js"></script>
    <meta charset="utf-8">
    <title>Detalles de Matias</title>
  </head>

  <body>
<!-- Header-->
    <div class="header">
      <img src="../System/Images/infiniteloop-logo.svg" class="logo">
      <ul class="navbar">
        <li><a href="index.asp"><img src="../system/Images/Home.png" height="15px" width="15px">Listado de empleados</a></li>
        <li><img src= "../System/Images/Details.png" height="15px" width="15px">Detalles ultimo empleado</li>
        <li><a href="Add.asp"><a href="Add.asp"><img src = "../System/Images/AddUser.png" height= "15px" width= "15px">Agregar empleado</a></li>
      </ul>

      <div class="breadcrumb">
        <ul>
          <li><a href="index.asp">Listado de empleados</a></li>
          <li>Detalles de Matias Sanchez</li>
        </ul>
      </div>
    </div>

    <!--Detalles empleado-->
    <div class="DivBox1">

      <% if Resultado > 0 then %>
        <div class="MensajeResult"><!--Si resultado es > a 0, crea el div-->
        <p>
          <%select case Resultado 'Genera un mensaje personalizado segun la accion ejecutada
          case 1
            response.write("Se añadio el usuario.")
          case 2
            response.write("Se edito al usuario.")
          end select%>
        </div>
        <script>
          setTimeout(()=>{$(".MensajeResult").slideUp();},3000);
        </script>
        <% end if %>

    <div class="Detalles">
      <img src="../System/Images/Edit.png" height="50" width="50" id="DetEdit">
      <br>

        <h2><span class="Underline">Datos personales</span></h2>
        <ul class="ListStyle">
          <li><span class="Underline">Nombre:</span> <%=ObjRs("EmployeeName")%></li>
          <li><span class="Underline">Apellido:</span> <%=ObjRs("EmployeeSurname")%></li>
          <li><span class="Underline">DNI:</span> <%=ObjRs("EmployeeDNI")%></li>
          <li><span class="Underline">Cargo:</span> <%=ObjRs("RoleName")%></li>
          <li><span class="Underline">Direccion:</span> <%=ObjRs("EmployeeAddress")%></li>
        </ul>
        <br>
        <h2><span class="Underline">Contacto</span></h2>
        <ul class="ListStyle">
          <li><span class="Underline">Email:</span> <%=ObjRs("EmployeeMail")%></li>
          <%Do While NOT ObjRs.EOF%>
          <li><span class="Underline">Telefono:</span> <%=ObjRs("EmployeePhone")%></li>
          <%Loop%>
        </ul>

        <form method="GET" action="Details.asp" id="FormDetails">
        <input type="hidden" value="1" name="DoAction">
        <input type="hidden" value="<%=IDEmp%>" name="IDEmp">
        </form>
    </div>
  </div>

  </body>

</html>
