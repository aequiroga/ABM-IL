<!--#include file="..\System\DataFunctions.asp"-->
  <%
  dim Nombre
  dim Apellido
  dim Fecha 
  dim DNI
  dim Cargo
  dim Telefono 
  dim Direccion
  dim Email
  dim DoAction
  dim IDEmp
  IDEmp = Request.QueryString("IDEmp")
  DoAction = request.Form("DoAction")
  if DoAction = 1 then
    Val = EditarEmpleado(IDEmp,Nombre,Apellido,DNI,Fecha,Cargo,Telefono,Email,Direccion)
    if Val=1 then
      response.redirect("Details.asp?Result=2")
      response.end()
    end if
  end if
   if IDEmp < 0 then
    response.redirect("index.asp?Result=-1")
    response.end()
  else
    SET ObjRS = DetallesEmpleado(IDEmp)
    if ObjRS.EOF then
      response.redirect("index.asp?Result=-1")
      response.end()
    end if
  end if
  %>
<!DOCTYPE html>
<html>

  <head>
    <link rel="stylesheet" type="text/css" href="../System/Style.css">
    <meta charset="utf-8">
    <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-3.3.1.min.js"></script>
    <script src="../System/Forms.js"></script>
    <title>Editar información Matias</title>
  </head>


  <script>
  $(document).ready(()=>{
      var PhoneStr = "<% =Telefono %>";

      console.log(PhoneStr)

      var splitt = PhoneStr.split(", ");//Separa string en array
      $("#inputTelefono0").val(splitt[0]);
    for(x in splitt) {
      console.log("aa"+x);
      //Evade el primer numero, que ya fue ingresado
      if (x > 0) addTelefono(splitt[x]);//Ejecuta funcion, mandando un numero segun el indice

      console.log("bb"+splitt[x]+"c"+x);
      }
  })
  </script>

  <body>
    <!-- Header-->
    <div class="header">
      <img src="../System/Images/infiniteloop-logo.svg" class="logo">
      <ul class="navbar">
        <li><a href="index.asp"><img src="../system/Images/Home.png" height="15px" width="15px">Listado de empleados</a></li>
        <li><a href="Details.asp"><img src= "../System/Images/Details.png" height="15px" width="15px">Detalles ultimo empleado</a></li>
        <li><a href="Add.asp"><img src = "../System/Images/AddUser.png" height= "15px" width= "15px">Agregar empleado</a></li>
      </ul>

      <ul class="breadcrumb">
        <li><a href="index.asp">Listado de empleados</a></li>
        <li><a href="Details.asp">Detalles de Matias Sanchez</a></li>
        <li>Editar empleado</li>
      </ul>
    </div>

<!--Input de edicion-->
      <div class="DivBox1">

        <% if Val = -1 then %>
          <div class="MensajeError"><!--Si resultado es -1, crea el div-->
          <p> Se ha producido un error.
          </div>
          <script>
            setTimeout(()=>{$(".MensajeError").slideUp();},3000);
          </script> 
        <%end if%>

        <div class="InputUser">
          <form method="POST" id="FormEdit"action="Edit.asp" autocomplete="on">
        Ingrese el nombre:
        <br>
        <input type="text" class="AlphaForm" id="inputNombre" placeholder="Nombre" name="EditNombre" value="<% = Nombre %>">
        <br>
        Ingrese el apellido:
        <br>
        <input type="text" class="AlphaForm" id="inputApellido" placeholder="Apellido" name="EditApellido" value="<% = Apellido %>">
        <br>
        Ingrese la fecha de nacimiento:
        <br>
        <input type="date" id="inputFecha" name="EditFecha" value="<% = Fecha %>">
        <br>
        Ingrese el DNI:
        <br>
        <input type="text" class="DNINumForm" id="inputDNI" placeholder="**.***.***" name="EditDNI" value="<% = DNI %>">
        <br>
        Seleccione el cargo:
        <br>
        <select id="inputCargo" name="EditCargo">
          <option value="0">Seleccione</option>
          <option value="1">Desarrollador<% if Cargo = 1 then%> (Seleccionado)<%end if%></option>
          <option value="2">Administrador<% if Cargo = 2 then%> (Seleccionado)<%end if%></option>
          <option value="3">Lider<% if Cargo = 3 then%> (Seleccionado)<%end if%></option>
          <option value="4">Diseñador<% if Cargo = 4 then%> (Seleccionado)<%end if%></option>
          <option value="5">Jefe<% if Cargo = 5 then%> (Seleccionado)<%end if%></option>
        </select>
        <br>
        Ingrese el telefono:
        <br>
        <div class="FormInputTel" telID = 0>
        <input type="text" class="NumForm Telefono" id="inputTelefono0" placeholder="Telefono"  name="NewTelefono" telID = 0>
        <img src="../System/Images/Add.png" height="28px" width="28px" class="buttonAddTel" id="buttonAddTel">
        </div>
        <br>
        Ingrese la direccion:
        <br>
        <input type="text" placeholder="Direccion" id="inputDireccion" class="AlphaNumForm" name="EditDirecccion" value="<%response.write(Direccion)%>">
        <br>
        Ingrese el E-Mail:
        <br>
        <input type="email" placeholder="E-Mail" id="inputemail" name="EditEmail" value="<%response.write(Email)%>">
        <br>
        <input type="hidden" name="DoAction" value="1">
          </form>
          <button class="ButtonForm" id="BotonVolverI">Volver</button>
          <input type="submit" class="ButtonForm" id="ValidarForm" value="Añadir cambios al sistema" form="FormEdit">
          <button class="RojoCritico"id="BotonEliminarU" >Eliminar Usuario</button>
        </div>
      </div>

  </body>

</html>
