<!--#include file="../System/DataFunctions.asp"-->
<%
  dim Result
  DoAction = Request.Form("DoAct")
  if DoAction = 1 then
    Val = BorrarEmpleado(IDEmp)
  if Val = 1 then
    response.redirect("index.asp?Result=1")
    response.end()
    end if
  end if
%>
<html>

  <head>
    <link rel="stylesheet" type="text/css" href="../System/Style.css">
    <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-3.3.1.min.js"></script>
    <script src="../System/Delete.js"></script>
    <meta charset="utf-8">
    <title>Eliminar empleado</title>
  </head>


  <body>
<!--Input de header-->
    <div class="header">
      <img src="../System/Images/infiniteloop-logo.svg" class="logo">
      <ul class="navbar">
        <li><a href="index.asp"><img src="../system/Images/Home.png" height="15px" width="15px">Listado de empleados</a></li>
        <li><a href="Details.asp"><img src= "../System/Images/Details.png" height="15px" width="15px">Detalles ultimo empleado</a></li>
        <li><a href="Add.asp"><img src = "../System/Images/AddUser.png" height= "15px" width= "15px">Agregar empleado</a></li>
      </ul>

      <div class="breadcrumb">
        <ul>
          <li><a href="index.asp">Listado de empleados</a></li>
          <li><a href="Details.asp">Detalles de Matias Sanchez</a></li>
          <li><a href="Edit.asp">Editar detalles</a></li>
          <li>Eliminar empleado</li>
        </ul>
      </div>
    </div>

<!-- Confirmacion de delete-->
    <div class="DivBox2">

      <% if Val = -1 then %>
        <div class="MensajeError2"><!--Si resultado es -1, crea el div-->
          <p> Se ha producido un error.
        </div>
        <script>
          setTimeout(()=>{$(".MensajeError2").slideUp();},3000);
        </script> 
      <%end if%>

      <div class="Delete">
      <h2 class="RojoCritico">¿Esta seguro de que desea eliminar al usuario del sistema?</h2>
      <img src="../System/Images/Warning.png" height: "200" width="200">
      <br>
      <a href="Edit.asp"><button style="height:50px;width:150px">No, volver a la página anterior</button></a>
      <br>
      <br>
      <input type="checkbox" id="DeleteConfirmacion"> Confirmar la acción</input>
      <br>
      <form method="POST" action="Delete.asp" id="FormDelete">
      <input type="hidden" value="1" name="DoAct">
      </form>
      <input type="submit" class="RojoCritico" id="BotonDelete" form="FormDelete" value="Si,eliminar"></button>
      <p id="Surprise" class="RojoCritico">Se debe confirmar la acción para proceder</p>
      <p> Se eliminara al usuario Matias Sanchez.</p>
      <p> Una vez eliminado el usuario, se lo conservara en la BD por 30 días.</p>
      <p> La excepción a esto se dara si se eliminan varios usuarios.</p>
    </div>
    </div>
  </body>

</html>
