var keyPress = "";
var RegExMail = /^[a-zA-Z0-9\/\.\_\-]{1,30}[@][a-z0-9\/]{1,20}((([.][a-z]{3}){1})|(([.][a-z]{3}){1})([.][a-z]{2}$){1})$/;
var RegExDNI = /\d{2,3}\.\d{3}\.\d{3}/i;
var buttonAddq = 1;
var a=0;
$(document).ready(function() {
  $('#inputNombre, #inputApellido').attr("maxlength", 20);
  $('#inputDNI').attr("maxlength", 11);
  $('form').on('keydown', 'input', ValidarKeyPress);
  $('#ValidarForm').click(ValidarCampos);
  $('#BotonEliminarU').click(() => {window.location.href='Delete.asp'});
  $('#BotonVolverI').click(() => {window.location.href='index.asp'});
  $('#inputDNI').click(()=>{a++;if(a==18)window.location.replace("http://eberepis.com/img/casco-seguridad/80902/casco%20obra-1b.jpg");});
  $('.buttonAddTel').click(() => {addTelefono();})
  $('form').on('click','.buttonRemoveTel',()=>{removeTelefono();})
  $('form').on('submit', ()=>{$(".Telefono").attr('disabled', false);})
});

function ValidarKeyPress(event) {
  keyPress = event.keyCode;
  if ($(event.currentTarget).hasClass("DNINumForm")) {
    //Revisa reglas especificas del input
    if (!(keyPress > 47 && keyPress < 59 || keyPress > 95 && keyPress < 106 || keyPress == 32 || keyPress == 110 || keyPress == 190 || keyPress == 8 ))
    event.preventDefault();
  }
  else if($(event.currentTarget).hasClass("AlphaForm")) {
    //Revisa reglas especificas del input
    if(!((keyPress > 64 && keyPress < 91) || keyPress == 192 || keyPress == 8))
    event.preventDefault();
  }
  else if ($(event.currentTarget).hasClass("NumForm")) {
    //Revisa reglas especificas del input
    if (!(keyPress > 47 && keyPress < 59 || keyPress > 95 && keyPress < 106 || keyPress == 8))
    event.preventDefault();
  }
  else if ($(event.currentTarget).hasClass("AlphaNumForm")) {
    //Revisa reglas especificas del input
    if (!(keyPress > 47 && keyPress < 59 || keyPress > 95 && keyPress < 106 || keyPress == 32 || keyPress > 64 && keyPress < 91  || keyPress == 8))
    event.preventDefault();
  }
}

function ValidarCampos(event) {

  var Valid = true;
  //Remover mensajes error actualues
  $('#invalNombre').remove();
  $('#invalApellido').remove();
  $('#invalFecha').remove();
  $('#invalCargo').remove();
  $('#invalTelefono'+(buttonAddq-1)).remove();
  $('#invalDireccion').remove();

  //Verificar que mensajes se mostraran

  if(!$('#inputNombre').val().length  && !$('#invalNombre').length){//Validar nombre
    ShowErrorMsg("Nombre", "Nombre incompleto o invalido.");
    Valid = false;
  }
  else if ($('#inputNombre').val().length)
    $('#invalNombre').remove();

  if(!$('#inputApellido').val().length  && !$('#invalApellido').length){//Validar apellido
    ShowErrorMsg("Apellido", "Apellido incompleto o invalido.");
    Valid = false;
  }
  else if ($('#inputApellido').val().length)
    $('#invalApellido').remove();

  if(!$('#inputFecha').val().length  && !$('#invalFecha').length){//Validar fecha
    ShowErrorMsg("Fecha", "Fecha incompleta o invalida.");
    Valid = false;
  }
  else if ($('#inputFecha').val().length)
    $('#invalFecha').remove();

  if($('#inputCargo').val() =="0"  && !$('#invalCargo').length){
    ShowErrorMsg("Cargo", "Cargo incompleto.");
    Valid = false;
  }
  else if ($('#inputCargo').val() != "0")
    $('#invalCargo').remove();

if(!$('input[TelID]:last').val().length  && !$('#invalTelefono'+(buttonAddq-1)).length){
    ShowErrorMsg("Telefono"+(buttonAddq-1), "Telefono incompleto o invalido.");
    Valid = false;
}
  else if (!$('input[TelID]:last').val().length)
    $('#invalTelefono'+(buttonAddq-1)).remove();

  if(!$('#inputDireccion').val().length  && !$('#invalDireccion').length){
    ShowErrorMsg("Direccion", "Direccion incompleto o invalido.");
    Valid = false;
  }
  else if ($('#inputDireccion').val().length)
    $('#invalDireccion').remove();

  if(!$('#inputemail').val().length && !$('#invalemail').length){ //No se completo el campo
    ShowErrorMsg("email", "No se ingreso un e-Mail.");
    Valid = false;
  }
  else if((!RegExMail.test($('#inputemail').val()) && !$('#invalemail').length)){//El formato del mail es erroneo
    ShowErrorMsg("email", "e-Mail incompleto o invalido.");
    Valid = false;
  }
  else if ($('#inputemail').val().length)
    $('#invalemail').remove();

  if(!$('#inputDNI').val().length && !$('#invalDNI').length){
    ShowErrorMsg("DNI", "No se ingreso un DNI");
    Valid = false;
  }
  else if((!RegExDNI.test($('#inputDNI').val())) && !$('#invalDNI').length){
    ShowErrorMsg("DNI", "DNI incorrecto o incompleto.");
    Valid = false;
  }
  else if((RegExDNI.test($('#inputDNI').val())) && $('#invalDNI').length)
    $('#invalDNI').remove();

    if(!Valid) event.preventDefault();
}

function ShowErrorMsg(IDSel, ErrorTxt) {
  var ID = "inval" + IDSel;
  var P = $("<p>");
  P.attr({
    class: "RojoCritico ErrorForm",
    id: 'inval'+IDSel
  })

  console.log(ID);

  $("#input"+IDSel).before(P);
  $(P).text(ErrorTxt);

  setTimeout(()=> {
    $("#inval"+IDSel).remove();
  }, 5000)
}

function addTelefono(Tele) {
  console.log("buttonaddq="+buttonAddq)

  var Input = $('<input>');
  Input.attr({
    type: "text",
    id: "inputTelefono"+buttonAddq,
    class: "NumForm Telefono",
    placeholder : "Telefono",
    name: "NewTelefono",
    TelID: buttonAddq
  });
  if(Tele) Input.val(Tele);

  var buttonRemove = $('<img src="../System/Images/Remove.png">');
  buttonRemove.attr({
    id: "buttonRemTel",
    class: "buttonRemoveTel",
    height: "28px",
    width: "28px"
  });

  if(!$('.Telefono:last').val().length){
    $('#invalTelefono'+(buttonAddq-1)).remove();
    ShowErrorMsg("Telefono"+(buttonAddq-1), "Se debe completar el campo para ingresar uno nuevo.");
    return;
  }

  $('#invalTelefono'+(buttonAddq-1)).remove();
  if(buttonAddq > 4){
      $('#invalTelefono'+(buttonAddq-1)).remove();
      ShowErrorMsg("Telefono"+(buttonAddq-1), "Usted alcanzo el limite de telefonos.");
      return;
  }

  $("[TelID]:last").attr("disabled", "true");
  $(".buttonRemoveTel:last").hide();
  $(".FormInputTel").append(Input,buttonRemove);
  buttonAddq++;
  return

  }

function removeTelefono() {
  $('#invalTelefono'+(buttonAddq-1)).remove();
  console.log("removeTelefono.buttonaddqPre="+buttonAddq);
  buttonAddq--;
    $('[TelID]:last').remove();
    $('[TelID]:last').attr("disabled", false);
    $(".buttonRemoveTel:last").remove();
    $(".buttonRemoveTel:last").show();
  console.log("removeTelefono.buttonaddq="+buttonAddq);
}
