
//index.asp
$(document).ready(function() {
  //Separa por color las filas
  $("tbody tr:odd td").addClass("tdOdd");
  $("tbody tr:even td").addClass("tdEven");

  //Verificacion keydown
  $('.FormAlfaNum').keydown(function (event) {
    keyPress = event.keyCode; //Guarda en variable el codigo de la tecla apretada
    if (keyPress == 8) {} //Revisa que el keyCode sea backspace
    else if (keyPress > 47 && keyPress < 59 || keyPress > 95 && keyPress < 106 || keyPress > 64 && keyPress < 91) { console.log(keyPress);}//Revisa reglas especificas del input
    else { //Si no se cumplieron las reglas, se cancela el evento
      console.log(keyPress);
      event.preventDefault();
  }})

  
$("#DetEdit").on("click",()=>{
  $("#FormDetails").submit();
  window.location.href = "Edit.asp"
})

});

//Mensaje usuario eliminado
/*function Mensaje(Codigo) {
    $(".MensajePost").css({
      "text-align": "center",
      "margin": "0% 0% 5% 0%",
      "padding": "2%",
      "border": "1px solid black",
    });
    
    if(Codigo == 1){
      $(".MensajePost").css({
        "background": "#22bf95",
      });
      $(".MensajePost p").text("El usuario se elimino exitosamente");
    }
    else if(Codigo == 2){
      $(".MensajePost").css({
        "background": "#da2f2f",
        "color": "white"
      });
      $(".MensajePost p").text("Error");
    }
    else if(Codigo == 3){
      $(".MensajePost").css({
        "background": "#dd743c",
        "color": "white"
      });
      $(".MensajePost p").text("Error");
    }
    $(".MensajePost").show();
    setTimeout(()=>{$(".MensajePost").fadeOut();},3000);
}*/

//Details.asp
