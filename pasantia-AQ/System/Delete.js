$(document).ready(function () {
  $('#Surprise').hide(); //Esconder pedido de confirmacion

  $('#BotonDelete').click((event)=> {
    if($('#DeleteConfirmacion').is(':not(:checked)')){
      event.preventDefault();
      $('#Surprise').show();
    } //Mostrar advertencia de confirmacion
  }).on('mouseleave', function () {
    if($('#DeleteConfirmacion').is(':not(:checked)'))
    $('#Surprise').fadeOut(); //Esconder advertencia de confirmacion
  })
});
