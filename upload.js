function readURL(input) {
  if (input.files && input.files[0]) {
    var reader = new FileReader();
    reader.onload = function (e) {
      $('.file-upload-wrap').hide();
      $('.file-upload-file').html(e.target.name);
      $('.file-upload-content').show();
      $('.file-title').html(input.files[0].name);
    };
    reader.readAsDataURL(input.files[0]);
  } else {
    removeUpload();
  }
}

function removeUpload() {
  $('.file-upload-input').val('');
  $('.file-upload-content').hide();
  $('.file-upload-wrap').show();
}
$('.file-upload-wrap').bind('dragover', function () {
  $('.file-upload-wrap').addClass('file-dropping');
});
$('.file-upload-wrap').bind('dragleave', function () {
  $('.file-upload-wrap').removeClass('file-dropping');
});
