
document.getElementById('fileInput').addEventListener('change', function(evt) {
    alert("File selected: " + evt.target.files[0].name);
});
