function convFecha(fecha) {
  const date = fecha.split('-');
  return `${date[2]}/${date[1]}/${date[0]}`;
}

function generate(template = "https://github.com/jcoppis/agronomia/raw/master/templates/formEgresados.docx") {
  const formElem = new FormData(document.querySelector('form'));
  let data = {};
  for(const val of formElem.entries()) {
    data[val[0]] = val[1];
  }
  data['nota'] = 10;
  data['nroActa'] = 1006;
  data['materia'] = 'Apicultura dtdi thdithdithdi hdithdithdi hdithdithdi';
  data['fechaEgreso'] = convFecha('2020-05-02');
  data['prom'] = '8.5';
  data['date'] = '2020-01-01';
  data['textoSecAcademico'] = 'Sra. Secretaria Académica';
  data['secAcademico'] = 'Mgtr. Liliana M. Gallez';

  saveToFile(data, template);
}

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

function saveToFile(data, template) {
  loadFile(template, function (error, content) {
    if (error) { throw error };
    var zip = new PizZip(content);
    var doc = new window.docxtemplater().loadZip(zip)
    doc.setData(data);
    try {
      // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
      doc.render()
    }
    catch (error) {
      // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object).
      var e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties,
      }
      console.log(JSON.stringify({ error: e }));
      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors.map(function (error) {
          return error.properties.explanation;
        }).join("\n");
        console.log('errorMessages', errorMessages);
        // errorMessages is a humanly readable message looking like this :
        // 'The tag beginning with "foobar" is unopened'
      }
      throw error;
    }
    var out = doc.getZip().generate({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }) //Output the document using Data-URI
    saveAs(out, "output.docx")
  })
}