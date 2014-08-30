var Handlebars = require('handlebars');
var JSZip = require('jszip');
var fs = require('fs');

var templateData = require('./demo/data').data;

fs.readFile('demo/template.docx', function(erorr, data) {
  var zip = new JSZip(data);
  var templateText = zip.file('word/document.xml').asText();
  var template = Handlebars.compile(templateText);
  var compiled = template(templateData);

  zip.file('word/document.xml', compiled);

  // @todo urls and images, check document.xml.rels

  var buffer = zip.generate({type: 'nodebuffer'});
  fs.writeFile('demo/out_template.docx', buffer);
});
