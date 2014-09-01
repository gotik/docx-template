var Handlebars = require('handlebars');
var JSZip = require('jszip');
var fs = require('fs');
var xmldoc = require('xmldom');
var DOMParser = xmldoc.DOMParser;
var XMLSerializer = xmldoc.XMLSerializer;
var domParser = new DOMParser();
var xmlSerializer = new XMLSerializer();

var templateData = require('./demo/data').data;

fs.readFile('demo/template.docx', function(erorr, data) {
  var zip = new JSZip(data);
  var documentText = zip.file('word/document.xml').asText();
  var relationsText = zip.file('word/_rels/document.xml.rels').asText();
  var template = Handlebars.compile(documentText);
  var compiled = template(templateData);

  zip.file('word/document.xml', compiled);


  var xml = domParser.parseFromString(relationsText, 'text/xml');
  var root = xml.getElementsByTagName('Relationships')[0];

  // demo imgs
  for (var i=1; i<10; i++) {
    root.appendChild(createImage(xml, i));
  }

  zip.file('word/_rels/document.xml.rels', xmlSerializer.serializeToString(xml));


  var buffer = zip.generate({type: 'nodebuffer'});
  fs.writeFile('demo/out_template.docx', buffer);
});


function createImage(xml, i) {
  var relation = xml.createElement('Relationship');
  relation.setAttribute('Target', 'media/image01.jpg');
  relation.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
  relation.setAttribute('Id', 'dId'+i);

  return relation;
}