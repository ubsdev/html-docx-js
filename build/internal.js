var _, documentTemplate, fs, utils;

fs = require('fs');

documentTemplate = require('./templates/document');

utils = require('./utils');

_ = {
  merge: require('lodash.merge')
};

module.exports = {
  generateDocument: function(zip) {
    var buffer;
    buffer = zip.generate({
      type: 'arraybuffer'
    });
    if (global.Blob) {
      return new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
    } else if (global.Buffer) {
      return new Buffer(new Uint8Array(buffer));
    } else {
      throw new Error("Neither Blob nor Buffer are accessible in this environment. " + "Consider adding Blob.js shim");
    }
  },
  renderDocumentFile: function(documentOptions) {
    var a, pageHeight, pageSizes, pageWidth, templateData;
    if (documentOptions == null) {
      documentOptions = {};
    }
    templateData = {};
    if (documentOptions === null) {
      documentOptions = {};
    }
    pageSizes = ['Letter', 'Tabloid', 'Legal', 'Statement', 'Executive', 'A3', 'A4', 'A5'];
    pageWidth = 12240;
    pageHeight = 15840;
    if (documentOptions.height) {
      pageHeight = documentOptions.height;
    }
    if (documentOptions.width) {
      pageWidth = documentOptions.width;
    }
    if (!documentOptions.height && !documentOptions.height) {
      if (documentOptions.size) {
        a = pageSizes.indexOf(documentOptions.size);
        console.log(a);
        if (a === -1) {
          throw new Error('Size should be ' + pageSizes.toString());
        }
      } else {
        documentOptions.size = 'letter';
      }
      switch (documentOptions.size) {
        case 'Letter':
          pageWidth = 12240;
          pageHeight = 15840;
          break;
        case 'Tabloid':
          pageWidth = 15840;
          pageHeight = 24480;
          break;
        case 'Legal':
          pageWidth = 12240;
          pageHeight = 20160;
          break;
        case 'Statement':
          pageWidth = 7920;
          pageHeight = 12240;
          break;
        case 'Executive':
          pageWidth = 10437.165;
          pageHeight = 15120;
          break;
        case 'A3':
          pageWidth = 16837.795;
          pageHeight = 23811.024;
          break;
        case 'A4':
          pageWidth = 11905.511;
          pageHeight = 16837.795;
          break;
        case 'A5':
          pageWidth = 8390.551;
          pageHeight = 11905.511;
          break;
        default:
          pageHeight = 15840;
          pageWidth = 12240;
      }
    }
    templateData = _.merge({
      margins: {
        top: 1440,
        right: 1440,
        bottom: 1440,
        left: 1440,
        header: 720,
        footer: 720,
        gutter: 0
      }
    }, (function() {
      switch (documentOptions.orientation) {
        case 'landscape':
          return {
            height: pageHeight,
            width: pageWidth,
            orient: 'landscape'
          };
        default:
          return {
            width: pageWidth,
            height: pageHeight,
            orient: 'portrait'
          };
      }
    })(), {
      margins: documentOptions.margins
    });
    return documentTemplate(templateData);
  },
  addFiles: function(zip, htmlSource, documentOptions) {
    zip.file('[Content_Types].xml', fs.readFileSync(__dirname + '/assets/content_types.xml'));
    zip.folder('_rels').file('.rels', fs.readFileSync(__dirname + '/assets/rels.xml'));
    return zip.folder('word').file('document.xml', this.renderDocumentFile(documentOptions)).file('afchunk.mht', utils.getMHTdocument(htmlSource)).folder('_rels').file('document.xml.rels', fs.readFileSync(__dirname + '/assets/document.xml.rels'));
  }
};
