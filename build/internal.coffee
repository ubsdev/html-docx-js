fs = require 'fs'
documentTemplate = require './templates/document'
utils = require './utils'
_ = merge: require 'lodash.merge'

module.exports =
  generateDocument: (zip) ->
    buffer = zip.generate(type: 'arraybuffer')
    if global.Blob
      new Blob [buffer],
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    else if global.Buffer
      new Buffer new Uint8Array(buffer)
    else
      throw new Error "Neither Blob nor Buffer are accessible in this environment. " +
        "Consider adding Blob.js shim"

  renderDocumentFile: (documentOptions = {}) ->

    templateData = {}
    if documentOptions == null
      documentOptions = {}
    pageSizes = [
      'Letter'
      'Tabloid'
      'Legal'
      'Statement'
      'Executive'
      'A3'
      'A4'
      'A5'
    ]
    pageWidth = 12240
    pageHeight = 15840
    if documentOptions.height
      pageHeight = documentOptions.height
    if documentOptions.width
      pageWidth = documentOptions.width
    if !documentOptions.height and !documentOptions.height
      if documentOptions.size
        a = pageSizes.indexOf(documentOptions.size)
        console.log a
        if a == -1
          throw new Error('Size should be ' + pageSizes.toString())
      else
        documentOptions.size = 'letter'
      switch documentOptions.size
        when 'Letter'
          pageWidth = 12240
          pageHeight = 15840
        when 'Tabloid'
          pageWidth = 15840
          pageHeight = 24480
        when 'Legal'
          pageWidth = 12240
          pageHeight = 20160
        when 'Statement'
          pageWidth = 7920
          pageHeight = 12240
        when 'Executive'
          pageWidth = 10437.165
          pageHeight = 15120
        when 'A3'
          pageWidth =  16837.795
          pageHeight = 23811.024
        when 'A4'
          pageWidth =  11905.511
          pageHeight = 16837.795
        when 'A5'
          pageWidth = 8390.551
          pageHeight = 11905.511
        else
          pageHeight = 15840
          pageWidth = 12240

    templateData = _.merge margins:
      top: 1440
      right: 1440
      bottom: 1440
      left: 1440
      header: 720
      footer: 720
      gutter: 0
    ,
      switch documentOptions.orientation
        when 'landscape' then height: pageHeight, width: pageWidth, orient: 'landscape'
        else width: pageWidth, height: pageHeight, orient: 'portrait'
    ,
      margins: documentOptions.margins

    documentTemplate(templateData)

  addFiles: (zip, htmlSource, documentOptions) ->
    zip.file '[Content_Types].xml', fs.readFileSync __dirname + '/assets/content_types.xml'
    zip.folder('_rels').file '.rels', fs.readFileSync __dirname + '/assets/rels.xml'
    zip.folder 'word'
      .file 'document.xml', @renderDocumentFile documentOptions
      .file 'afchunk.mht', utils.getMHTdocument htmlSource
      .folder '_rels'
        .file 'document.xml.rels', fs.readFileSync __dirname + '/assets/document.xml.rels'
