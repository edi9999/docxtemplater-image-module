First module for docxtemplater.

[![Build Status](https://travis-ci.org/open-xml-templating/docxtemplater-image-module.svg?branch=master&style=flat)](https://travis-ci.org/open-xml-templating/docxtemplater-image-module)
[![Download count](http://img.shields.io/npm/dm/docxtemplater-image-module.svg?style=flat)](https://www.npmjs.org/package/docxtemplater-image-module)
[![Current tag](http://img.shields.io/npm/v/docxtemplater-image-module.svg?style=flat)](https://www.npmjs.org/package/docxtemplater-image-module)
[![Issues closed](http://issuestats.com/github/open-xml-templating/docxtemplater-image-module/badge/issue?style=flat)](http://issuestats.com/github/open-xml-templating/docxtemplater-image-module)

# Installation:

You will need docxtemplater v1: `npm install docxtemplater`

install this module: `npm install docxtemplater-image-module`

# Usage

Your docx should contain the text: `{%myImage}`

```coffee
ImageModule = require 'docxtemplater-image-module'

imageModule = new ImageModule({ centered: false })
docx = new DocxGen()
  .attachModule(imageModule)
  .load(content)
  .setData({ myImage: { path: 'examples/image.png', size: [650, 200] })
  .render()

buffer = docx
  .getZip()
  .generate( {type: 'nodebuffer' })

  fs.writeFile("test.docx",buffer);
```

# Options

 You can center the images using new ImageModule({centered:true}) instead

# Notice

 For the imagereplacer to work, the image tag: `{%image}` need to be in its own `<w:p>`, so that means that you have to put a new line after and before the tag.

# Building

 You can build the coffee into js by running `gulp` (this will watch the directory for changes)

# Testing

You can test that everything works fine using the command `mocha`. This will also create 3 docx files under the root directory that you can open to check if the docx are correct
