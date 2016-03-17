This is the master branch, which is not completely stable yet (version 1). No npm module has been released to npm yet. For up to date documentation, please switch to the 0.x branch.

This version is only compatible with docxtemplater version 2

The v1 has added two mandatory arguments: `opts.getImage` and `opts.getSize`. See Usage (v1)

[![Build Status](https://travis-ci.org/open-xml-templating/docxtemplater-image-module.svg?branch=master&style=flat)](https://travis-ci.org/open-xml-templating/docxtemplater-image-module) [![Download count](http://img.shields.io/npm/dm/docxtemplater-image-module.svg?style=flat)](https://www.npmjs.org/package/docxtemplater-image-module) [![Current tag](http://img.shields.io/npm/v/docxtemplater-image-module.svg?style=flat)](https://www.npmjs.org/package/docxtemplater-image-module) [![Issues closed](http://issuestats.com/github/open-xml-templating/docxtemplater-image-module/badge/issue?style=flat)](http://issuestats.com/github/open-xml-templating/docxtemplater-image-module)

Installation:
=============

You will need docxtemplater v1: `npm install docxtemplater`

install this module: `npm install docxtemplater-image-module`

Usage (v1)
==========

Your docx should contain the text: `{%image}`

```
var ImageModule=require('docxtemplater-image-module')

var opts = {}
opts.centered = false;
opts.getImage=function(tagValue, tagName) {
    return fs.readFileSync(tagValue,'binary');
}

opts.getSize=function(img,tagValue, tagName) {
    return [150,150];
}

var imageModule=new ImageModule(opts);

var docx=new DocxGen()
    .attachModule(imageModule)
    .load(content)
    .setData({image:'examples/image.png'})
    .render()

var buffer= docx
        .getZip()
        .generate({type:"nodebuffer"})

fs.writeFile("test.docx",buffer);
```

To understand what `img`, `tagValue`, `tagName` mean, lets take an example :

If your template is :

```
{%myImage}

and your data:

{
    "myImage":'sampleImage.png'
}

tagValue will be equal to "sampleImage.png", tagName will be equal to "myImage" and img will be what ever the getImage function returned
```

One of the most useful cases of this is to set the images to be the size of that image.

For this, you will need to install the [npm package 'image-size'](https://www.npmjs.com/package/image-size) then, write:

```
opts = {centered:false}
opts.getImage=function(tagValue) {
    return fs.readFileSync(tagValue,'binary');
}
opts.getSize=function(img) {
   sizeOf=require('image-size');
   sizeObj=sizeOf(img);
   console.log(sizeObj);
   return [sizeObj.width,sizeObj.height];
}
imageModule=new ImageModule(opts)
```

Centering images
================

You can center the images using opts.centered=true instead

Size and path based on placeholder
==================================

You can have customizable image loader using the template's placeholder name.

```
opts.getImage = function (tagValue, tagName) {
    if(tagName === 'logo')
        return fs.readFileSync(__dirname + '/logos/' + tagValue);

    return fs.readFileSync(__dirname + '/images/' + tagValue);
};
```

The same thing can be used to customize image size.

```
opts.getSize = function (img, tagValue, tagName) {
    if(tagName === 'logo')
        return [100, 100];

    return [300, 300];
};
```

Notice
======

For the imagereplacer to work, the image tag: `{%image}` needs to be in its own `<w:p>`, so that means that you have to put a new line after and before the tag.

Testing
=======

You can test that everything works fine using the command `mocha`. This will also create some docx files under the root directory that you can open to check if the generation was correct.

Instructions for Soomo Dev
=======================

You can build a release for the browser with the following commands

```
npm install -g gulp jasmine-node uglify-js browserify
npm install
browserify -x docxtemplater ./es6/index.js -s FootnoteModule > <%soomo-core-directory%>/vendor/assets/javascripts/docxtemplater-footnote.v1.0.0.js
```
* Use branch turabian-footnotes on core
* To get things setup you could use this docx [template](http://assets.soomopublishing.com.s3.amazonaws.com/courses/USHistory2/HIS114MilestoneOneTemplate.Turabian_Tarun_runs.docx)



