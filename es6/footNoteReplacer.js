"use strict";

var DocUtils = require("./docUtils");
var DocxQrCode = require("./docxQrCode");
var PNG = require("png-js");
var base64encode = require("./base64").encode;

module.exports = class FootNoteReplacer {
	constructor(xmlTemplater, footNoteManager) {
		this.xmlTemplater = xmlTemplater;
		this.footNoteManager = footNoteManager;
		this.imageSetter = this.imageSetter.bind(this);
		this.imgMatches = [];
		this.xmlTemplater.numQrCode = 0;
	}
	findImages() {
		this.imgMatches = DocUtils.pregMatchAll(/<w:drawing[^>]*>.*?<a:blip.r:embed.*?<\/w:drawing>/g, this.xmlTemplater.content);
		return this;
	}
	replaceImages() {
		this.qr = [];
		this.xmlTemplater.numQrCode += this.imgMatches.length;
		var iterable = this.imgMatches;
		for (var imgNum = 0, match; imgNum < iterable.length; imgNum++) {
			match = iterable[imgNum];
			this.replaceImage(match, imgNum);
		}
		return this;
	}
	imageSetter(docxqrCode) {
		if (docxqrCode.callbacked === true) { return; }
		docxqrCode.callbacked = true;
		docxqrCode.xmlTemplater.numQrCode--;
		this.footNoteManager.setImage(`word/media/${docxqrCode.imgName}`, docxqrCode.data, {binary: true});
		return this.popQrQueue(this.footNoteManager.fileName + "-" + docxqrCode.num, false);
	}
	getXmlImg(match) {
		var baseDocument = `<?xml version="1.0" ?>
		<w:document
		mc:Ignorable="w14 wp14"
		xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
			xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
			xmlns:o="urn:schemas-microsoft-com:office:office"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
			xmlns:v="urn:schemas-microsoft-com:vml"
		xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
			xmlns:w10="urn:schemas-microsoft-com:office:word"
		xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
			xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
			xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
			xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
			xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
			xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
			xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
			xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${match[0]}</w:document>
			`;
		var f = function (e) {
			if (e === "fatalError") {
				throw new Error("fatalError");
			}
		};
		return DocUtils.str2xml(baseDocument, f);
	}
	replaceImage(match, imgNum) {
		var num = parseInt(Math.random() * 10000, 10);
		var xmlImg;
		try {
			xmlImg = this.getXmlImg(match);
		}
		catch (e) {
			return;
		}
		var tagrId = xmlImg.getElementsByTagName("a:blip")[0];
		if (tagrId === undefined) { throw new Error("tagRiD undefined !"); }
		var rId = tagrId.getAttribute("r:embed");
		var tag = xmlImg.getElementsByTagName("wp:docPr")[0];
		if (tag === undefined) { throw new Error("tag undefined"); }
		// if image is already a replacement then do nothing
		if (tag.getAttribute("name").substr(0, 6) === "Copie_") { return; }
		var imgName = this.footNoteManager.getImageName();
		this.pushQrQueue(this.footNoteManager.fileName + "-" + num, true);
		var newId = this.footNoteManager.addImageRels(imgName, "");
		this.xmlTemplater.imageId++;
		var oldFile = this.footNoteManager.getImageByRid(rId);
		this.footNoteManager.setImage(this.footNoteManager.getFullPath(imgName), oldFile.data, {binary: true});
		tag.setAttribute("name", `${imgName}`);
		tagrId.setAttribute("r:embed", `rId${newId}`);
		var imageTag = xmlImg.getElementsByTagName("w:drawing")[0];
		if (imageTag === undefined) { throw new Error("imageTag undefined"); }
		var replacement = DocUtils.xml2Str(imageTag);
		this.xmlTemplater.content = this.xmlTemplater.content.replace(match[0], replacement);

		return this.decodeImage(oldFile, imgName, num, imgNum);
	}
	decodeImage(oldFile, imgName, num, imgNum) {
		var mockedQrCode = {xmlTemplater: this.xmlTemplater, imgName: imgName, data: oldFile.asBinary(), num: num};
		if (!/\.png$/.test(oldFile.name)) {
			return this.imageSetter(mockedQrCode);
		}
		return ((imgName) => {
			var base64 = base64encode(oldFile.asBinary());
			var binaryData = new Buffer(base64, "base64");
			var png = new PNG(binaryData);
			var finished = (a) => {
				png.decoded = a;
				try {
					this.qr[imgNum] = new DocxQrCode(png, this.xmlTemplater, imgName, num, this.getDataFromString);
					return this.qr[imgNum].decode(this.imageSetter);
				}
				catch (e) {
					return this.imageSetter(mockedQrCode);
				}
			};
			return png.decode(finished);
		})(imgName);
	}
};
