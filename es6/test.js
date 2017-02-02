"use strict";
/* eslint-disable no-console */

const fs = require("fs");
const Docxtemplater = require("docxtemplater");
const path = require("path");
const JSZip = require("jszip");
const ImageModule = require("./index.js");
const testutils = require("docxtemplater/js/tests/utils");
const shouldBeSame = testutils.shouldBeSame;

const fileNames = [
	"imageExample.docx",
	"imageHeaderFooterExample.docx",
	"imageLoopExample.docx",
	"expectedNoImage.docx",
	"expectedHeaderFooter.docx",
	"expectedOneImage.docx",
	"expectedCentered.docx",
	"expectedLoopCentered.docx",
	"withoutRels.docx",
	"expectedWithoutRels.docx",
	"tagImage.pptx",
	"expectedTagImage.pptx",
];

beforeEach(function () {
	this.opts = {
		getImage: function (tagValue) {
			return fs.readFileSync(tagValue, "binary");
		},
		getSize: function () {
			return [150, 150];
		},
		centered: false,
	};

	this.loadAndRender = function () {
		const fileType = (testutils.pptX[this.name]) ? "pptx" : "docx";
		const file = (fileType === "pptx") ? testutils.pptX[this.name] : testutils.docX[this.name];
		this.doc = new Docxtemplater();
		this.doc.setOptions({fileType});
		this.opts.fileType = fileType;
		const inputZip = new JSZip(file.loadedContent);
		this.doc.loadZip(inputZip).setData(this.data);
		const imageModule = new ImageModule(this.opts);
		this.doc.attachModule(imageModule);
		this.renderedDoc = this.doc.render();
		const doc = this.renderedDoc;
		shouldBeSame({doc, expectedName: this.expectedName});
	};
});

function testStart() {
	describe("{%image}", function () {
		it("should work with one image", function () {
			this.name = "imageExample.docx";
			this.expectedName = "expectedOneImage.docx";
			this.data = {image: "examples/image.png"};
			this.fileType = "docx";
			this.loadAndRender();
		});

		it("should work without initial rels", function () {
			this.name = "withoutRels.docx";
			this.expectedName = "expectedWithoutRels.docx";
			this.data = {image: "examples/image.png"};
			this.loadAndRender();
		});

		it("should work with image tag == null", function () {
			this.name = "imageExample.docx";
			this.expectedName = "expectedNoImage.docx";
			this.data = {};
			this.loadAndRender();
		});

		it("should work with centering", function () {
			this.name = "imageExample.docx";
			this.expectedName = "expectedCentered.docx";
			this.opts.centered = true;
			this.data = {image: "examples/image.png"};
			this.loadAndRender();
		});

		it("should work with loops", function () {
			this.name = "imageLoopExample.docx";
			this.expectedName = "expectedLoopCentered.docx";
			this.opts.centered = true;
			this.data = {images: ["examples/image.png", "examples/image2.png"]};
			this.loadAndRender();
		});

		it("should work with image in header/footer", function () {
			this.name = "imageHeaderFooterExample.docx";
			this.expectedName = "expectedHeaderFooter.docx";
			this.data = {image: "examples/image.png"};
			this.loadAndRender();
		});

		it("should work with PPTX documents", function () {
			this.name = "tagImage.pptx";
			this.expectedName = "expectedTagImage.pptx";
			this.data = {image: "examples/image.png"};
			this.loadAndRender();
		});
	});
}

testutils.setExamplesDirectory(path.resolve(__dirname, "..", "examples"));
testutils.setStartFunction(testStart);
fileNames.forEach(function (filename) {
	if (filename.indexOf(".pptx") === filename.length - 5) {
		testutils.loadFile(filename, testutils.loadPptx);
	}
	else {
		testutils.loadFile(filename, testutils.loadDocx);
	}
});
testutils.start();
