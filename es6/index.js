"use strict";

var SubContent = require("docxtemplater").SubContent;
var XmlTemplater = require("docxtemplater").XmlTemplater;
var FileTypeConfig = require("docxtemplater").FileTypeConfig;
var DocUtils = require("./docUtils");

class FootnoteModule {
	constructor(options) { 
		this.options = options || {}; 
	}

	
	handleEvent(event) {
		if (event === "rendering-file") {
			var gen = this.manager.getInstance("gen");
			this.zip = gen.zip;
			this.addFootNotes();
			return gen;
		} else if (event === "rendered") {
			return this.finished();
		}
	}

	updateFile(fileName, data, options = {}) {
		this.zip.remove(fileName);
		return this.zip.file(fileName, data, options);
	}

	addFootNotes() {
		var file = this.zip.files["word/footnotes.xml"];
		var xmlString = DocUtils.decodeUtf8(file.asText());

		var output = "";
		var footnotes = this.options.footnotes;
		for (var counter = 0; counter < footnotes.length; counter++) {
			var referenceNumber = counter + 1
			output += "<w:footnote w:id='" + referenceNumber + "'><w:p><w:pPr><w:pStyle w:val='FootnoteText'/></w:pPr><w:r><w:rPr><w:rStyle w:val='FootnoteReference'/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space='preserve'>" + footnotes[counter] + "</w:t></w:r></w:p></w:footnote>"
		}

		xmlString = xmlString.replace("</w:footnotes>", output + "</w:footnotes>")
		this.updateFile("word/footnotes.xml",xmlString);
	}

	get(data) {
		return null;
	}

	handle(type, data) {
		return null;
	}

	finished() {}

	on(event, data) {
		if (event === "error") {
			throw data;
		}
	}

	replaceBy(text, outsideElement) {
		return text;
	}
	
	replaceTag() {
		return "";
	}
}

module.exports = FootnoteModule;
