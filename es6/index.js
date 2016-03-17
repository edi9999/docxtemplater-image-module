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
			// this.addStyles();
			// this.addStylesWithEffects();
			// this.updateSettings();
			// this.loadFootNoteRels();
			// this.addFootNoteContentType();
			// this.addFootNoteRels();
			// this.createFootNotesFile();
			// this.addEndNoteRels();
			// this.addEndNotesContentType()
			// this.createEndNotesFile();
			this.addFootNotes();
			return gen;
		} else if (event === "rendered") {
			return this.finished();
		}
	}

	loadFootNoteRels() {
		var file = this.zip.files[`word/_rels/${this.endFileName}.xml.rels`] || this.zip.files["word/_rels/document.xml.rels"];
		if (file === undefined) { return; }
		var content = DocUtils.decodeUtf8(file.asText());
		this.xmlDoc = DocUtils.str2xml(content);
		// Get all Rids
		var RidArray = [];
		var iterable = this.xmlDoc.getElementsByTagName("Relationship");
		for (var i = 0, tag; i < iterable.length - 1; i++) {
			tag = iterable[i];
			RidArray.push(parseInt(tag.getAttribute("Id").substr(3), 10));
		}
		this.maxRid = DocUtils.maxArray(RidArray);
		return this;
	}

	addFootNoteContentType() {
		var content = this.zip.files["[Content_Types].xml"].asText();
		var xmlDoc = DocUtils.str2xml(content);
		var addTag = true;
		
		var types = xmlDoc.getElementsByTagName("Types")[0];
		var newTag = xmlDoc.createElement("Override");
		newTag.namespaceURI = null;
		newTag.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml");
		newTag.setAttribute("PartName", "/word/footnotes.xml");
		types.appendChild(newTag);
		this.updateFile("[Content_Types].xml", DocUtils.encodeUtf8(DocUtils.xml2Str(xmlDoc)));
	}

	addEndNotesContentType() {
		var content = this.zip.files["[Content_Types].xml"].asText();
		var xmlDoc = DocUtils.str2xml(content);
		var addTag = true;
		
		var types = xmlDoc.getElementsByTagName("Types")[0];
		var newTag = xmlDoc.createElement("Override");
		newTag.namespaceURI = null;
		newTag.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml");
		newTag.setAttribute("PartName", "/word/endnotes.xml");
		types.appendChild(newTag);
		this.updateFile("[Content_Types].xml", DocUtils.encodeUtf8(DocUtils.xml2Str(xmlDoc)));
	}

	addFootNoteRels() {
		this.maxRid++;
		var file = this.zip.files[`word/_rels/${this.endFileName}.xml.rels`] || this.zip.files["word/_rels/document.xml.rels"];
		var content = DocUtils.decodeUtf8(file.asText());
		var xmlDoc = DocUtils.str2xml(content);
		var relationships = xmlDoc.getElementsByTagName("Relationships")[0];
		var newTag = xmlDoc.createElement("Relationship");
		newTag.namespaceURI = null;
		newTag.setAttribute("Id", `rId${this.maxRid}`);
		newTag.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes");
		newTag.setAttribute("Target", 'footnotes.xml');
		relationships.appendChild(newTag);
		this.updateFile(file.name, DocUtils.encodeUtf8(DocUtils.xml2Str(xmlDoc)));
	}

	addEndNoteRels() {
		this.maxRid++;
		var file = this.zip.files[`word/_rels/${this.endFileName}.xml.rels`] || this.zip.files["word/_rels/document.xml.rels"];
		var content = DocUtils.decodeUtf8(file.asText());
		var xmlDoc = DocUtils.str2xml(content);
		var relationships = xmlDoc.getElementsByTagName("Relationships")[0];
		var newTag = xmlDoc.createElement("Relationship");
		newTag.namespaceURI = null;
		newTag.setAttribute("Id", `rId${this.maxRid}`);
		newTag.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes");
		newTag.setAttribute("Target", 'endnotes.xml');
		relationships.appendChild(newTag);
		this.updateFile(file.name, DocUtils.encodeUtf8(DocUtils.xml2Str(xmlDoc)));
	}


	createFootNotesFile() {
		var prefix = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes mc:Ignorable="w14 wp14" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
		var content = this.addFootNotes();
		var suffix = '</w:footnotes>';
		var xmlString = prefix + content + suffix;

		this.zip.file("word/footnotes.xml",xmlString, {});
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

	createEndNotesFile() {
		var xmlString = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes mc:Ignorable="w14 wp14" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:endnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>'
		this.zip.file("word/endnotes.xml",xmlString, {});
	}

	addStyles() {
		var file = this.zip.files["word/styles.xml"];
		var xmlString = DocUtils.decodeUtf8(file.asText());
		xmlString = xmlString.replace("</w:styles>", "<w:style w:styleId='FootnoteText' w:type='paragraph'><w:name w:val='footnote text'/><w:basedOn w:val='Normal'/><w:link w:val='FootnoteTextChar'/><w:uiPriority w:val='99'/><w:unhideWhenUsed/><w:pPr><w:spacing w:line='240' w:lineRule='auto'/></w:pPr></w:style><w:style w:customStyle='1' w:styleId='FootnoteTextChar' w:type='character'><w:name w:val='Footnote Text Char'/><w:basedOn w:val='DefaultParagraphFont'/><w:link w:val='FootnoteText'/><w:uiPriority w:val='99'/></w:style><w:style w:styleId='FootnoteReference' w:type='character'><w:name w:val='footnote reference'/><w:basedOn w:val='DefaultParagraphFont'/><w:uiPriority w:val='99'/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val='superscript'/></w:rPr></w:style></w:styles>")
		this.updateFile("word/styles.xml",xmlString);

	}

	addStylesWithEffects() {
		var file = this.zip.files["word/stylesWithEffects.xml"];
		var xmlString = DocUtils.decodeUtf8(file.asText());
		xmlString = xmlString.replace("</w:styles>", "<w:style w:styleId='FootnoteText' w:type='paragraph'><w:name w:val='footnote text'/><w:basedOn w:val='Normal'/><w:link w:val='FootnoteTextChar'/><w:uiPriority w:val='99'/><w:unhideWhenUsed/><w:pPr><w:spacing w:line='240' w:lineRule='auto'/></w:pPr></w:style><w:style w:customStyle='1' w:styleId='FootnoteTextChar' w:type='character'><w:name w:val='Footnote Text Char'/><w:basedOn w:val='DefaultParagraphFont'/><w:link w:val='FootnoteText'/><w:uiPriority w:val='99'/></w:style><w:style w:styleId='FootnoteReference' w:type='character'><w:name w:val='footnote reference'/><w:basedOn w:val='DefaultParagraphFont'/><w:uiPriority w:val='99'/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val='superscript'/></w:rPr></w:style></w:styles>")
		this.updateFile("word/stylesWithEffects.xml",xmlString);

	}

	updateSettings() {
		var file = this.zip.files["word/settings.xml"];
		var xmlString = DocUtils.decodeUtf8(file.asText());
		xmlString = xmlString.replace("</w:settings>","<w:footnotePr><w:footnote w:id='-1'/><w:footnote w:id='0'/></w:footnotePr><w:endnotePr><w:endnote w:id='-1'/><w:endnote w:id='0'/></w:endnotePr></w:settings>");
		// var xmlDoc = DocUtils.str2xml(xmlString);
		// var settings = xmlDoc.getElementsByTagName("w:settings")[0];
		// var rsids = xmlDoc.getElementsByTagName("w:rsids")[0];
		// var newTag = xmlDoc.createElement("w:rsid");
		// newTag.namespaceURI = null;
		// newTag.setAttribute("w:val", "002A7BE7");
		// rsids.appendChild(newTag);
		// this.updateFile("word/settings.xml",DocUtils.encodeUtf8(DocUtils.xml2Str(xmlDoc)));
		this.updateFile("word/settings.xml",xmlString);
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
