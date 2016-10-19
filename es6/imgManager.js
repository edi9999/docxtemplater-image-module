"use strict";

const DocUtils = require("./docUtils");

const imageExtensions = ["gif", "jpeg", "jpg", "emf", "png"];
const imageListRegex = /[^.]+\.([^.]+)/;

module.exports = class ImgManager {
	constructor(zip, fileName) {
		this.zip = zip;
		this.fileName = fileName;
		this.endFileName = this.fileName.replace(/^.*?([a-z0-9]+)\.xml$/, "$1");
	}
	getImageList() {
		const imageList = [];
		Object.keys(this.zip.files).forEach(function (path) {
			const extension = path.replace(imageListRegex, "$1");
			if (imageExtensions.indexOf(extension) >= 0) {
				imageList.push({path: path, files: this.zip.files[path]});
			}
		});
		return imageList;
	}
	setImage(fileName, data, options) {
		options = options || {};
		this.zip.remove(fileName);
		return this.zip.file(fileName, data, options);
	}
	hasImage(fileName) {
		return this.zip.files[fileName] != null;
	}
	loadImageRels() {
		const file = this.zip.files[`word/_rels/${this.endFileName}.xml.rels`];
		let content;
		if(file!=null){
			content = DocUtils.decodeUtf8(file.asText());
		}else{
			content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>";
			this.addPartRels(`/word/_rels/${this.endFileName}.xml.rels`,"application/vnd.openxmlformats-package.relationships+xml")
		}
		this.xmlDoc = DocUtils.str2xml(content);
		// Get all Rids
		const RidArray = [0];
		const iterable = this.xmlDoc.getElementsByTagName("Relationship");
		for (let i = 0, tag; i < iterable.length; i++) {
			tag = iterable[i];
			const id = tag.getAttribute("Id");
			if (/^rId[0-9]+$/.test(id)) {
				RidArray.push(parseInt(id.substr(3), 10));
			}
		}
		this.maxRid = DocUtils.maxArray(RidArray);
		this.imageRels = [];
		return this;
	}
// Add an extension type in the [Content_Types.xml], is used if for example you want word to be able to read png files (for every extension you add you need a contentType)
	addExtensionRels(contentType, extension) {
		const content = this.zip.files["[Content_Types].xml"].asText();
		const xmlDoc = DocUtils.str2xml(content);
		let addTag = true;
		const defaultTags = xmlDoc.getElementsByTagName("Default");
		for (let i = 0, tag; i < defaultTags.length; i++) {
			tag = defaultTags[i];
			if (tag.getAttribute("Extension") === extension) { addTag = false; }
		}
		if (addTag) {
			const types = xmlDoc.getElementsByTagName("Types")[0];
			const newTag = xmlDoc.createElement("Default");
			newTag.namespaceURI = null;
			newTag.setAttribute("ContentType", contentType);
			newTag.setAttribute("Extension", extension);
			types.appendChild(newTag);
			return this.setImage("[Content_Types].xml", DocUtils.encodeUtf8(DocUtils.xml2Str(xmlDoc)));
		}
	}

	addPartRels(partName, contentType) {
		const content = this.zip.files["[Content_Types].xml"].asText();
		const xmlDoc = DocUtils.str2xml(content);
		let addTag = true;
		const defaultTags = xmlDoc.getElementsByTagName("Override");
		for (let i = 0, tag; i < defaultTags.length; i++) {
			tag = defaultTags[i];
			if (tag.getAttribute("PartName") === partName) { addTag = false; }
		}
		if (addTag) {
			const types = xmlDoc.getElementsByTagName("Types")[0];
			const newTag = xmlDoc.createElement("Override");
			newTag.namespaceURI = null;
			newTag.setAttribute("ContentType", contentType);
			newTag.setAttribute("PartName", partName);
			types.appendChild(newTag);
			return this.setImage("[Content_Types].xml", DocUtils.encodeUtf8(DocUtils.xml2Str(xmlDoc)));
		}
	}
	// Adding an image and returns it's Rid
	addImageRels(imageName, imageData, i) {
		if (i == null) {
			i = 0;
		}
		const realImageName = i === 0 ? imageName : imageName + `(${i})`;
		if ((this.zip.files[`word/media/${realImageName}`] != null)) {
			return this.addImageRels(imageName, imageData, i + 1);
		}
		this.maxRid++;
		const file = {
			name: `word/media/${realImageName}`,
			data: imageData,
			options: {
				base64: false,
				binary: true,
				compression: null,
				date: new Date(),
				dir: false,
			},
		};
		this.zip.file(file.name, file.data, file.options);
		const extension = realImageName.replace(/[^.]+\.([^.]+)/, "$1");
		this.addExtensionRels(`image/${extension}`, extension);
		const relationships = this.xmlDoc.getElementsByTagName("Relationships")[0];
		const newTag = this.xmlDoc.createElement("Relationship");
		newTag.namespaceURI = null;
		newTag.setAttribute("Id", `rId${this.maxRid}`);
		newTag.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
		newTag.setAttribute("Target", `media/${realImageName}`);
		relationships.appendChild(newTag);
		this.setImage(`word/_rels/${this.endFileName}.xml.rels`, DocUtils.encodeUtf8(DocUtils.xml2Str(this.xmlDoc)));
		return this.maxRid;
	}
	getImageName(id) {
		id = id || 0;
		const nameCandidate = "Copie_" + id + ".png";
		const fullPath = this.getFullPath(nameCandidate);
		if (this.hasImage(fullPath)) {
			return this.getImageName(id + 1);
		}
		return nameCandidate;
	}
	getFullPath(imgName) { return `word/media/${imgName}`; }
	// This is to get an image by it's rId (returns null if no img was found)
	getImageByRid(rId) {
		const relationships = this.xmlDoc.getElementsByTagName("Relationship");
		for (let i = 0, relationship; i < relationships.length; i++) {
			relationship = relationships[i];
			const cRId = relationship.getAttribute("Id");
			if (rId === cRId) {
				const path = relationship.getAttribute("Target");
				if (path.substr(0, 6) === "media/") {
					return this.zip.files[`word/${path}`];
				}
				throw new Error("Rid is not an image");
			}
		}
		throw new Error("No Media with this Rid found");
	}
};
