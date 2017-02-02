"use strict";

const DocUtils = require("docxtemplater").DocUtils;
const DOMParser = require("xmldom").DOMParser;

function isNaN(number) {
	return !(number === number);
}

const ImgManager = require("./imgManager");
const moduleName = "open-xml-templating/docxtemplater-image-module";

function getInner({part, left, right, postparsed}) {
	const xmlString = postparsed.slice(left + 1, right).reduce(function (concat, item) {
		return concat + item.value;
	}, "");
	part.off = {x: 0, y: 0};
	part.ext = {cx: 0, cy: 0};
	const xmlDoc = new DOMParser().parseFromString("<xml>" + xmlString + "</xml>");
	const off = xmlDoc.getElementsByTagName("a:off");
	if (off.length > 0) {
		part.off.x = off[0].getAttribute("x");
		part.off.y = off[0].getAttribute("y");
	}
	const ext = xmlDoc.getElementsByTagName("a:ext");
	if (ext.length > 0) {
		part.ext.cx = ext[0].getAttribute("cx");
		part.ext.cy = ext[0].getAttribute("cy");
	}
	return part;
}

class ImageModule {
	constructor(options) {
		this.options = options || {};
		if (this.options.centered == null) { this.options.centered = false; }
		if (this.options.getImage == null) { throw new Error("You should pass getImage"); }
		if (this.options.getSize == null) { throw new Error("You should pass getSize"); }
		this.qrQueue = [];
		this.imageNumber = 1;
	}
	optionsTransformer(options, docxtemplater) {
		const relsFiles = docxtemplater.zip.file(/\.xml\.rels/)
			.concat(docxtemplater.zip.file(/\[Content_Types\].xml/))
			.map((file) => file.name);
		this.fileTypeConfig = docxtemplater.fileTypeConfig;
		this.zip = docxtemplater.zip;
		options.xmlFileNames = options.xmlFileNames.concat(relsFiles);
		return options;
	}
	set(options) {
		if (options.zip) {
			this.zip = options.zip;
		}
		if (options.xmlDocuments) {
			this.xmlDocuments = options.xmlDocuments;
		}
	}
	parse(placeHolderContent) {
		const module = moduleName;
		const type = "placeholder";
		if (placeHolderContent.substring(0, 2) === "%%") {
			return {type, value: placeHolderContent.substr(2), module, centered: true};
		}
		if (placeHolderContent.substring(0, 1) === "%") {
			return {type, value: placeHolderContent.substr(1), module, centered: false};
		}
		return null;
	}
	postparse(parsed) {
		let expandTo;
		if (this.options.fileType === "pptx") {
			expandTo = "p:sp";
		}
		else {
			expandTo = this.options.centered ? "w:p" : "w:t";
		}
		return DocUtils.traits.expandToOne(parsed, {moduleName, getInner, expandTo});
	}
	render(part, options) {
		this.imgManager = new ImgManager(this.zip, options.filePath, this.xmlDocuments);
		if (!part.type === "placeholder" || part.module !== moduleName) {
			return null;
		}
		try {
			const tagValue = options.scopeManager.getValue(part.value);
			if (!tagValue) {
				throw new Error("tagValue is empty");
			}
			const imgBuffer = this.options.getImage(tagValue, part.value);
			const rId = this.imgManager.addImageRels(this.getNextImageName(), imgBuffer);
			const sizePixel = this.options.getSize(imgBuffer, tagValue, part.value);
			return this.getRenderedPart(part, rId, sizePixel);
		}
		catch (e) {
			return {value: this.fileTypeConfig.tagTextXml};
		}
	}
	getRenderedPart(part, rId, sizePixel) {
		const size = [this.convertPixelsToEmus(sizePixel[0]), this.convertPixelsToEmus(sizePixel[1])];
		const centered = (this.options.centered || part.centered);
		let newText;
		if (this.options.fileType === "pptx") {
			newText = this.getPptRender(part, rId, size, centered);
		}
		else {
			newText = this.getDocxRender(part, rId, size, centered);
		}
		return {value: newText};
	}
	getPptRender(part, rId, size, centered) {
		const offset = {x: parseInt(part.off.x, 10), y: parseInt(part.off.y, 10)};
		const cellCX = parseInt(part.ext.cx, 10) || 1;
		const cellCY = parseInt(part.ext.cy, 10) || 1;
		const imgW = parseInt(size[0], 10) || 1;
		const imgH = parseInt(size[1], 10) || 1;

		if (centered) {
			offset.x = offset.x + (cellCX / 2) - (imgW / 2);
			offset.y = offset.y + (cellCY / 2) - (imgH / 2);
		}

		return this.getPptImageXml(rId, [imgW, imgH], offset);
	}
	getDocxRender(part, rId, size, centered) {
		return (centered) ? this.getImageXmlCentered(rId, size) : this.getImageXml(rId, size);
	}
	getNextImageName() {
		const name = `image_generated_${this.imageNumber}.png`;
		this.imageNumber++;
		return name;
	}
	convertPixelsToEmus(pixel) {
		return Math.round(pixel * 9525);
	}
	getImageXml(rId, size) {
		if (isNaN(rId)) {
			throw new Error("rId is NaN, aborting");
		}
		return `<w:drawing>
		<wp:inline distT="0" distB="0" distL="0" distR="0">
			<wp:extent cx="${size[0]}" cy="${size[1]}"/>
			<wp:effectExtent l="0" t="0" r="0" b="0"/>
			<wp:docPr id="2" name="Image 2" descr="image"/>
			<wp:cNvGraphicFramePr>
				<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
			</wp:cNvGraphicFramePr>
			<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
				<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
						<pic:nvPicPr>
							<pic:cNvPr id="0" name="Picture 1" descr="image"/>
							<pic:cNvPicPr>
								<a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
							</pic:cNvPicPr>
						</pic:nvPicPr>
						<pic:blipFill>
							<a:blip r:embed="rId${rId}">
								<a:extLst>
									<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
										<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
									</a:ext>
								</a:extLst>
							</a:blip>
							<a:srcRect/>
							<a:stretch>
								<a:fillRect/>
							</a:stretch>
						</pic:blipFill>
						<pic:spPr bwMode="auto">
							<a:xfrm>
								<a:off x="0" y="0"/>
								<a:ext cx="${size[0]}" cy="${size[1]}"/>
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst/>
							</a:prstGeom>
							<a:noFill/>
							<a:ln>
								<a:noFill/>
							</a:ln>
						</pic:spPr>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
		</wp:inline>
	</w:drawing>
		`.replace(/\t|\n/g, "");
	}
	getImageXmlCentered(rId, size) {
		if (isNaN(rId)) {
			throw new Error("rId is NaN, aborting");
		}
		return `<w:p>
			<w:pPr>
				<w:jc w:val="center"/>
			</w:pPr>
			<w:r>
				<w:rPr/>
				<w:drawing>
					<wp:inline distT="0" distB="0" distL="0" distR="0">
					<wp:extent cx="${size[0]}" cy="${size[1]}"/>
					<wp:docPr id="0" name="Picture" descr=""/>
					<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
						<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
						<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
							<pic:nvPicPr>
							<pic:cNvPr id="0" name="Picture" descr=""/>
							<pic:cNvPicPr>
								<a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
							</pic:cNvPicPr>
							</pic:nvPicPr>
							<pic:blipFill>
							<a:blip r:embed="rId${rId}"/>
							<a:stretch>
								<a:fillRect/>
							</a:stretch>
							</pic:blipFill>
							<pic:spPr bwMode="auto">
							<a:xfrm>
								<a:off x="0" y="0"/>
								<a:ext cx="${size[0]}" cy="${size[1]}"/>
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst/>
							</a:prstGeom>
							<a:noFill/>
							<a:ln w="9525">
								<a:noFill/>
								<a:miter lim="800000"/>
								<a:headEnd/>
								<a:tailEnd/>
							</a:ln>
							</pic:spPr>
						</pic:pic>
						</a:graphicData>
					</a:graphic>
					</wp:inline>
				</w:drawing>
			</w:r>
		</w:p>
		`.replace(/\t|\n/g, "");
	}
	getPptImageXml(rId, size, off) {
		if (isNaN(rId)) {
			throw new Error("rId is NaN, aborting");
		}
		return `<p:pic>
			<p:nvPicPr>
				<p:cNvPr id="6" name="Picture 2"/>
				<p:cNvPicPr>
					<a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
				</p:cNvPicPr>
				<p:nvPr/>
			</p:nvPicPr>
			<p:blipFill>
				<a:blip r:embed="rId${rId}" cstate="print">
					<a:extLst>
						<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
							<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
						</a:ext>
					</a:extLst>
				</a:blip>
				<a:srcRect/>
				<a:stretch>
					<a:fillRect/>
				</a:stretch>
			</p:blipFill>
			<p:spPr bwMode="auto">
				<a:xfrm>
					<a:off x="${off.x}" y="${off.y}"/>
					<a:ext cx="${size[0]}" cy="${size[1]}"/>
				</a:xfrm>
				<a:prstGeom prst="rect">
					<a:avLst/>
				</a:prstGeom>
				<a:noFill/>
				<a:ln>
					<a:noFill/>
				</a:ln>
				<a:effectLst/>
				<a:extLst>
					<a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}">
						<a14:hiddenFill xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
							<a:solidFill>
								<a:schemeClr val="accent1"/>
							</a:solidFill>
						</a14:hiddenFill>
					</a:ext>
					<a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}">
						<a14:hiddenLine xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" w="9525">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:miter lim="800000"/>
							<a:headEnd/>
							<a:tailEnd/>
						</a14:hiddenLine>
					</a:ext>
					<a:ext uri="{AF507438-7753-43E0-B8FC-AC1667EBCBE1}">
						<a14:hiddenEffects xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
							<a:effectLst>
								<a:outerShdw dist="35921" dir="2700000" algn="ctr" rotWithShape="0">
									<a:schemeClr val="bg2"/>
								</a:outerShdw>
							</a:effectLst>
						</a14:hiddenEffects>
					</a:ext>
				</a:extLst>
			</p:spPr>
		</p:pic>
		`.replace(/\t|\n/g, "");
	}
}

module.exports = ImageModule;
