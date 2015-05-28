SubContent=require('docxtemplater').SubContent
ImgManager=require('./imgManager')
ImgReplacer=require('./imgReplacer')
fs=require('fs')

jsxml = require("node-jsxml");
Namespace = jsxml.Namespace
QName = jsxml.QName
XML = jsxml.XML
XMLList = jsxml.XMLList

class ImageModule
	constructor:(@options={})->
		if !@options.centered? then @options.centered=false
		@qrQueue=[]
		@imageNumber=1
		@fileType="presentation" #TODO: handle this properly
	handleEvent:(event,eventData)->
		if event=='rendering-file'
			@renderingFileName=eventData
			gen=@manager.getInstance('gen')
			@imgManager=new ImgManager(gen.zip,@renderingFileName)
			@imgManager.loadImageRels()
		if event=='rendered'
			if @qrQueue.length==0 then @finished()
	get:(data)->
		if data=='loopType'
			templaterState=@manager.getInstance('templaterState')
			if templaterState.textInsideTag[0]=='%'
				return 'image'
		null
	getNextImageName:()->
		name="image_generated_#{@imageNumber}.png"
		@imageNumber++
		name
	replaceBy:(text,outsideElement)->
		xmlTemplater=@manager.getInstance('xmlTemplater')
		templaterState=@manager.getInstance('templaterState')
		subContent=new SubContent(xmlTemplater.content)
			.getInnerTag(templaterState)
			.getOuterXml(outsideElement)
		xmlTemplater.replaceXml(subContent,text)
	getPositionFromText:()->
		xmlTemplater=@manager.getInstance('xmlTemplater')
		templaterState=@manager.getInstance('templaterState')
		subContent=new SubContent(xmlTemplater.content)
			.getInnerTag(templaterState)
			.getOuterXml('p:sp').text
		return_value={}
		xml = new XML(subContent)
		xform = xml.child('spPr').child('xfrm')
		return_value.txtX = xform.child('off').attribute('x').getValue()
		return_value.txtY = xform.child('off').attribute('y').getValue()
		return_value.txtW = xform.child('ext').attribute('cx').getValue()
		return_value.txtH = xform.child('ext').attribute('cy').getValue()
		return_value

	convertPixelsToEmus:(pixel)->
		Math.round(pixel * 9525)
	getSizeFromData:(imgData)->
		[150,150]
	getImageFromData:(imgData)->
		fs.readFileSync(imgData)
	replaceTag:->
		scopeManager=@manager.getInstance('scopeManager')
		templaterState=@manager.getInstance('templaterState')

		tag = templaterState.textInsideTag.substr(1)
		imgData=scopeManager.getValueFromScope(tag)

		tagXml=@manager.getInstance('xmlTemplater').tagXml
		startEnd= "<#{tagXml}></#{tagXml}>"
		if imgData=='undefined' then return @replaceBy(startEnd,tagXml)
		try
			imgBuffer=@getImageFromData(imgData)
		catch e
			return @replaceBy(startEnd,tagXml)
		imageRels=@imgManager.loadImageRels();
		console.log('ImageRels', imageRels)
		if imageRels
			rId=imageRels.addImageRels(@getNextImageName(),imgBuffer)

			sizePixel=@getSizeFromData(imgBuffer)
			size=[@convertPixelsToEmus(sizePixel[0]),@convertPixelsToEmus(sizePixel[1])]

			if @options.centered==false
				outsideElement=tagXml
				newText=@getImageXml(rId,size)
			if @options.centered==true
				outsideElement=tagXml.substr(0,1)+':p'
				newText=@getImageXmlCentered(rId,size)
			if @fileType == 'presentation'
				console.log("File is Presentation")
				positionOfTextBox = @getPositionFromText()
				newText=@getPresentationImageXml(rId, positionOfTextBox.txtX, positionOfTextBox.txtY, positionOfTextBox.txtW, positionOfTextBox.txtH)
				outsideElement = 'p:sp'
			console.log('About to replace BY', newText)
			@replaceBy(newText,outsideElement)
	replaceQr:->
		xmlTemplater=@manager.getInstance('xmlTemplater')
		imR=new ImgReplacer(xmlTemplater,@imgManager)
		imR.getDataFromString=(result,cb)=>
			if @getImageFromDataAsync?
				@getImageFromDataAsync(result,cb)
			else
				cb(null,@getImageFromData(result))
		imR.pushQrQueue=(num)=>
			@qrQueue.push(num)
		imR.popQrQueue=(num)=>
			found = @qrQueue.indexOf(num)
			if found!=-1
				@qrQueue.splice(found,1)
			else @on('error',new Error("qrqueue #{num} is not in qrqueue"))
			if @qrQueue.length==0 then @finished()
		try
			imR
				.findImages()
				.replaceImages()
		catch e
			@on('error',e)
	finished:->
	on:(event,data)->
		if event=='error'
			throw data
	handle:(type,data)->
		if type=='replaceTag' and data=='image'
			@replaceTag()
		if type=='xmlRendered' and @options.qrCode
			@replaceQr()
		null
	getPresentationImageXml:(rId, x,y,cx,cy)->
		return """
		<p:pic><p:nvPicPr><p:cNvPr id="6" name="Picture 2"/><p:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
		<p:blipFill><a:blip r:embed="rId#{rId}" cstate="print">
		<a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
		<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip>
		<a:srcRect/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr bwMode="auto">
		<a:xfrm><a:off x="#{x}" y="#{y}"/><a:ext cx="#{cx}" cy="#{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:extLst><a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}"><a14:hiddenFill xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a14:hiddenFill></a:ext><a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}"><a14:hiddenLine xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" w="9525"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a14:hiddenLine></a:ext><a:ext uri="{AF507438-7753-43E0-B8FC-AC1667EBCBE1}"><a14:hiddenEffects xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a:effectLst><a:outerShdw dist="35921" dir="2700000" algn="ctr" rotWithShape="0"><a:schemeClr val="bg2"/></a:outerShdw></a:effectLst></a14:hiddenEffects></a:ext></a:extLst></p:spPr></p:pic>
		"""

	getImageXml:(rId,size)->
		return """
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="#{size[0]}" cy="#{size[1]}"/>
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
                    <a:blip r:embed="rId#{rId}">
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
                      <a:ext cx="#{size[0]}" cy="#{size[1]}"/>
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
		"""
	getImageXmlCentered:(rId,size)->
		"""
		<w:p>
		  <w:pPr>
			<w:jc w:val="center"/>
		  </w:pPr>
		  <w:r>
			<w:rPr/>
			<w:drawing>
			  <wp:inline distT="0" distB="0" distL="0" distR="0">
				<wp:extent cx="#{size[0]}" cy="#{size[1]}"/>
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
						<a:blip r:embed="rId#{rId}"/>
						<a:stretch>
						  <a:fillRect/>
						</a:stretch>
					  </pic:blipFill>
					  <pic:spPr bwMode="auto">
						<a:xfrm>
						  <a:off x="0" y="0"/>
						  <a:ext cx="#{size[0]}" cy="#{size[1]}"/>
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
		"""

module.exports=ImageModule
