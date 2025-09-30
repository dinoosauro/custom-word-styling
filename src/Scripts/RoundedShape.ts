

interface Props {
    /**
     * The width of the rectangle, in points
     */
    width: number,
    /**
     * The height of the rectangle, in points
     */
    height: number,
    /**
     * The border radius, between 0 (square) and 0.5 (circle)
     */
    borderRadius: number
    /**
     * The color used in the plain color mode. You can change the transparency (or delete it) from the `plainColorSettings` object
     */
    color: string,
    /**
     * If passed, a gradient will be added as the background
     */
    gradient?: {
        /**
         * The angle of the gradient
         */
        angle: number,
        /**
         * A nested array, composed of [the position of the color in the gradient (from 0 to 1), the hex color, and the transparency of the color]
         */
        colors: [number, string, number][]
    },
    /**
     * If passed, a background image will be added in the shape
     */
    image?: {
        /**
         * The base64 of the image, either in PNG or in JPEG
         */
        base64: string,
        isPng: boolean,
        height: number,
        width: number,
        /**
         * How the rounded rectangle should be resized so that the image won't be stretched. 
         * - If `none`, the image will be stretched;
         * - If `keepWidth`, the width will remain the same, and the height will be edited;
         * - If `keepHeight`, the height will remain the same, and the widht will be edited.
         */
        scaleType: "none" | "keepWidth" | "keepHeight"
    },
    /**
     * Settings about the border of the rectangle. If you don't want to add a border, don't pass this value.
     */
    border?: {
        /**
         * If provided, a plain color will be added to the border
         */
        plainColor?: string,
        /**
         * Transparency of the plain color for the border
         */
        plainColorTransparency?: number
        /**
         * If passed, the border will be a linear gradient
         */
        gradient?: {
            /**
            * The angle of the gradient
            */
            angle: number,
            /**
             * A nested array, composed of [the position of the color in the gradient (from 0 to 1), the hex color, and the transparency of the color]
             */
            colors: [number, string, number][]
        },
        size: number,
        /**
         * Type of the border line (ex: single line, double line etc.)
         */
        borderType: string,
        /**
         * Style of the border line (ex: normal, dashed, dotted etc.)
         */
        borderLineStyle: string
    },
    /**
     * Custom settings for the plain color, passed in the `color` property of this object.
     */
    plainColorSettings?: {
        /**
         * Trasparency of the plain color
         */
        transparency?: number,
        /**
         * If the plain color shouldn't be added
         */
        deleteBackground?: boolean
    },
    /**
     * The type of the shape to create
     */
    shapeType: string
}
/**
 * Create a new shape, and add it in the document
 */
export default async function CreateRoundedShape({ color, gradient, image, width, height, borderRadius, border, plainColorSettings, shapeType }: Props) {
    // Placeholder for the shape style I want
    const XMLDoc = new DOMParser().parseFromString(`<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
    <pkg:part pkg:name="/_rels/.rels"
        pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
        <pkg:xmlData>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1"
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    Target="word/document.xml" />
            </Relationships>
        </pkg:xmlData>
    </pkg:part>
    <pkg:part pkg:name="/word/document.xml"
        pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
            <w:document
                xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
                xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
                xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
                xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
                xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
                xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
                xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
                xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
                xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
                xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
                xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:oel="http://schemas.microsoft.com/office/2019/extlst"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                xmlns:v="urn:schemas-microsoft-com:vml"
                xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:w10="urn:schemas-microsoft-com:office:word"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
                xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
                xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
                xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"
                xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
                xmlns:w16sdtfl="http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock"
                xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
                xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
                xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
                xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
                xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14">
                <w:body>
                    <w:p>
                        <w:r>
                            <w:rPr>
                                <w:noProof />
                            </w:rPr>
                            <mc:AlternateContent>
                                <mc:Choice Requires="wps">
                                    <w:drawing>
                                        <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
                                            simplePos="0" relativeHeight="251659264" behindDoc="0"
                                            locked="0" layoutInCell="1" allowOverlap="1">
                                            <wp:simplePos x="0" y="0" />
                                            <wp:positionH relativeFrom="column">
                                                <wp:posOffset>179680</wp:posOffset>
                                            </wp:positionH>
                                            <wp:positionV relativeFrom="paragraph">
                                                <wp:posOffset>-29286</wp:posOffset>
                                            </wp:positionV>
                                            <wp:extent cx="2648102" cy="1580083" />
                                            <wp:effectExtent l="0" t="0" r="6350" b="0" />
                                            <wp:wrapNone />
                                            <wp:docPr id="1739854477"
                                                name="Rettangolo con angoli arrotondati 1" />
                                            <wp:cNvGraphicFramePr />
                                            <a:graphic
                                                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                                <a:graphicData
                                                    uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                                    <wps:wsp>
                                                        <wps:cNvSpPr />
                                                        <wps:spPr>
                                                            <a:xfrm>
                                                                <a:off x="0" y="0" />
                                                                <a:ext cx="2648102" cy="1580083" />
                                                            </a:xfrm>
                                                            <a:prstGeom prst="roundRect">
                                                                <a:avLst>
                                                                    <a:gd name="adj"
                                                                        fmla="val 50000" />
                                                                </a:avLst>
                                                            </a:prstGeom>
                                                            <a:solidFill>
                                                                <a:srgbClr val="363F34" />
                                                            </a:solidFill>
                                                            <a:ln>
                                                                <a:noFill />
                                                            </a:ln>
                                                        </wps:spPr>
                                                        <wps:style>
                                                            <a:lnRef idx="2">
                                                                <a:schemeClr val="accent1">
                                                                    <a:shade val="15000" />
                                                                </a:schemeClr>
                                                            </a:lnRef>
                                                            <a:fillRef idx="1">
                                                                <a:schemeClr val="accent1" />
                                                            </a:fillRef>
                                                            <a:effectRef idx="0">
                                                                <a:schemeClr val="accent1" />
                                                            </a:effectRef>
                                                            <a:fontRef idx="minor">
                                                                <a:schemeClr val="lt1" />
                                                            </a:fontRef>
                                                        </wps:style>
                                                        <wps:bodyPr rot="0" spcFirstLastPara="0"
                                                            vertOverflow="overflow"
                                                            horzOverflow="overflow" vert="horz"
                                                            wrap="square" lIns="91440" tIns="45720"
                                                            rIns="91440" bIns="45720" numCol="1"
                                                            spcCol="0" rtlCol="0" fromWordArt="0"
                                                            anchor="ctr" anchorCtr="0" forceAA="0"
                                                            compatLnSpc="1">
                                                            <a:prstTxWarp prst="textNoShape">
                                                                <a:avLst />
                                                            </a:prstTxWarp>
                                                            <a:noAutofit />
                                                        </wps:bodyPr>
                                                    </wps:wsp>
                                                </a:graphicData>
                                            </a:graphic>
                                        </wp:anchor>
                                    </w:drawing>
                                </mc:Choice>
                                <mc:Fallback>
                                    <w:pict>
                                        <v:roundrect 
                                            style="position:absolute;margin-left:14.15pt;margin-top:-2.3pt;width:208.5pt;height:124.4pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;v-text-anchor:middle"
                                            arcsize="10923f"
                                            fillcolor="#363f34" stroked="f" strokeweight="1pt">
                                            <v:stroke joinstyle="miter" />
                                        </v:roundrect>
                                    </w:pict>
                                </mc:Fallback>
                            </mc:AlternateContent>
                        </w:r>
                    </w:p>
                    <w:p/>
                    <w:sectPr>
                        <w:pgSz w:w="11906" w:h="16838" />
                        <w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1134"
                            w:header="708" w:footer="708" w:gutter="0" />
                        <w:cols w:space="708" />
                        <w:docGrid w:linePitch="360" />
                    </w:sectPr>
                </w:body>
            </w:document>
        </pkg:xmlData>
    </pkg:part>
</pkg:package>`, "application/xml")
    await Word.run(async (ctx) => {
        /**
         * The main Rectangle object
         */
        const rect = XMLDoc.getElementsByTagName("v:roundrect")[0];
        if (typeof image !== "undefined") { // Apply the image as the background
            // We need to get the OOXML of the document for two reasons:
            // - We need to add an ID for the new image file
            // - For some reason, Word would reject the new shape if the stylesheet isn't appended at the end.
            const alreadyAddedOOxml = ctx.document.body.getOoxml();
            await ctx.sync();
            const parsedAddedOoxml = new DOMParser().parseFromString(alreadyAddedOOxml.value, "application/xml");
            /**
             * All the links to other XML/image files
             */
            const availableRelationships = Array.from(parsedAddedOoxml.getElementsByTagName("pkg:part")).find(i => i.getAttribute("pkg:name") === "/word/_rels/document.xml.rels")?.getElementsByTagName("Relationship");
            /**
             * Unique ID for the current file. Might be updated to the "rIdX" syntax
             */
            let uuid: string = crypto.randomUUID();
            /**
             * The number of the image, that'll be part of the image file name (ex: `image1.jpg`)
             */
            let imageNumber = +Math.random().toString().substring(2);
            if (availableRelationships) {
                const availableRelationshipsArr = Array.from(availableRelationships);
                /**
                 * All the rIds that have already been added in the document
                 */
                const entryNumber = availableRelationshipsArr.map(i => parseInt(i.getAttribute("Id")?.substring(3) ?? "0"));
                // We'll now look to get the first ID number that hasn't already been used
                let tempUuid = availableRelationshipsArr.length + 1;
                while (entryNumber.includes(tempUuid)) tempUuid++;
                // And now we'll do the same, but for images.
                const availableImages = availableRelationshipsArr.filter(i => i.getAttribute("Target")?.startsWith("media/image")).map(i => parseInt(i.getAttribute("Target")?.substring("media/image".length) ?? "0"));
                let tempImageNumber = availableImages.length;
                while (availableImages.includes(tempImageNumber)) tempImageNumber++;
                uuid = `rId${tempUuid}`;
                imageNumber = tempImageNumber;
            }
            // Now we'll create the Relationships XML node, where we'll add the reference to the new image file
            const relationshipPart = XMLDoc.createElement("pkg:part");
            for (const [key, value] of [
                ["pkg:name", "/word/_rels/document.xml.rels"],
                ["pkg:contentType", "application/vnd.openxmlformats-package.relationships+xml"],
                ["pkg:padding", "256"]
            ]) relationshipPart.setAttribute(key, value);
            const pkgData = XMLDoc.createElement("pkg:xmlData");
            const relationshipGroup = XMLDoc.createElement("Relationships");
            relationshipGroup.setAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
            const relationship = XMLDoc.createElement("Relationship");
            for (const [key, value] of [
                ["Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"],
                ["Id", uuid],
                ["Target", `media/image${imageNumber}.${image.isPng ? "png" : "jpg"}`]
            ]) relationship.setAttribute(key, value);
            relationshipGroup.append(relationship);
            pkgData.append(relationshipGroup);
            relationshipPart.append(pkgData);
            XMLDoc.getElementsByTagName("pkg:package")[0].append(relationshipPart);

            // And now we'll create a new XML part, where we'll add the base64 of the image
            const imagePart = XMLDoc.createElement("pkg:part");
            for (const [key, value] of [
                ["pkg:name", `/word/media/image${imageNumber}.${image.isPng ? "png" : "jpg"}`],
                ["pkg:contentType", `image/${image.isPng ? "png" : "jpeg"}`],
                ["pkg:compression", "store"]
            ]) imagePart.setAttribute(key, value);
            const binaryData = XMLDoc.createElement("pkg:binaryData");
            binaryData.textContent = image.base64;
            imagePart.append(binaryData);
            XMLDoc.getElementsByTagName("pkg:package")[0].append(imagePart);
            // And we'll also append the source theme, since otherwise the Word Office.JS library would throw an Exception
            const themes = Array.from(parsedAddedOoxml.getElementsByTagName("pkg:part")).find(i => i.getAttribute("pkg:name") === "/word/theme/theme1.xml" || i.getAttribute("pkg:name") === "/word/styles.xml");
            if (themes) XMLDoc.getElementsByTagName("pkg:package")[0].append(themes);
            // Now, we need to update the placeholder rectangle: we need to remove the reference to the plain color, and add the ones that indicate to Work that an image is being used in the background.
            const imageFill = XMLDoc.createElement("a:blipFill");
            const blip = XMLDoc.createElement("a:blip");
            blip.setAttribute("r:embed", uuid);
            const stretch = XMLDoc.createElement("a:stretch");
            const fillRect = XMLDoc.createElement("a:fillRect");
            stretch.append(fillRect);
            imageFill.append(blip, stretch);
            XMLDoc.getElementsByTagName("w:drawing")[0].getElementsByTagName("a:solidFill")[0].remove();
            XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("wps:spPr")[0].insertBefore(imageFill, XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("wps:spPr")[0].getElementsByTagName("a:ln")[0]);
            rect.removeAttribute("fillcolor");
            const fillInsideRect = XMLDoc.createElement("v:fill");
            for (const [key, value] of [
                ["r:id", uuid],
                ["o:title", ""],
                ["recolor", "t"],
                ["rotate", "t"],
                ["type", "frame"]
            ]) fillInsideRect.setAttribute(key, value);
            rect.prepend(fillInsideRect);
        } else if (typeof gradient !== "undefined") { // We'll apply a gradient as the background of an image
            // Let's start by sorting all the gradients according to their position
            gradient.colors.sort((a, b) => a[0] - b[0]);
            rect.setAttribute("fillcolor", `${gradient.colors[0][1]}`); // Update the fillcolor value with the first entry of the gradient
            // Let's create the fill node that'll be added in the rectangle
            const fillNode = XMLDoc.createElement("v:fill");
            for (const [key, value] of [
                ["color2", `${gradient.colors[gradient.colors.length - 1][1]}`],
                ["colors", gradient.colors.map(([percent, color]) => `${(percent === 0 ? "00" : percent === 1 ? "01" : percent).toString().substring(1)} ${color}`).join(";")],
                ["focus", "100%"],
                ["type", "gradient"]
            ]) fillNode.setAttribute(key, value);
            rect.prepend(fillNode);
            // And now we need to replace the reference to the plain color with the reference to the gradient. We need to specify again the gradient, and the position here is specified in tens of thousands apprently.
            XMLDoc.getElementsByTagName("w:drawing")[0].getElementsByTagName("a:solidFill")[0].remove();
            const gradientFill = getGradientFill(XMLDoc, gradient)
            XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("wps:spPr")[0].insertBefore(gradientFill, XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("wps:spPr")[0].getElementsByTagName("a:ln")[0]);
        } else if (plainColorSettings?.deleteBackground) { // Remove the solid color background, and add the noFill node
            XMLDoc.getElementsByTagName("w:drawing")[0].getElementsByTagName("a:solidFill")[0].remove();
            rect.removeAttribute("fillcolor");
            const noFill = XMLDoc.createElement("a:noFill");
            XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("wps:spPr")[0].insertBefore(noFill, XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("wps:spPr")[0].getElementsByTagName("a:ln")[0]);
        } else if (typeof color !== "undefined") { // Update the plain color
            rect.setAttribute("fillcolor", color);
            XMLDoc.getElementsByTagName("w:drawing")[0].getElementsByTagName("a:solidFill")[0].firstElementChild?.setAttribute("val", color.substring(1));
            if (typeof plainColorSettings?.transparency !== "undefined" && plainColorSettings.transparency !== 0) { // And update the transparency
                const alpha = XMLDoc.createElement("a:alpha");
                alpha.setAttribute("val", (100_000 - (plainColorSettings.transparency * 100_000)).toString());
                XMLDoc.getElementsByTagName("w:drawing")[0].getElementsByTagName("a:solidFill")[0].getElementsByTagName("a:srgbClr")[0].prepend(alpha);
            }
        }

        if (border?.gradient) { // Apply a linear gradient to the border
            const gradientFill = getGradientFill(XMLDoc, border.gradient);
            XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("a:ln")[0].getElementsByTagName("a:noFill")[0].remove();
            XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("a:ln")[0].prepend(gradientFill);
            rect.setAttribute("strokecolor", border.gradient.colors[0][1]);
        } else if (border?.plainColor) { // Apply a plain color to the border
            const solidFill = XMLDoc.createElement("a:solidFill");
            const srgbColor = XMLDoc.createElement("a:srgbClr");
            srgbColor.setAttribute("val", border.plainColor.substring(1));
            if (typeof border.plainColorTransparency !== "undefined" && border.plainColorTransparency !== 0) { // Add the alpha node with the transparency value.
                const alpha = XMLDoc.createElement("a:alpha");
                alpha.setAttribute("val", (100_000 - (border.plainColorTransparency * 100_000)).toString());
                srgbColor.prepend(alpha); 
            }
            solidFill.prepend(srgbColor);
            XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("a:ln")[0].getElementsByTagName("a:noFill")[0].remove();
            XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("a:ln")[0].prepend(solidFill);
            rect.setAttribute("strokecolor", border.plainColor);
        }
        if (border) { // Update border styling
            XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("a:ln")[0].setAttribute("cmpd", border.borderType); // Single/double line etc.
            if (border.borderLineStyle !== "normal") { // Update the line style (dashed, dotted etc.). We need to update both the `a:prstDash` and the `v:stroke` properties, that are divided in that property by the first space.
                const prstDash = XMLDoc.createElement("a:prstDash");
                prstDash.setAttribute("val", border.borderLineStyle.substring(0, border.borderLineStyle.indexOf(" ")));
                XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("a:ln")[0].prepend(prstDash);
                rect.getElementsByTagName("v:stroke")[0].setAttribute("dashstyle", border.borderLineStyle.substring(border.borderLineStyle.indexOf(" ") + 1));
            }
        }
        // Update the border radius of the rounded rectangle
        XMLDoc.getElementsByTagName("a:prstGeom")[0].getElementsByTagName("a:gd")[0].setAttribute("fmla", `val ${Math.floor(borderRadius * 100000)}`)
        rect.setAttribute("arcsize", borderRadius === 0 ? "0" : borderRadius.toString().substring(1));
        // And we'll update the border size
        if (typeof border?.size !== "undefined") XMLDoc.getElementsByTagName("a:graphicData")[0].getElementsByTagName("a:ln")[0].setAttribute("w", (+border.size * 12700).toString());
        /**
         * The XML to add in the Word document
         */
        let outputOoxml = new XMLSerializer().serializeToString(XMLDoc).replaceAll(`xmlns=""`, ""); // We need to replace empty xmlns since otherwise Word would throw an error
        if (typeof image !== "undefined" && image.scaleType !== "none") { // We need to update the height so that the image isn't stretched
            let strWidth = outputOoxml.substring(outputOoxml.indexOf(image.scaleType === "keepWidth" ? "width:" : "height:") + (image.scaleType === "keepWidth" ? 6 : 7));
            strWidth = strWidth.substring(0, strWidth.indexOf("pt"));
            const widthNumber = +strWidth.replace("pt", "");
            const outputHeight = ((widthNumber * (image.scaleType === "keepWidth" ? image.height : image.width)) / (image.scaleType === "keepWidth" ? image.width : image.height)).toFixed(2);
            outputOoxml = outputOoxml.replaceAll(`${image.scaleType === "keepWidth" ? "height" : "width"}:${image.scaleType === "keepWidth" ? "124.4" : "208.5"}pt`, `${image.scaleType === "keepWidth" ? "height" : "width"}:${outputHeight}pt`).replaceAll(`c${image.scaleType === "keepWidth" ? "y" : "x"}="${image.scaleType === "keepWidth" ? "1580083" : "2648102"}"`, `c${image.scaleType === "keepWidth" ? "y" : "x"}="${Math.round(+outputHeight * 12700)}"`);
        }
        // Now we apply the custom width/height set by the user
        if (image?.scaleType !== "keepWidth") outputOoxml = outputOoxml.replaceAll("height:124.4pt", `height:${height}pt`).replaceAll(`cy="1580083"`, `cy="${Math.round(+height * 12700)}"`);
        if (image?.scaleType !== "keepHeight") outputOoxml = outputOoxml.replaceAll("width:208.5pt", `width:${width}px`).replaceAll(`cx="2648102"`, `cx="${Math.round(+width * 12700)}"`);
        // And finally we add the temp rounded rectangle to the document, and we'll update its shape to the one asked by the user.
        const result = ctx.document.body.insertOoxml(outputOoxml, "End");
        await ctx.sync();
        const shapes = result.shapes.load();
        await ctx.sync();
        shapes.items[0].geometricShapeType = shapeType as "RoundRectangle";
        await ctx.sync();
    })
}

/**
 * Get the node that tells Word to make the current entry a linear gradient.
 * @param XMLDoc the XML Document that is being edited
 * @param gradient gradient information
 * @returns the `a:gradFill` node
 */
function getGradientFill(XMLDoc: Document, gradient: {angle: number, colors: [number, string, number][]}) {
    const gradientFill = XMLDoc.createElement("a:gradFill");
    const gradientList = XMLDoc.createElement("a:gsLst");
    for (const color of gradient.colors) {
        const entry = XMLDoc.createElement("a:gs");
        entry.setAttribute("pos", color[0] === 0 ? "0" : color[0] === 1 ? "100000" : color[0].toFixed(5).substring(2));
        const srgb = XMLDoc.createElement("a:srgbClr");
        srgb.setAttribute("val", color[1].substring(1).toUpperCase());
        if (color[2] !== 0) {
            const alpha = XMLDoc.createElement("a:alpha");
            alpha.setAttribute("val", (100_000 - (color[2] * 100_000)).toString());
            srgb.append(alpha);
        }
        entry.append(srgb);
        gradientList.append(entry);
    }
    const lin = XMLDoc.createElement("a:lin");
    for (const [key, value] of [
        ["ang", (gradient.angle * 10800000 / 180).toString()], // I don't know the real unit, I just observed the XML between 0 degrees and 90 degrees and I made the proportion
        ["scaled", "0"]
    ]) lin.setAttribute(key, value);
    gradientFill.append(gradientList);
    gradientFill.append(lin);
    return gradientFill;
}