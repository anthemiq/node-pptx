/* eslint-disable no-prototype-builtins */

let { PptFactoryHelper } = require('../../helpers/ppt-factory-helper');
let { PptxContentHelper } = require('../../helpers/pptx-content-helper');
const { Xml } = require('../../xmlnode');

class SlideFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    addSlide(slideName, layoutName) {
        const relsKey = `ppt/slides/_rels/${slideName}.xml.rels`;
        const slideKey = `ppt/slides/${slideName}.xml`;
        const layoutKey = `ppt/slideLayouts/${layoutName}.xml`;

        this.content[relsKey] = Xml.create('Relationships', [
            Xml.create('Relationship', null, {
                Id: 'rId1',
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
                Target: `../slideLayouts/${layoutName}.xml`,}),
        ], {
            xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
        });

        // add the actual slide itself (use the layout template as the source; note: layout templates are NOT the same as master slide templates)
        const baseSlideContent = this.content[layoutKey].clone();
        baseSlideContent.setName('p:sld');
        baseSlideContent.removeAttr('preserve');
        baseSlideContent.removeAttr('type');
        return this.content[slideKey] = baseSlideContent;
    }

    removeSlide(slideName) {
        delete this.content[`ppt/slides/_rels/${slideName}.xml.rels`];
        delete this.content[`ppt/slides/${slideName}.xml`];
    }

    moveSlide(sourceSlideNum, destinationSlideNum) {
        if (destinationSlideNum > sourceSlideNum) {
            // move slides between start and destination backwards (e.g. slide 4 becomes 3, 3 becomes 2, etc.)
            for (let i = sourceSlideNum; i < destinationSlideNum; i++) {
                this.swapSlide(i, i + 1);
            }
        } else if (destinationSlideNum < sourceSlideNum) {
            // move slides between start and destination forward (e.g. slide 2 becomes 3, 3 becomes 4, etc.)
            for (let i = sourceSlideNum - 1; i >= destinationSlideNum; i--) {
                this.swapSlide(i, i + 1);
            }
        }
    }

    swapSlide(slideNum1, slideNum2) {
        let slideKey1 = `ppt/slides/slide${slideNum1}.xml`;
        let slideKey2 = `ppt/slides/slide${slideNum2}.xml`;
        let slideRelsKey1 = `ppt/slides/_rels/slide${slideNum1}.xml.rels`; // you need to swap rels in case slide layouts are used
        let slideRelsKey2 = `ppt/slides/_rels/slide${slideNum2}.xml.rels`;

        [this.content[slideKey1], this.content[slideKey2]] = [this.content[slideKey2], this.content[slideKey1]];
        [this.content[slideRelsKey1], this.content[slideRelsKey2]] = [this.content[slideRelsKey2], this.content[slideRelsKey1]];
    }

    addImageToSlideRelationship(slide, target) {
        let relsKey = `ppt/slides/_rels/${slide.name}.xml.rels`;
        let rId = `rId${this.content[relsKey].childCount() + 1}`;

        this.content[relsKey].push(
            Xml.create('Relationship', null, {
                Id: rId,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                Target: target,
            })
        );

        return rId;
    }

    addHyperlinkToSlideRelationship(slide, target) {
        if (!target) return '';

        let relsKey = `ppt/slides/_rels/${slide.name}.xml.rels`;
        let rId = `rId${this.content[relsKey].childCount() + 1}`;

         this.content[relsKey].push(
            Xml.create('Relationship', null, {
                Id: rId,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                Target: target,
                TargetMode: 'External',
            })
        );

        return rId;
    }

    addSlideTargetRelationship(slide, target) {
        if (!target) return '';

        let relsKey = `ppt/slides/_rels/${slide.name}.xml.rels`;
        let rId = `rId${this.content[relsKey].childCount() + 1}`;

        this.content[relsKey].push(
            Xml.create('Relationship', null, {
                Id: rId,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                Target: target,
            })
        );

        return rId;
    }

    addChartToSlideRelationship(slide, chartName) {
        let relsKey = `ppt/slides/_rels/${slide.name}.xml.rels`;
        let rId = `rId${this.content[relsKey].childCount() + 1}`;

        this.content[relsKey].push(
            Xml.create('Relationship', null, {
                Id: rId,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
                Target: `../charts/${chartName}.xml`,
            })
        );

        return rId;
    }

    addImage(slide, image, imageObjectName, rId) {
        let slideKey = `ppt/slides/${slide.name}.xml`;
        let objectCount = 0;

        //-----------------------------------------------------------------------------------------------------------------------------
        // TODO: Mark - something similar to this needs to be done in Factories/index.js -> PptxContentHelper.extractSlideObjectInfo():
        const spTreeRoot = this.content[slideKey].get('p:cSld/p:spTree');

        spTreeRoot.get('p:nvGrpSpPr').forEach(function(elem) {
            if (elem.name() === 'p:cNvPr') {
                objectCount++;
            }
        });

        // won't have sp nodes on a blank slide
        if (spTreeRoot.get('p:sp')) {
            spTreeRoot.forEach(function(elem) {
                if (elem.name() === 'p:sp' && elem.get('p:nvSpPr/p:cNvPr')) {
                    objectCount++;
                } else if (elem.name() === 'p:pic') {
                    objectCount++;
                }
            });
        }
        //-----------------------------------------------------------------------------------------------------------------------------

        // TODO: once the object count extractor is done, use _this_ ID in the "p:cNvPr" node below... (instead of "objectCount+1")
        let picObjectId = slide.getNextObjectId();

        let newImageBlock = Xml.create('p:pic', Xml.createTree({
            'p:nvPicPr': {
                'p:cNvPr': { [Xml.ATTR_KEY]: { id: objectCount + 1, name: `${imageObjectName} ${objectCount + 1}`, descr: imageObjectName } },
                'p:cNvPicPr': { 'a:picLocks': { [Xml.ATTR_KEY]: { noChangeAspect: '1' } } },
                'p:nvPr': null,
            },
            'p:blipFill': {
                'a:blip': { [Xml.ATTR_KEY]: { 'r:embed': rId, cstate: 'print' } },
                'a:stretch': { 'a:fillRect': null },
            },
            'p:spPr':  {
                'a:xfrm': {
                    'a:off': { [Xml.ATTR_KEY]: { x: image.x(), y: image.y() } },
                    'a:ext': { [Xml.ATTR_KEY]: { cx: image.cx(), cy: image.cy() } },
                },
                'a:prstGeom': { [Xml.ATTR_KEY]: { prst: 'rect' }, 'a:avLst': null, },
            },
        }));

        if (typeof image.options.url === 'string' && image.options.url.length > 0) {
            newImageBlock.get('p:nvPicPr/p:cNvPr').push(Xml.create('a:hlinkClick', null, { 'r:id': image.options.rIdForHyperlink }));

            if (image.options.url[0] === '#') {
                newImageBlock.get('p:nvPicPr/p:cNvPr/a:hlinkClick').setAttr('action', 'ppaction://hlinksldjump');
            }
        }

        spTreeRoot.push(newImageBlock);

        return newImageBlock;
    }

    addText(slide, textBox) {
        let slideKey = `ppt/slides/${slide.name}.xml`;
        let objectId = slide.getNextObjectId();
        let options = textBox.options;

        // construct the bare minimum structure of a shape block (text objects are a special case of shape)
        const newTextBlock = Xml.create('p:sp',
            PptFactoryHelper.createBaseShapeBlock(objectId, 'Text', textBox.x(), textBox.y(), textBox.cx(), textBox.cy())
        );

        // now add the nodes which turn a shape block into a text block
        newTextBlock.get('p:nvSpPr/p:cNvSpPr').setAttr('txBox', '1');
        newTextBlock.get('p:txBody/a:bodyPr').setAttr('rtlCol', '0');

        if (options.backgroundColor) {
            newTextBlock.get('p:spPr').push(Xml.create('a:solidFill', [PptFactoryHelper.createColorBlock(options.backgroundColor)]));
        } else {
            newTextBlock.get('p:spPr').push(Xml.create('a:noFill'));
        }

        PptFactoryHelper.addTextValuesToBlock(newTextBlock.get('p:txBody'), textBox, options);
        PptFactoryHelper.setTextBodyProperties(newTextBlock.get('p:txBody/a:bodyPr'), textBox, options);
        PptFactoryHelper.setShapeProperties(newTextBlock.get('p:spPr'), options);

        if (typeof options.url === 'string' && options.url.length > 0) {
            if (options.applyHrefOnShapeOnly) {
                newTextBlock.get('p:nvSpPr/p:cNvPr').push(Xml.create('a:hlinkClick', null, { 'r:id': options.rIdForHyperlink }));

                if (options.url[0] === '#') {
                    newTextBlock.get('p:nvSpPr/p:cNvPr/a:hlinkClick').setAttr('action', 'ppaction://hlinksldjump');
                }
            } else {
                newTextBlock.get('p:txBody/a:p/a:r/a:rPr').push(Xml.create('a:hlinkClick', null, { 'r:id': options.rIdForHyperlink }));

                if (options.url[0] === '#') {
                    newTextBlock.get('p:txBody/a:p/a:r/a:rPr/a:hlinkClick').setAttr('action', 'ppaction://hlinksldjump');
                }
            }
        }

        this.content[slideKey].get('p:cSld/p:spTree').push(newTextBlock);

        return newTextBlock;
    }

    addShape(slide, shape) {
        let slideKey = `ppt/slides/${slide.name}.xml`;
        let objectId = slide.getNextObjectId();
        let options = shape.options;
        let type = shape.shapeType;
        let shapeColor = options.color || '00AA00';

        if (options.textAlign === undefined) {
            options.textAlign = 'center'; // for shapes, we always want text defaulted to the center
        }

        const newShapeBlock = Xml.create('p:sp',
            PptFactoryHelper.createBaseShapeBlock(objectId, 'Shape', shape.x(), shape.y(), shape.cx(), shape.cy())
        );

        newShapeBlock.get('p:spPr/a:prstGeom').setAttr('prst', type.name);
        newShapeBlock.get('p:spPr').push(Xml.create('a:solidFill', [PptFactoryHelper.createColorBlock(shapeColor)]));
        newShapeBlock.get('p:txBody/a:bodyPr').setAttr('rtlCol', '0');

        PptFactoryHelper.addTextValuesToBlock(newShapeBlock.get('p:txBody'), shape, options);
        PptFactoryHelper.setTextBodyProperties(newShapeBlock.get('p:txBody/a:bodyPr'), shape, options);
        PptFactoryHelper.setShapeProperties(newShapeBlock.get('p:spPr'), options, type.avLst);

        if (typeof options.url === 'string' && options.url.length > 0) {
            newShapeBlock.get('p:nvSpPr/p:cNvPr').push(Xml.create('a:hlinkClick', null, { 'r:id': options.rIdForHyperlink }));

            if (options.url[0] === '#') {
                newShapeBlock.get('p:nvSpPr/p:cNvPr/a:hlinkClick').setAttr('action', 'ppaction://hlinksldjump');
            }
        }

        this.content[slideKey].get('p:cSld/p:spTree').push(newShapeBlock);

        return newShapeBlock;
    }

    addChart(slide, chart) {
        let slideKey = `ppt/slides/${slide.name}.xml`;
        let chartKey = `ppt/charts/${chart.name}.xml`;

        let newGraphicFrameBlock = PptFactoryHelper.createBaseChartFrameBlock(chart.x(), chart.y(), chart.cx(), chart.cy()); // goes onto the slide
        let newChartSpaceBlock = PptFactoryHelper.createBaseChartSpaceBlock(); // goes into the chart XML
        let seriesDataBlock = PptFactoryHelper.createSeriesDataBlock(chart.chartData);

        newChartSpaceBlock.get('c:chart/c:plotArea/c:barChart/c:ser').pushAll(seriesDataBlock);

        this.content[chartKey] = newChartSpaceBlock;
        this.content[slideKey].get('p:cSld/p:spTree').push(newGraphicFrameBlock);

        return newGraphicFrameBlock;
    }

    setBackgroundColor(slide, color) {
        let slideKey = `ppt/slides/${slide.name}.xml`;
        let slideContent = this.content[slideKey].get('p:cSld');

        if (!slideContent.get('p:bg')) {
            slideContent.insert(0, Xml.create('p:bg'));
        }

        slideContent.get('p:bg').remove('p:bgPr');
        slideContent.get('p:bg').push(Xml.create('p:bgPr', [
            Xml.create('a:solidFill', [PptFactoryHelper.createColorBlock(color)]),
            Xml.create('a:effectLst'),
        ]));
    }
}

module.exports.SlideFactory = SlideFactory;
