const fs = require('fs');
const { Xml } = require('../../xmlnode');

let { PptxContentHelper } = require('../../helpers/pptx-content-helper');
let { PptxUnitHelper } = require('../../helpers/unit-helper');

class PresentationFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/presentation.xml`);
        this.content['ppt/presentation.xml'] = Xml.parse(xml);
    }

    addSlideRefIdToGlobalList(rId) {
        let maxId = 255; // slide ID #'s start at some magic number of 256 (so init to 255, then the first slide will have 256)
        let presentationContent = this.content['ppt/presentation.xml'];

        if (!presentationContent.get('p:sldIdLst')) {
            // if we don't have a <p:sldIdLst> node yet (i.e. no slides), insert a new sldIdLst node under the sldMasterIdLst node
            let existingNodes = PptxContentHelper.extractNodes(presentationContent);
            let masterIdLstIndex = existingNodes.findIndex(node => node.name === 'p:sldMasterIdLst');

            presentationContent.push(existingNodes[masterIdLstIndex]);

            existingNodes.splice(masterIdLstIndex, 1); // delete <p:sldMasterIdLst> from the list so the call to restoreNodes() doesn't add it again

            presentationContent.push(Xml.create('p:sldIdLst'));

            PptxContentHelper.restoreNodes(presentationContent, existingNodes);
        } else {
            if (presentationContent.get('p:sldIdLst')) {
                presentationContent.get('p:sldIdLst').forEach(function(node) {
                    if (+node.attr('id') > maxId) maxId = +node.attr('id');
                });
            }
        }

        presentationContent.get('p:sldIdLst').push(
            Xml.create('p:sldId', null, {
                id: `${+maxId + 1}`,
                'r:id': rId,
            })
        );
    }

    removeSlideRefIdFromGlobalList(rId) {
        let slideIdListIndex = -1;
        let presentationContent = this.content['ppt/presentation.xml'];

        if (presentationContent.get('p:sldIdLst')) {
            presentationContent.get('p:sldIdLst').forEach(function(node, index) {
                if (node.name() === 'p:sldId' && node.attr('r:id') === rId) {
                    slideIdListIndex = index;
                }
            });
        }

        if (slideIdListIndex !== -1) {
            presentationContent.get('p:sldIdLst').removeAt(slideIdListIndex);

            if (presentationContent.get('p:sldIdLst').childCount() === 0) {
                presentationContent.remove('p:sldIdLst');
            }
        }
    }

    setLayout(layout) {
        const slideSizeBlock = this.content['ppt/presentation.xml'].get('p:sldSz');
        const originalCx = slideSizeBlock.attr('cx');
        const originalCy = slideSizeBlock.attr('cy');

        slideSizeBlock.setAttr('cx', layout.width);
        slideSizeBlock.setAttr('cy', layout.height);
        slideSizeBlock.setAttr('type', layout.type);

        // note: seems like there is no "type" attribute on the note sizes
        this.content['ppt/presentation.xml'].get('p:notesSz').setAttr('cx', layout.width);
        this.content['ppt/presentation.xml'].get('p:notesSz').setAttr('cy', layout.height);

        slideSizeBlock.removeAttr('type');

        if (originalCx !== layout.width || originalCy !== layout.height) {
            let slideNumberOffsetInches = 0.25; // number of inches of padding between slide number and right side of slide
            let size = this.getSlideNumberShapeSizeFromLayout('slideLayout1');

            // cx and cy props will be -1 if the slide number shape doesn't exist (it exists when making a pptx from
            // scratch using this library, but probably won't exist when loading an external PPTX - there's no way
            // for this library to detect whether a third-party PPTX has an auto slide number shape without knowing
            // the object name and the layout name in which it would reside)
            if (size.cx !== -1 && size.cy !== -1) {
                let newX = layout.width - size.cx - PptxUnitHelper.fromInches(slideNumberOffsetInches);
                let newY = layout.height - size.cy;

                this.moveSlideNumberOnLayoutTemplate('slideLayout1', newX, newY);
            }
        }
    }

    moveSlideNumberOnLayoutTemplate(layoutName, x, y) {
        let slideNumberNode = this.getSlideNumberShapeNodeFromLayout(layoutName);

        if (slideNumberNode) {
            slideNumberNode.get('p:spPr/a:xfrm/a:off').setAttr('x', x);
            slideNumberNode.get('p:spPr/a:xfrm/a:off').setAttr('y', y);
        }
    }

    getSlideNumberShapeSizeFromLayout(layoutName) {
        let slideNumberNode = this.getSlideNumberShapeNodeFromLayout(layoutName);

        if (slideNumberNode) {
            return {
                cx: slideNumberNode.get('p:spPr/a:xfrm/a:ext').attr('cx'),
                cy: slideNumberNode.get('p:spPr/a:xfrm/a:ext').attr('cy'),
            };
        }

        return { cx: -1, cy: -1 };
    }

    getSlideNumberShapeNodeFromLayout(layoutName) {
        let layoutKey = `ppt/slideLayouts/${layoutName}.xml`;

        if (this.content[layoutKey]) {
            let templateSlideLayoutContent = this.content[layoutKey];
            let shapesRoot = templateSlideLayoutContent.get('p:cSld/p:spTree');

            if (shapesRoot) {
                let slideNumberShapeIndex = -1;

                shapesRoot.forEach(function(node, index) {
                    if (node.name() === 'p:sp' && node.get('p:nvSpPr/p:cNvPr').attr('name') === 'Slide Number Placeholder 1') {
                        slideNumberShapeIndex = index;
                    }
                });

                if (slideNumberShapeIndex !== -1) return shapesRoot.at(slideNumberShapeIndex);
            }
        }
    }
}

module.exports.PresentationFactory = PresentationFactory;
