/* eslint-disable no-prototype-builtins */

class PptxContentHelper {
    // Given the "content" block from the root (ex: PowerPointFactory.content), this function will pull out every slide and return very basic info on them.
    // (Right now, it's just the slide layout name used on each slide and the relId for that layout-to-slide relationship.)
    static extractInfoFromSlides(content) {
        let slideInformation = {}; // index is slide name

        for (let key in content) {
            if (key.substr(0, 16) === 'ppt/slides/slide') {
                let slideName = key.substr(11, key.lastIndexOf('.') - 11);
                let slideKey = `ppt/slides/${slideName}.xml`;
                let slideRelsKey = `ppt/slides/_rels/${slideName}.xml.rels`;
                let slideLayoutRelsNode = content[slideRelsKey].filter(function(element) {
                    return element.name() === "Relationship"
                        && element.attr('Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
                })[0];

                let relId = slideLayoutRelsNode.attr('Id');
                let target = slideLayoutRelsNode.attr('Target');
                let layout = target.substr(target.lastIndexOf('/') + 1);
                layout = layout.substr(0, layout.indexOf('.'));

                let objectInfo = PptxContentHelper.extractSlideObjectInfo(content[slideKey]);

                slideInformation[slideName] = {
                    layout: { relId: relId, name: layout },
                    objectCount: objectInfo.objectCount,
                    highestObjectId: objectInfo.highestObjectId,
                };
            }
        }

        return slideInformation;
    }

    static extractSlideObjectInfo(content) {
        const objectInfo = {
            objectCount: 0,
            highestObjectId: 0,
        };

        const shapeTree = content.get('p:cSld/p:spTree');
        shapeTree.forEach((shapeNode) => {
            switch (shapeNode.name()) {
            case 'p:nvGrpSpPr':
                objectInfo.highestObjectId = Math.max(objectInfo.highestObjectId, Number(shapeNode.get('p:cNvPr')?.attr('id')));
                objectInfo.objectCount++;
                break;
            case 'p:sp':
                objectInfo.highestObjectId = Math.max(objectInfo.highestObjectId, Number(shapeNode.get('p:nvSpPr/p:cNvPr')?.attr('id')));
                objectInfo.objectCount++;
                break;
            case 'p:pic':
                objectInfo.highestObjectId = Math.max(objectInfo.highestObjectId, Number(shapeNode.get('p:nvPicPr/p:cNvPr')?.attr('id')));
                objectInfo.objectCount++;
                break;
            case 'p:graphicFrame':
                objectInfo.highestObjectId = Math.max(objectInfo.highestObjectId, Number(shapeNode.get('p:nvGraphicFramePr/p:cNvPr')?.attr('id')));
                objectInfo.objectCount++;
                break;
            }
        });

        return objectInfo;
    }

    static extractNodes(contentBlock) {
        const existingNodes = [];

        for (let i = 0, count = contentBlock.childCount(); i < count; i++) {
            existingNodes.push(contentBlock.removeAt(0));
        }

        return existingNodes;
    }

    static restoreNodes(contentBlock, nodes) {
        nodes.forEach((node) => contentBlock.push(node));
    }
}

module.exports.PptxContentHelper = PptxContentHelper;
