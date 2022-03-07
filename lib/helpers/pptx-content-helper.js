/* eslint-disable no-prototype-builtins */

class PptxContentHelper {
    // Given the "content" block from the root (ex: PowerPointFactory.content), this function will pull out every slide and return very basic info on them.
    // (Right now, it's just the slide layout name used on each slide and the relId for that layout-to-slide relationship.)
    static extractInfoFromSlides(content) {
        let slideInformation = {}; // index is slide name

        for (let key in content) {
            if (key.substr(0, 16) === 'ppt/slides/slide') {
                let slideName = key.substr(11, key.lastIndexOf('.') - 11);
                let slideRelsKey = `ppt/slides/_rels/${slideName}.xml.rels`;
                let slideLayoutRelsNode = content[slideRelsKey].filter(function(element) {
                    return element.name() === "Relationship"
                        && element.attr('Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
                })[0];

                let relId = slideLayoutRelsNode.attr('Id');
                let target = slideLayoutRelsNode.attr('Target');
                let layout = target.substr(target.lastIndexOf('/') + 1);
                layout = layout.substr(0, layout.indexOf('.'));

                let objectInfo = PptxContentHelper.extractSlideObjectInfo(content, slideName);

                slideInformation[slideName] = { layout: { relId: relId, name: layout }, objectCount: objectInfo.objectCount };
            }
        }

        return slideInformation;
    }

    static extractSlideObjectInfo(content, slideName) {
        let objectInfo = {
            objectCount: 0,
        };

        // TODO... Mark: you can add code here...

        return objectInfo;
    }

    static extractNodes(contentBlock) {
        const existingNodes = [];

        for (let i = 0, count = contentBlock.childCount(); i < count; i++) {
            existingNodes.push(contentBlock.removeAt(0)[0]);
        }

        return existingNodes;
    }

    static restoreNodes(contentBlock, nodes) {
        nodes.forEach((node) => contentBlock.push(node));
    }
}

module.exports.PptxContentHelper = PptxContentHelper;
