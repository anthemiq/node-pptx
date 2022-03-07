const fs = require('fs');
const { Xml } = require('../../xmlnode');
const uuidv4 = require('uuid/v4');

class PptRelsFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/_rels/presentation.xml.rels`);
        this.content['ppt/_rels/presentation.xml.rels'] = Xml.parse(xml);
    }

    addPresentationToSlideRelationship(slideName) {
        const rId = `rId-${uuidv4()}`;

        this.content['ppt/_rels/presentation.xml.rels'].push(
            Xml.create('Relationship', null, {
                Id: rId,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                Target: `slides/${slideName}.xml`,
            })
        );

        return rId;
    }

    removePresentationToSlideRelationship(slideName) {
        let rId = -1;
        let relationshipIndex = -1;
        let slideType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
        let target = `slides/${slideName}.xml`;

        this.content['ppt/_rels/presentation.xml.rels'].forEach(function(element, index) {
            if (element.name() === 'Relationship' && element.attr('Type') === slideType && element.attr('Target') === target) {
                rId = element.attr('Id');
                relationshipIndex = index;
                return;
            }
        });

        if (relationshipIndex !== -1) {
            this.content['ppt/_rels/presentation.xml.rels'].removeAt(relationshipIndex);
        }

        return rId;
    }
}

module.exports.PptRelsFactory = PptRelsFactory;
