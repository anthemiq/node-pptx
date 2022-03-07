const fs = require('fs');
const { Xml } = require('../../xmlnode');

class SlideLayoutsFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        let xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/slideLayouts/_rels/slideLayout1.xml.rels`);
        this.content[`ppt/slideLayouts/_rels/slideLayout1.xml.rels`] = Xml.parse(xml);

        xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/slideLayouts/slideLayout1.xml`);
        this.content[`ppt/slideLayouts/slideLayout1.xml`] = Xml.parse(xml);
    }
}

module.exports.SlideLayoutsFactory = SlideLayoutsFactory;
