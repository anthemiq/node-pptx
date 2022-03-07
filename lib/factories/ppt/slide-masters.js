const fs = require('fs');
const { Xml } = require('../../xmlnode');

class SlideMastersFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        let xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/slideMasters/_rels/slideMaster1.xml.rels`);
        this.content['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = Xml.parse(xml);

        xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/slideMasters/slideMaster1.xml`);
        this.content['ppt/slideMasters/slideMaster1.xml'] = Xml.parse(xml);
    }
}

module.exports.SlideMastersFactory = SlideMastersFactory;
