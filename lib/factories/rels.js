const fs = require('fs');
const { Xml } = require('../xmlnode');

class RelsFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../fragments/_rels/.rels`);
        this.content['_rels/.rels'] = Xml.parse(xml);
    }
}

module.exports.RelsFactory = RelsFactory;
