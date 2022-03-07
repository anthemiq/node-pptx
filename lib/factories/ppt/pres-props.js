const fs = require('fs');
const { Xml } = require('../../xmlnode');

class PresPropsFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/presProps.xml`);
        this.content['ppt/presProps.xml'] = Xml.parse(xml);
    }
}

module.exports.PresPropsFactory = PresPropsFactory;
