const fs = require('fs');
const { Xml } = require('../../xmlnode');

class TableStylesFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/tableStyles.xml`);
        this.content['ppt/tableStyles.xml'] = Xml.parse(xml);
    }
}

module.exports.TableStylesFactory = TableStylesFactory;
