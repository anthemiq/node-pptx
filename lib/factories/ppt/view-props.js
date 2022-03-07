const fs = require('fs');
const { Xml } = require('../../xmlnode');

class ViewPropsFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/viewProps.xml`);
        this.content['ppt/viewProps.xml'] = Xml.parse(xml);
    }
}

module.exports.ViewPropsFactory = ViewPropsFactory;
