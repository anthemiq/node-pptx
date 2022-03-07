const fs = require('fs');
const { Xml } = require('../../xmlnode');

class ThemeFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../../fragments/ppt/theme/theme1.xml`);
        this.content['ppt/theme/theme1.xml'] = Xml.parse(xml);
    }
}

module.exports.ThemeFactory = ThemeFactory;
