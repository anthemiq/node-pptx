const fs = require('fs');
const { Xml } = require('../../xmlnode');

class AppFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../../fragments/docProps/app.xml`);
        this.content['docProps/app.xml'] = Xml.parse(xml);
    }

    setProperties(props) {
        if (props.company) this.content['docProps/app.xml'].get('Company').setText(props.company);
    }

    getProperties() {
        let props = {};
        let propertiesRef = this.content['docProps/app.xml'];

        if (propertiesRef?.get('Company')) {
            props.company = propertiesRef['Company'];
        }

        return props;
    }

    incrementSlideCount() {
        if (!this.content['docProps/app.xml'].get('Slides')) {
            this.content['docProps/app.xml'].push(Xml.create('Slides', [
                Xml.createText('0')
            ]));
        }

        let slideCount = +this.content['docProps/app.xml'].get('Slides').text();
        this.content['docProps/app.xml'].get('Slides').setText(`${slideCount + 1}`);
    }

    decrementSlideCount() {
        if (this.content['docProps/app.xml'].get('Slides')) {
            let slideCount = +this.content['docProps/app.xml'].get('Slides').text();
            this.content['docProps/app.xml'].get('Slides').setText(`${slideCount !== 0 ? slideCount - 1 : 0}`);
        }
    }

    setSlideCount(count) {
        if (!this.content['docProps/app.xml'].get('Slides')) {
            this.content['docProps/app.xml'].push(Xml.create('Slides', [
                Xml.createText('0')
            ]));
        }

        this.content['docProps/app.xml'].get('Slides').setText(`${count}`);
    }
}

module.exports.AppFactory = AppFactory;
