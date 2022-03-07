const fs = require('fs');
const { Xml } = require('../../xmlnode');

class CoreFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../../fragments/docProps/core.xml`);
        this.content['docProps/core.xml'] = Xml.parse(xml);

        //this.updateTimeStamps(); // FIXME: doesn't work for some reason...
    }

    setProperties(props) {
        if (props.title) this.content['docProps/core.xml'].get('dc:title').setText(props.title);
        if (props.author) this.content['docProps/core.xml'].get('dc:creator').setText(props.author);
        //if (props.revision) this.content['docProps/core.xml'].get('dc:revision').setText(props.revision); // FIXME: doesn't work for some reason, causes corrupt file
        if (props.subject) this.content['docProps/core.xml'].get('dc:subject').setText(props.subject);
    }

    getProperties() {
        let props = {};

        props.title = this.content['docProps/core.xml'].get('dc:title').text();
        props.author = this.content['docProps/core.xml'].get('dc:creator').text();
        props.revision = this.content['docProps/core.xml'].get('dc:revision').text();
        props.subject = this.content['docProps/core.xml'].get('dc:subject').text();

        return props;
    }

    updateTimeStamps() {
        this.updateCreatedDateTimeStamp();
        this.updatedModifiedDateTimeStamp();
    }

    // for now we won't need to update created date without modified date because every save is considered a "new" pptx, but these functions are separated for future support
    updateCreatedDateTimeStamp() {
        this.content['docProps/core.xml'].get('dcterms:created').setText(new Date().toISOString());
    }

    updatedModifiedDateTimeStamp() {
        this.content['docProps/core.xml'].get('dcterms:modified').setText(new Date().toISOString());
    }
}

module.exports.CoreFactory = CoreFactory;
