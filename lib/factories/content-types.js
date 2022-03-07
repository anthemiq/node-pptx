const fs = require('fs');
const { Xml } = require('../xmlnode');

class ContentTypeFactory {
    constructor(parentFactory, args) {
        this.parentFactory = parentFactory;
        this.content = parentFactory.content;
        this.args = args;
    }

    build() {
        const xml = fs.readFileSync(`${__dirname}/../fragments/[Content_Types].xml`);
        this.content['[Content_Types].xml'] = Xml.parse(xml);

        this.addDefaultMediaContentTypes();
    }

    addDefaultMediaContentTypes() {
        // it's OK to have these content type definitions in the file even if they are not used anywhere in the pptx
        this.addDefaultContentType('png', 'image/png');
        this.addDefaultContentType('gif', 'image/gif');
        this.addDefaultContentType('jpg', 'image/jpg');
    }

    addDefaultContentType(extension, contentType) {
        let contentTypeExists = false;

        this.content['[Content_Types].xml'].forEach(function(element) {
            if (element.name = 'Default' && element.attr('Extension') === extension) {
                contentTypeExists = true;
                return;
            }
        });

        if (!contentTypeExists) {
            this.content['[Content_Types].xml'].push(Xml.create('Default', null, {
                Extension: extension,
                ContentType: contentType,
            }));
        }
    }

    addContentType(partName, contentType) {
        let contentTypeExists = false;

        this.content['[Content_Types].xml'].forEach(function(element) {
            if (element.name = 'Override' && element.attr('PartName') === partName && element.attr('ContentType') === contentType) {
                contentTypeExists = true;
                return;
            }
        });

        if (!contentTypeExists) {
            this.content['[Content_Types].xml'].push(Xml.create('Override', null, {
                PartName: partName,
                ContentType: contentType,
            }));
        }
    }

    removeContentType(partName, contentType) {
        let contentTypeIndex = -1;

        this.content['[Content_Types].xml'].forEach(function(element, index) {
            if (element.name = 'Override' && element.attr('PartName') === partName && element.attr('ContentType') === contentType) {
                contentTypeIndex = index;
                return;
            }
        });

        if (contentTypeIndex !== -1) {
            this.content['[Content_Types].xml'].removeAt(contentTypeIndex);
        }
    }
}

module.exports.ContentTypeFactory = ContentTypeFactory;
