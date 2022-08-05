/* eslint-disable no-prototype-builtins */

let { Shape } = require('./shape');
let { Image, RenderedImage } = require('./image');
let { Chart } = require('./chart');
let { TextBox, RenderedTextBox } = require('./text-box');
const { GraphicFrame, RenderedGraphicFrame } = require('./graphic-frame');

class Slide {
    constructor(args) {
        Object.assign(this, args);

        this.content = args.content;
        this.powerPointFactory = args.powerPointFactory;
        this.name = args.name;
        this.externalObjectCount = args.externalObjectCount || 0;
        this.highestObjectId = args.highestObjectId || 0;
        this.layoutName = args.layoutName || 'slideLayout1';

        this.elements = []; // Currently not used for anything.
        this.defaults = {};
    }

    layout(layoutName) {
        this.layoutName = layoutName;

        return this;
    }

    getLayout() {
        return this.layoutName;
    }

    textColor(color) {
        this.defaults.textColor = color;
    }

    backgroundColor(color) {
        this.powerPointFactory.setBackgroundColor(this, color);
    }

    processConfig(config, pptxObject) {
        if (typeof config === 'function') {
            config(pptxObject);
        } else if (typeof config === 'object') {
            // calls the corresponding setter functions if the user passed in a "property object"
            for (let configKey in config) {
                if (config.hasOwnProperty(configKey)) {
                    if (typeof pptxObject[configKey] === 'function') {
                        pptxObject[configKey](config[configKey]);
                    }
                }
            }
        } else {
            throw new Error('Invalid config passed to Slide.processConfig().');
        }
    }

    async addImage(config) {
        let image = new Image();

        try {
            this.processConfig(config, image);
        } catch (err) {
            throw new Error(`Exception in Slide.addImage() when calling this.processConfig(). ${err.message}`);
        }

        try {
            if (image.sourceType === 'file' || image.sourceType === 'base64') {
                this.powerPointFactory.addImage(this, image);
            } else if (image.sourceType === 'url') {
                await this.powerPointFactory.addImageFromRemoteUrl(this, image);
            }

            this.elements.push(image);

            return this;
        } catch (err) {
            let imageSource = '(base64 binary)';

            if (image.sourceType === 'file') {
                imageSource = image.source;
            } else if (image.sourceType === 'url') {
                imageSource = image.downloadUr;
            }

            throw new Error(`Failed to add image to slide. Image source: ${imageSource}. Exception info: ${err.message}`);
        }
    }

    addText(config) {
        try {
            let textBox = new TextBox();

            this.processConfig(config, textBox);

            // need to make a copy of defaults first, then merge options into that copy so the original defaults object stays immutable
            textBox.options = Object.assign(Object.assign({}, this.defaults), textBox.options);

            this.powerPointFactory.addText(this, textBox);
            this.elements.push(textBox);

            return this;
        } catch (err) {
            console.log(err);
            throw new Error(`Failed to add text to slide. Exception info: ${err.message}`);
        }
    }

    addShape(config) {
        try {
            let shape = new Shape();

            this.processConfig(config, shape);
            this.powerPointFactory.addShape(this, shape);
            this.elements.push(shape);

            return this;
        } catch (err) {
            throw new Error(`Failed to add shape to slide. Exception info: ${err.message}`);
        }
    }

    async addChart(config) {
        try {
            let chart = new Chart();

            this.processConfig(config, chart);
            await this.powerPointFactory.addChart(this, chart);
            this.elements.push(chart);

            return this;
        } catch (err) {
            throw new Error(`Failed to add chart to slide. Exception info: ${err.message}`);
        }
    }

    moveTo(destinationSlideNum) {
        try {
            let thisSlideNum = Number(this.name.replace('slide', ''));

            this.powerPointFactory.moveSlide(thisSlideNum, destinationSlideNum);
        } catch (err) {
            throw new Error(`Failed to move slide to new position #: ${destinationSlideNum}. Exception info: ${err.message}`);
        }
    }

    rename(newName) {
        this.name = newName;
    }

    getSlideXmlAsString() {
        return this.content.serialize();
    }

    getNumElements() {
        return this.content.get('p:cSld/p:spTree').childCount();
    }

    getNextObjectId() {
        return ++this.highestObjectId;
    }

    getShapes() {
        const renderedShapes = [];
        this.content.get('p:cSld/p:spTree').forEach((elem, index) => {
            let shape = this._createShape(elem, index);
            if (shape) {
                renderedShapes.push(shape);
            }
        });
        return renderedShapes;
    }

    getShapeAt(index) {
      const elem = this.content.get('p:cSld/p:spTree').at(index);
      if (elem) {
        return this._createShape(elem, index);
      } else {
        return null;
      }
    }

    insertShapeAt(index, shape, shapes) {
      this._insertElementAt(index, shape.content);
      shape.index = index;

      if (shapes) {
        shapes.forEach((s) => {
          if (s !== shape && s.index >= index) {
            ++s.index;
          }
        });
      }
    }

    _insertElementAt(index, element) {
        const treeNode = this.content.get('p:cSld/p:spTree');
        treeNode.insert(index, element);
    }

    removeShape(shape, shapes) {
        this._removeElementAt(shape.index);

        if (shapes) {
          shapes.forEach((s) => {
            if (s.index > shape.index) {
              --s.index;
            }
          });
        }

        return shape;
    }

    _removeElementAt(index) {
        const treeNode = this.content.get('p:cSld/p:spTree');
        return treeNode.removeAt(index);
    }

    replaceShape(oldShape, newShape) {
        const treeNode = this.content.get('p:cSld/p:spTree');
        treeNode.replace(oldShape.index, newShape.content);
        newShape.index = oldShape.index;
    }

    _createShape(element, index) {
        switch (element.name()) {
            case 'p:sp':
                return new RenderedTextBox(element, index);
            case 'p:pic':
                return new RenderedImage(element, index);
            case 'p:graphicFrame':
                return new RenderedGraphicFrame(element, index);
        }
        return null;
    }

    addHyperlinkToShape(shape, url, index) {
        const rId = this.powerPointFactory.pptFactory.slideFactory.addHyperlinkToSlideRelationship(this,  url);
        shape.addHyperlink(rId, index);
    }
}

module.exports.Slide = Slide;
