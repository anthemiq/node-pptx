let { ElementProperties, ShapeClass } = require('./element-properties');

class RenderedGraphicFrame extends ElementProperties {
    constructor(content) {
        super();
        this.content = content;
        this.properties = this.content.get('a:xfrm');
    }

    class() {
        return ShapeClass.GraphicFrame;
    }
}

module.exports.RenderedGraphicFrame = RenderedGraphicFrame;
