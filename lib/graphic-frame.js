let { ElementProperties, ShapeClass } = require('./element-properties');
const { RenderedTextBox } = require('./text-box');

class RenderedGraphicFrame extends ElementProperties {
    constructor(content, index) {
        super();
        this.content = content;
        this.index = index;
        this.properties = this.content.get('a:xfrm');
        this.rows = this.content.get('a:graphic/a:graphicData/a:tbl').filter((child) => child.name() === 'a:tr');
    }

    class() {
        return ShapeClass.GraphicFrame;
    }

    rowCount() {
        return this.rows?.length;
    }

    columnCount() {
        return this.content.get('a:graphic/a:graphicData/a:tbl/a:tblGrid').filter((child) => child.name() === 'a:gridCol').length;
    }

    shape(row, col) {
        return new RenderedTextBox(this.rows[row].filter((child) => child.name() === 'a:tc')[col], -1, 'a');
    }
}

module.exports.RenderedGraphicFrame = RenderedGraphicFrame;
