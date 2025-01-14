let { PptxUnitHelper } = require('./helpers/unit-helper');

const ShapeClass = Object.freeze({
    Text: 1,
    Image: 2,
    GraphicFrame: 3,
});

class ElementProperties {
    constructor() {
        this._x = 0;
        this._y = 0;
        this._cx = 0;
        this._cy = 0;
        this.options = {};
    }

    setPropertyContent(properties) {
        this.properties = properties;

        this.x(this._x);
        this.y(this._y);
        this.cx(this._cx);
        this.cy(this._cy);
    }

    x(val) {
        if (arguments.length === 0) {
            if (this.properties !== undefined) {
                return PptxUnitHelper.toPixels(this.properties.get('a:off').attr('x'));
            } else {
                return this._x;
            }
        } else {
            this._x = val;

            if (this.properties !== undefined) {
                this.properties.get('a:off').setAttr('x', PptxUnitHelper.fromPixels(val));
            }
        }

        return this;
    }

    y(val) {
        if (arguments.length === 0) {
            if (this.properties !== undefined) {
                return PptxUnitHelper.toPixels(this.properties.get('a:off').attr('y'));
            } else {
                return this._y;
            }
        } else {
            this._y = val;

            if (this.properties !== undefined) {
                this.properties.get('a:off').setAttr('y', PptxUnitHelper.fromPixels(val));
            }
        }

        return this;
    }

    cx(val) {
        if (arguments.length === 0) {
            if (this.properties !== undefined) {
                return PptxUnitHelper.toPixels(this.properties.get('a:ext').attr('cx'));
            } else {
                return this._cx;
            }
        } else {
            this._cx = val;

            if (this.properties !== undefined) {
                this.properties.get('a:ext').setAttr('cx', PptxUnitHelper.fromPixels(val));
            }
        }

        return this;
    }

    cy(val) {
        if (arguments.length === 0) {
            if (this.properties !== undefined) {
                return PptxUnitHelper.toPixels(this.properties.get('a:ext').attr('cy'));
            } else {
                return this._cy;
            }
        } else {
            this._cy = val;

            if (this.properties !== undefined) {
                this.properties.get('a:ext').setAttr('cy', PptxUnitHelper.fromPixels(val));
            }
        }

        return this;
    }
}

module.exports.ElementProperties = ElementProperties;
module.exports.ShapeClass = ShapeClass;
