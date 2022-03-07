let { ElementProperties } = require('./element-properties');

class Chart extends ElementProperties {
    constructor(args) {
        super();
        Object.assign(this, args);

        this.chartType = 'bar';

        this.cx(600);
        this.cy(400);
    }

    type(chartType) {
        this.chartType = chartType;

        return this;
    }

    data(chartData) {
        this.chartData = chartData;

        return this;
    }

    setContent(content) {
        this.content = content;
        super.setPropertyContent(this.content.get('p:xfrm'));
    }
}

module.exports.Chart = Chart;
