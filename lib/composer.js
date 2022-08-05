let { Presentation } = require('./presentation');

class Composer {
    constructor() {
        this.presentation = new Presentation();
    }

    getSlide(slideNameOrNumber) {
        return this.presentation.getSlide(slideNameOrNumber);
    }

    async loadFromFile(filePath) {
        await this.presentation.loadFromFile(filePath);
        return this;
    }

    async loadFromBuffer(buffer) {
        await this.presentation.loadFromBuffer(buffer);
        return this;
    }

    async compose(func) {
        await func(this.presentation);
        return this;
    }

    async save(destination) {
        await this.presentation.save(destination);
        return this;
    }
}

module.exports.Composer = Composer;
