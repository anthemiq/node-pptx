const Parse = require('xml-js');

class Xml {
    static CHILD_KEY = "@";
    static ATTR_KEY = "$";

    constructor(xml) {
        this.xml = xml;
    }

    static create(name, children, attrs) {
        const node = new Xml({
            type: 'element',
            name,
        });
        if (children) {
            children?.forEach((child) => node.push(child));
        }
        if (attrs) {
            node.xml[Xml.ATTR_KEY] = attrs;
        }
        return node;
    }

    static createText(text) {
        return new Xml({ type: 'text', text });
    }

    static createTree(obj) {
        const nodes = [];
        Object.keys(obj).forEach((key) => {
            if (key === Xml.ATTR_KEY) {
                return;
            }
            let node;
            if (obj[key] === null) {
                node = Xml.create(key);
            } else if (typeof obj[key] === 'object') {
                node = Xml.create(key, Xml.createTree(obj[key]), obj[key][Xml.ATTR_KEY]);
            } else {
                node = Xml.create(key, null, obj[key][Xml.ATTR_KEY]);
                node.setText(obj[key]);
            }
            nodes.push(node);
        });
        return nodes;
    }

    static parse(xmlString) {
        const result = Parse.xml2js(xmlString, { elementsKey: Xml.CHILD_KEY, attributesKey: Xml.ATTR_KEY });
        return new Xml(result[this.CHILD_KEY][0]);
    }

    name() {
        return this.xml.name;
    }

    setName(name) {
        this.xml.name = name;
    }

    childCount() {
        return this.xml[Xml.CHILD_KEY] ? this.xml[Xml.CHILD_KEY].length : 0;
    }

    children() {
        return this.xml[Xml.CHILD_KEY]?.map((child) => new Xml(child));
    }

    get(childPath) {
        const names = childPath.split("/");
        return names.reduce((node, name) => {
            const child = node?.xml[Xml.CHILD_KEY]?.find((elem) => {
                return elem.type === "element" && elem.name === name;
            });
            return child ? new Xml(child) : null;
        }, this);
    }
    
    getAll(childName) {
        return this.xml[Xml.CHILD_KEY]?.filter((child) => child.name === childName);
    }

    at(index) {
        const child = this.xml[Xml.CHILD_KEY]?.[index];
        return child ? new Xml(child) : null;
    }

    forEach(fn) {
        this.xml[Xml.CHILD_KEY]?.forEach((child, index, array) => fn(new Xml(child), index, array));
    }

    filter(fn) {
        const result = this.xml[Xml.CHILD_KEY]?.filter((child, index, array) => fn(new Xml(child), index, array));
        return result?.map((elem) => new Xml(elem));
    }

    push(child) {
        if (!this.xml[Xml.CHILD_KEY]) {
            this.xml[Xml.CHILD_KEY] = [];
        }
        if (child instanceof Xml) {
            this.xml[Xml.CHILD_KEY].push(child.xml);
        } else {
            this.xml[Xml.CHILD_KEY].push(child);
        }
    }

    pushAll(arr) {
        arr?.forEach((elem) => this.push(elem));
    }

    insert(index, child) {
        if (child instanceof Xml) {
            this.xml[Xml.CHILD_KEY].splice(index, 0, child.xml);
        } else {
            this.xml[Xml.CHILD_KEY].splice(index, 0, child);
        }
    }

    replace(index, child) {
        if (child instanceof Xml) {
            this.xml[Xml.CHILD_KEY][index] = child.xml;
        } else {
            this.xml[Xml.CHILD_KEY][index] = child;
        }
    }

    remove(name) {
        if (this.xml[Xml.CHILD_KEY]) {
            const node = this.xml[Xml.CHILD_KEY][name];
            delete this.xml[Xml.CHILD_KEY][name];
            return node;
        } else {
            return null;
        }
    }

    removeAt(index) {
        return this.xml[Xml.CHILD_KEY].splice(index, 1)[0];
    }

    text() {
        return this.xml[Xml.CHILD_KEY][0].text;
    }

    setText(text) {
        this.xml[Xml.CHILD_KEY][0].text = text;
    }

    attr(name) {
        return this.xml[Xml.ATTR_KEY]?.[name];
    }

    setAttr(name, value) {
        if (!this.xml[Xml.ATTR_KEY]) {
            this.xml[Xml.ATTR_KEY] = {};
        }
        this.xml[Xml.ATTR_KEY][name] = value;
    }

    setAttrs(attrs) {
        this.xml[Xml.ATTR_KEY] = attrs;
    }

    removeAttr(name) {
        delete this.xml[Xml.ATTR_KEY][name];
    }

    clone() {
        const clone = JSON.parse(JSON.stringify(this));
        return new Xml(clone.xml);
    }

    serialize() {
        // We don't keep the empty root node that the underlying XML library expects, so recreate it.
        return Parse.js2xml(Xml.create('', [this.xml]).xml, { elementsKey: Xml.CHILD_KEY, attributesKey: Xml.ATTR_KEY });
    }
}

module.exports.Xml = Xml;
