
export class VirtualSlide {
    public elements: any[] = [];
    public width = 960;
    public height = 540;
    public background = { color: '#ffffff' };

    constructor() { }

    getBackground() {
        return {
            setSolidFill: (hex: string) => this.background.color = hex
        };
    }

    insertImage(blob: any) {
        // Mock image insertion - treating it like a rect for now
        const shape = new VirtualShape('IMAGE', 0, 0, 100, 100);
        this.elements.push(shape);
        return {
            getWidth: () => shape.width,
            getHeight: () => shape.height,
            setWidth: (w: number) => { shape.width = w; return this; },
            setHeight: (h: number) => { shape.height = h; return this; },
            setLeft: (l: number) => { shape.left = l; return this; },
            setTop: (t: number) => { shape.top = t; return this; },
            remove: () => { this.elements = this.elements.filter(e => e !== shape); }
        };
    }

    insertLine(type: any, x1: number, y1: number, x2: number, y2: number) {
        const line = new VirtualLine(x1, y1, x2, y2);
        this.elements.push(line);
        return line;
    }

    insertShape(type: any, left: number, top: number, width: number, height: number) {
        const shape = new VirtualShape(type, left, top, width, height);
        this.elements.push(shape);
        return shape;
    }

    insertTable(rows: number, cols: number) {
        const table = new VirtualTable(rows, cols);
        this.elements.push(table);
        return table;
    }

    getPlaceholder(type: any) {
        // Return a dummy placeholder if asked, mostly for title
        if (String(type) === 'TITLE' || String(type) === 'CENTERED_TITLE') {
            const shape = new VirtualShape('RECTANGLE', 50, 20, 800, 60);
            shape.placeholderType = type;
            this.elements.push(shape);
            return {
                asShape: () => shape,
                remove: () => {
                    this.elements = this.elements.filter(e => e !== shape);
                }
            };
        }
        return null;
    }

    getPlaceholders() {
        return [];
    }

    getPageElements() {
        return this.elements.map(e => ({
            getObjectId: () => e.id,
            asShape: () => e instanceof VirtualShape ? e : null
        }));
    }

    group(elements: any[]) {
        // In virtual slide, grouping might just be a logical thing, 
        // but for now we can ignore deep grouping logic or just return a dummy group
        // The generator uses group to center things.

        // We can compute the bounding box of these elements to support 'getWidth/getHeight'
        // which are used for centering.
        const ids = elements.map(e => e.getObjectId());
        const targetElements = this.elements.filter(e => ids.includes(e.id));

        if (targetElements.length === 0) return null;

        return new VirtualGroup(targetElements);
    }
}

export class VirtualElement {
    public id = Math.random().toString(36).substring(7);
    public top: number;
    public left: number;
    public width: number;
    public height: number;

    constructor(left: number, top: number, width: number, height: number) {
        this.left = left;
        this.top = top;
        this.width = width;
        this.height = height;
    }

    getObjectId() { return this.id; }
    getWidth() { return this.width; }
    getHeight() { return this.height; }
    getLeft() { return this.left; }
    getTop() { return this.top; }
    setLeft(l: number) { this.left = l; }
    setTop(t: number) { this.top = t; }
}

export class VirtualGroup extends VirtualElement {
    constructor(public children: VirtualElement[]) {
        // Compute bounding box
        const minX = Math.min(...children.map(c => c.left));
        const maxX = Math.max(...children.map(c => c.left + c.width));
        const minY = Math.min(...children.map(c => c.top));
        const maxY = Math.max(...children.map(c => c.top + c.height));
        super(minX, minY, maxX - minX, maxY - minY);
    }

    setLeft(newLeft: number) {
        const diff = newLeft - this.left;
        this.left = newLeft;
        this.children.forEach(c => c.setLeft(c.left + diff));
    }

    setTop(newTop: number) {
        const diff = newTop - this.top;
        this.top = newTop;
        this.children.forEach(c => c.setTop(c.top + diff));
    }
}

export class VirtualLine extends VirtualElement {
    strokeColor = '#000000';
    weight = 1;
    x1: number;
    y1: number;
    x2: number;
    y2: number;

    constructor(x1: number, y1: number, x2: number, y2: number) {
        const minX = Math.min(x1, x2);
        const minY = Math.min(y1, y2);
        super(minX, minY, Math.abs(x2 - x1), Math.abs(y2 - y1));
        this.x1 = x1;
        this.y1 = y1;
        this.x2 = x2;
        this.y2 = y2;
    }

    // Override setLeft/Top to update line coords
    setLeft(l: number) {
        const diff = l - this.left;
        super.setLeft(l);
        this.x1 += diff;
        this.x2 += diff;
    }

    setTop(t: number) {
        const diff = t - this.top;
        super.setTop(t);
        this.y1 += diff;
        this.y2 += diff;
    }

    getLineFill() {
        return {
            setSolidFill: (hex: string) => this.strokeColor = hex
        };
    }

    setWeight(w: number) { this.weight = w; }
}

export class VirtualShape extends VirtualElement {
    type: string;
    fillColor = '#cccccc';
    borderColor = '#000000';
    borderWeight = 1;
    text = '';
    textStyle: any = { size: 12, color: '#000000', bold: false, align: 'LEFT' };
    placeholderType: string | null = null;
    contentAlignment = 'TOP';

    constructor(type: string, left: number, top: number, width: number, height: number) {
        super(left, top, width, height);
        this.type = String(type);
        if (this.type === 'TEXT_BOX') {
            this.fillColor = 'TRANSPARENT';
            this.borderColor = 'TRANSPARENT';
        }
    }

    getFill() {
        return {
            setSolidFill: (hex: string) => this.fillColor = hex,
            setTransparent: () => this.fillColor = 'TRANSPARENT'
        };
    }

    getBorder() {
        return {
            getLineFill: () => ({
                setSolidFill: (hex: string) => this.borderColor = hex
            }),
            setWeight: (w: number) => this.borderWeight = w,
            setTransparent: () => this.borderColor = 'TRANSPARENT'
        };
    }

    getText() {
        return {
            setText: (t: string) => this.text = t,
            asString: () => this.text,
            getRange: () => ({
                getTextStyle: () => this.getTextStyleMock()
            }),
            getTextStyle: () => this.getTextStyleMock(),
            getParagraphStyle: () => ({
                setParagraphAlignment: (a: any) => this.textStyle.align = String(a)
            })
        };
    }

    private getTextStyleMock() {
        const mock = {
            setFontSize: (s: number) => { this.textStyle.size = s; return mock; },
            setForegroundColor: (c: string) => { this.textStyle.color = c; return mock; },
            setBold: (b: boolean) => { this.textStyle.bold = b; return mock; },
            setFontFamily: () => mock,
        };
        return mock;
    }

    setContentAlignment(align: any) {
        this.contentAlignment = String(align);
    }

    getPlaceholderType() { return this.placeholderType; }
    asShape() { return this; }
}

export class VirtualTable extends VirtualElement {
    rows: number;
    cols: number;
    cells: VirtualCell[][];

    constructor(rows: number, cols: number) {
        super(0, 0, 100, 100);
        this.rows = rows;
        this.cols = cols;
        this.cells = [];
        for (let r = 0; r < rows; r++) {
            const rowArr = [];
            for (let c = 0; c < cols; c++) {
                rowArr.push(new VirtualCell());
            }
            this.cells.push(rowArr);
        }
    }

    getCell(r: number, c: number) {
        return this.cells[r][c];
    }
}

export class VirtualCell {
    text = '';
    fillColor = 'TRANSPARENT';
    borders = {
        top: new VirtualBorder(),
        bottom: new VirtualBorder(),
        left: new VirtualBorder(),
        right: new VirtualBorder()
    };
    textStyle: any = { size: 10, color: '#000000', bold: false, align: 'LEFT', family: 'Arial' };
    contentAlignment = 'TOP';

    getText() {
        return {
            setText: (t: string) => this.text = t,
            getTextStyle: () => this.getTextStyleMock()
        };
    }

    private getTextStyleMock() {
        const mock = {
            setFontSize: (s: number) => { this.textStyle.size = s; return mock; },
            setForegroundColor: (c: string) => { this.textStyle.color = c; return mock; },
            setBold: (b: boolean) => { this.textStyle.bold = b; return mock; },
            setFontFamily: (f: string) => { this.textStyle.family = f; return mock; },
        };
        return mock;
    }

    getFill() {
        return {
            setSolidFill: (hex: string) => this.fillColor = hex,
            setTransparent: () => this.fillColor = 'TRANSPARENT'
        };
    }

    getBorderTop() { return this.borders.top; }
    getBorderBottom() { return this.borders.bottom; }
    getBorderLeft() { return this.borders.left; }
    getBorderRight() { return this.borders.right; }

    setContentAlignment(a: any) { this.contentAlignment = String(a); }
}

export class VirtualBorder {
    color = '#000000';
    weight = 1;
    isVisible = true;

    setSolidFill(hex: string) { this.color = hex; this.isVisible = true; }
    setTransparent() { this.isVisible = false; }
    setWeight(w: number) { this.weight = w; }
}
