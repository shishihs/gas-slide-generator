import { SlideTheme } from '../config/SlideTheme';
import { DEFAULT_THEME } from '../config/DefaultTheme';
import { hexToRgb } from './ColorUtils';

/**
 * Get theme from layout or use default
 */
function getThemeFromLayout(layout: any): SlideTheme {
    if (layout && typeof layout.getTheme === 'function') {
        return layout.getTheme();
    }
    return DEFAULT_THEME;
}

export function applyTextStyle(textRange: GoogleAppsScript.Slides.TextRange, opt: { color?: string; size?: number; bold?: boolean; align?: GoogleAppsScript.Slides.ParagraphAlignment; fontType?: 'large' | 'small' }, theme: SlideTheme = DEFAULT_THEME) {
    const style = textRange.getTextStyle();
    let defaultColor;
    if (opt.fontType === 'large') {
        defaultColor = theme.colors.textPrimary;
    } else {
        defaultColor = theme.colors.textSmallFont;
    }
    style.setFontFamily(theme.fonts.family)
        .setForegroundColor(opt.color || defaultColor)
        .setFontSize(opt.size || theme.fonts.sizes.body)
        .setBold(opt.bold || false);
    if (opt.align) {
        try {
            textRange.getParagraphs().forEach(p => {
                p.getRange().getParagraphStyle().setParagraphAlignment(opt.align!);
            });
        } catch (e) { }
    }
}

export function setStyledText(shapeOrCell: GoogleAppsScript.Slides.Shape | GoogleAppsScript.Slides.TableCell, rawText: string, baseOpt?: any, theme: SlideTheme = DEFAULT_THEME) {
    const parsed = parseInlineStyles(rawText || '', theme.colors.primary);
    const tr = shapeOrCell.getText();
    tr.setText(parsed.output);
    applyTextStyle(tr, baseOpt || {}, theme);
    applyStyleRanges(tr, parsed.ranges);
}

export function setBulletsWithInlineStyles(shape: GoogleAppsScript.Slides.Shape, points: string[], theme: SlideTheme = DEFAULT_THEME) {
    const joiner = '\n\n';
    let combined = '';
    const ranges: any[] = [];
    (points || []).forEach((pt, idx) => {
        const parsed = parseInlineStyles(String(pt || ''), theme.colors.primary);
        const bullet = parsed.output;
        if (idx > 0) combined += joiner;
        const start = combined.length;
        combined += bullet;
        parsed.ranges.forEach(r => ranges.push({
            start: start + r.start,
            end: start + r.end,
            bold: r.bold,
            color: r.color
        }));
    });
    const tr = shape.getText();
    tr.setText(combined || '—');
    applyTextStyle(tr, {
        size: theme.fonts.sizes.body
    }, theme);
    try {
        tr.getParagraphs().forEach(p => {
            p.getRange().getParagraphStyle().setLineSpacing(130).setSpaceBelow(12);
        });
    } catch (e) { }
    applyStyleRanges(tr, ranges);
}

function checkSpacing(s: string, out: string, i: number, nextCharIndex: number) {
    let prefix = '';
    let suffix = '';
    if (out.length > 0 && !/\s$/.test(out)) {
        prefix = ' ';
    }
    if (nextCharIndex < s.length && !/^\s/.test(s[nextCharIndex])) {
        suffix = ' ';
    }
    return { prefix, suffix };
}

function parseInlineStyles(s: string, highlightColor: string) {
    const ranges = [];
    let out = '';
    let i = 0;
    while (i < s.length) {
        if (s[i] === '*' && s[i + 1] === '*' &&
            s[i + 2] === '[' && s[i + 3] === '[') {
            const contentStart = i + 4;
            const close = s.indexOf(']]**', contentStart);
            if (close !== -1) {
                let content = s.substring(contentStart, close);
                const nextCharIndex = close + 4;
                const { prefix, suffix } = checkSpacing(s, out, i, nextCharIndex);
                out += prefix;
                const start = out.length;
                content += suffix;
                out += content;
                const end = out.length - suffix.length;
                const rangeObj = {
                    start,
                    end,
                    bold: true,
                    color: highlightColor,
                };
                ranges.push(rangeObj);
                i = nextCharIndex;
                continue;
            }
        }
        if (s[i] === '[' && s[i + 1] === '[') {
            const close = s.indexOf(']]', i + 2);
            if (close !== -1) {
                let content = s.substring(i + 2, close);
                const nextCharIndex = close + 2;
                const { prefix, suffix } = checkSpacing(s, out, i, nextCharIndex);
                out += prefix;
                const start = out.length;
                content += suffix;
                out += content;
                const end = out.length - suffix.length;
                const rangeObj = {
                    start,
                    end,
                    bold: true,
                    color: highlightColor,
                };
                ranges.push(rangeObj);
                i = nextCharIndex;
                continue;
            }
        }
        if (s[i] === '*' && s[i + 1] === '*') {
            const close = s.indexOf('**', i + 2);
            if (close !== -1) {
                let content = s.substring(i + 2, close);
                if (content.indexOf('[[') === -1) {
                    const nextCharIndex = close + 2;
                    const { prefix, suffix } = checkSpacing(s, out, i, nextCharIndex);
                    out += prefix;
                    const start = out.length;
                    content += suffix;
                    out += content;
                    const end = out.length - suffix.length;
                    ranges.push({
                        start,
                        end,
                        bold: true,
                    });
                    i = nextCharIndex;
                    continue;
                } else {
                    i += 2;
                    continue;
                }
            }
        }
        out += s[i];
        i++;
    }
    return {
        output: out,
        ranges
    };
}

function applyStyleRanges(textRange: GoogleAppsScript.Slides.TextRange, ranges: any[]) {
    ranges.forEach(r => {
        try {
            const sub = textRange.getRange(r.start, r.end);
            if (sub.isEmpty()) return;
            const st = sub.getTextStyle();
            if (r.bold) st.setBold(true);
            if (r.color) st.setForegroundColor(r.color);
            if (r.size) st.setFontSize(r.size);
        } catch (e) { }
    });
}



export function insertImageFromUrlOrFileId(urlOrFileId: string): GoogleAppsScript.Base.BlobSource | null {
    if (!urlOrFileId || typeof urlOrFileId !== 'string') return null;
    function extractFileIdFromUrl(url: string) {
        if (!url || typeof url !== 'string') return null;
        const patterns = [
            /\/file\/d\/([a-zA-Z0-9_-]+)/,
            /id=([a-zA-Z0-9_-]+).*file/,
            /file\/([a-zA-Z0-9_-]+)/
        ];
        for (const pattern of patterns) {
            const match = url.match(pattern);
            if (match && match[1]) return match[1];
        }
        return null;
    }
    const fileIdPattern = /^[a-zA-Z0-9_-]{25,}$/;
    const extractedFileId = extractFileIdFromUrl(urlOrFileId);
    if (extractedFileId && fileIdPattern.test(extractedFileId)) {
        try {
            const file = DriveApp.getFileById(extractedFileId);
            return file.getBlob();
        } catch (e) {
            return null;
        }
    } else if (fileIdPattern.test(urlOrFileId)) {
        try {
            const file = DriveApp.getFileById(urlOrFileId);
            return file.getBlob();
        } catch (e) {
            return null;
        }
    } else if (urlOrFileId.startsWith('data:image/')) {
        try {
            const parts = urlOrFileId.split(',');
            if (parts.length !== 2) throw new Error('Invalid Base64 format.');
            const mimeMatch = parts[0].match(/:(.*?);/);
            if (!mimeMatch) return null;
            const mimeType = mimeMatch[1];
            const base64Data = parts[1];
            const decodedData = Utilities.base64Decode(base64Data);
            return Utilities.newBlob(decodedData, mimeType);
        } catch (e) {
            return null;
        }
    } else {
        // Attempt to just download it if it is a public URL, or leave if it's too complex. 
        // GAS insertImage can take URL, but this function returns Blob.
        // If it is regular URL, we might want to fetch it.
        try {
            // Naive fetch for URL
            if (urlOrFileId.startsWith('http')) {
                const response = UrlFetchApp.fetch(urlOrFileId);
                return response.getBlob();
            }
        } catch (e) { }

        return null;
    }
}

function createGradientRectangle(slide: GoogleAppsScript.Slides.Slide, x: number, y: number, width: number, height: number, colors: string[]) {
    const numStrips = Math.max(20, Math.floor(width / 2));
    const stripWidth = width / numStrips;
    const startColor = hexToRgb(colors[0]),
        endColor = hexToRgb(colors[1]);
    if (!startColor || !endColor) return null;
    const shapes = [];
    for (let i = 0; i < numStrips; i++) {
        const ratio = i / (numStrips - 1);
        const r = Math.round(startColor.r + (endColor.r - startColor.r) * ratio);
        const g = Math.round(startColor.g + (endColor.g - startColor.g) * ratio);
        const b = Math.round(startColor.b + (endColor.b - startColor.b) * ratio);
        const strip = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x + (i * stripWidth), y, stripWidth + 0.5, height);
        strip.getFill().setSolidFill(r, g, b);
        strip.getBorder().setTransparent();
        shapes.push(strip);
    }
    if (shapes.length > 1) {
        try {
            // @ts-ignore
            if (slide.group) return slide.group(shapes);
        } catch (e) { }
    }
    return shapes[0] || null;
}

function createPillShapeUnderline(slide: GoogleAppsScript.Slides.Slide, x: number, y: number, width: number, height: number, settings: any) {
    const shapes = [];
    const diameter = height;
    const radius = height / 2;
    const rectWidth = Math.max(0, width - diameter);
    if (width < diameter) {
        const centerCircle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, diameter, diameter);
        const color = settings.enableGradient ? settings.gradientStart : settings.primaryColor;
        centerCircle.getFill().setSolidFill(color);
        centerCircle.getBorder().setTransparent();
        return;
    }
    const leftCircle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, diameter, diameter);
    shapes.push(leftCircle);
    let mainShape;
    if (rectWidth > 0) {
        if (settings.enableGradient) {
            mainShape = createGradientRectangle(
                slide,
                x + radius,
                y,
                rectWidth,
                height,
                [settings.gradientStart, settings.gradientEnd]
            );
            if (mainShape) shapes.push(mainShape as GoogleAppsScript.Slides.Shape);
        } else {
            mainShape = slide.insertShape(
                SlidesApp.ShapeType.RECTANGLE,
                x + radius,
                y,
                rectWidth,
                height
            );
            mainShape.getFill().setSolidFill(settings.primaryColor);
            mainShape.getBorder().setTransparent();
            shapes.push(mainShape);
        }
    }
    const rightCircle = slide.insertShape(
        SlidesApp.ShapeType.ELLIPSE,
        x + width - diameter,
        y,
        diameter,
        diameter
    );
    shapes.push(rightCircle);
    if (settings.enableGradient) {
        leftCircle.getFill().setSolidFill(settings.gradientStart);
        rightCircle.getFill().setSolidFill(settings.gradientEnd);
    } else {
        leftCircle.getFill().setSolidFill(settings.primaryColor);
        rightCircle.getFill().setSolidFill(settings.primaryColor);
    }
    leftCircle.getBorder().setTransparent();
    rightCircle.getBorder().setTransparent();
    if (shapes.length > 1) {
        try {
            // @ts-ignore
            if (slide.group) slide.group(shapes);
        } catch (e) {
        }
    }
}


export function normalizeImages(arr: any[]) {
    return (arr || []).map(v => typeof v === 'string' ? {
        url: v
    } : (v && v.url ? v : null)).filter(Boolean).slice(0, 6);
}

export function renderImagesInArea(slide: GoogleAppsScript.Slides.Slide, layout: any, area: any, images: any[], imageUpdateOption: string = 'update') {
    if (!images || !images.length) return;
    if (imageUpdateOption !== 'update') {
        return;
    }
    const n = Math.min(6, images.length);
    let cols = n === 1 ? 1 : (n <= 4 ? 2 : 3);
    const rows = Math.ceil(n / cols);
    const gap = layout.pxToPt(10);
    const cellW = (area.width - gap * (cols - 1)) / cols,
        cellH = (area.height - gap * (rows - 1)) / rows;
    for (let i = 0; i < n; i++) {
        const r = Math.floor(i / cols),
            c = i % cols;
        try {
            const img = slide.insertImage(images[i].url);
            const scale = Math.min(cellW / img.getWidth(), cellH / img.getHeight());
            const w = img.getWidth() * scale,
                h = img.getHeight() * scale;
            img.setWidth(w).setHeight(h).setLeft(area.left + c * (cellW + gap) + (cellW - w) / 2).setTop(area.top + r * (cellH + gap) + (cellH - h) / 2);
        } catch (e) { }
    }
}

export function drawArrowBetweenRects(slide: GoogleAppsScript.Slides.Slide, a: any, b: any, arrowH: number, arrowGap: number, settings: any) {
    const fromX = a.left + a.width;
    const fromY = a.top + a.height / 2;
    const toX = b.left;
    const toY = b.top + b.height / 2;
    if (toX - fromX <= 0) return;
    const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, fromX, fromY, toX, toY);
    line.getLineFill().setSolidFill(settings.primaryColor);
    line.setWeight(1.5);
    line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
}

export function adjustShapeText_External(shape: GoogleAppsScript.Slides.Shape, preCalculatedWidthPt: number | null = null, widthOverride: number | null = null, heightOverride: number | null = null) {
    const PADDING_TOP_BOTTOM = 7.5;
    const PADDING_LEFT_RIGHT = 10.0;

    function getEffectiveCharCount(text: string) {
        let count = 0;
        if (!text) return 0;
        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            if (char.match(/[^\x00-\x7F\uFF61-\uFF9F]/)) {
                count += 1.0;
            } else {
                count += 0.6;
            }
        }
        return count;
    }

    function _isShapeShortBox(shape: GoogleAppsScript.Slides.Shape, baseFontSize: number | null, heightOverride: number | null = null) {
        if (!baseFontSize || baseFontSize === 0) {
            return false;
        }
        const boxHeight = (heightOverride !== null) ? heightOverride : shape.getHeight();
        return (boxHeight <= (baseFontSize * 2));
    }

    // Note: Simplified for portability. Original had more logic.
    // ...

    return { isOverflow: false, details: 'Simplified implementation in cleanup.' };
}

/**
 * Add footer to slide with optional page number
 */
export function addFooter(slide: GoogleAppsScript.Slides.Slide, layout: any, pageNum: number, settings: any, creditImageBlob: GoogleAppsScript.Base.BlobSource | null) {
    const theme = getThemeFromLayout(layout);

    if (theme.footerText && theme.footerText.trim() !== '') {
        const leftRect = layout.getRect('footer.leftText');
        const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftRect.left, leftRect.top, leftRect.width, leftRect.height);
        const tr = leftShape.getText();
        tr.setText(theme.footerText);
        applyTextStyle(tr, {
            size: theme.fonts.sizes.footer,
            fontType: 'large'
        });
        try {
            leftShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) { }
    }

    if (pageNum > 0 && settings && settings.showPageNumber) {
        const rightRect = layout.getRect('footer.rightPage');
        const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightRect.left, rightRect.top, rightRect.width, rightRect.height);
        const tr = rightShape.getText();
        tr.setText(String(pageNum));
        applyTextStyle(tr, {
            size: theme.fonts.sizes.footer,
            color: theme.colors.primary,
            align: SlidesApp.ParagraphAlignment.END
        });
        try {
            rightShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) { }
    }
}

function safeGetRect(layout: any, path: string) {
    try {
        const rect = layout.getRect(path);
        if (rect &&
            (typeof rect.left === 'number' || rect.left === undefined) &&
            typeof rect.top === 'number' &&
            typeof rect.width === 'number' &&
            typeof rect.height === 'number') {
            if (rect.left === undefined) {
                return null;
            }
            return rect;
        }
        return null;
    } catch (e) {
        return null;
    }
}

function estimateTextWidthPt(text: string, fontSize: number): number {
    let count = 0;
    if (!text) return 0;
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        if (char.match(/[^\x00-\x7F\uFF61-\uFF9F]/)) {
            count += 1.0;
        } else {
            count += 0.6;
        }
    }
    return count * fontSize;
}

export function drawStandardTitleHeader(slide: GoogleAppsScript.Slides.Slide, layout: any, key: string, title: string, settings: any, preCalculatedWidthPt: number | null = null, imageUpdateOption: string = 'update') {
    const theme = getThemeFromLayout(layout);
    const titleRect = safeGetRect(layout, `${key}.title`);
    if (!titleRect) {
        return;
    }
    const initialFontSize = theme.fonts.sizes.contentTitle;
    const optimalHeight = layout.pxToPt(initialFontSize + 8);
    const cmToPt = 28.3465;
    const verticalShiftPt = 0.3 * cmToPt;
    const adjustedTop = titleRect.top + verticalShiftPt;
    const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
        titleRect.left,
        adjustedTop,
        titleRect.width,
        optimalHeight
    );
    setStyledText(titleShape, title || '', {
        size: initialFontSize,
        bold: true,
        fontType: 'large'
    }, theme);
    try {
        titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) { }

    if (settings.showTitleUnderline && title) {
        const uRect = safeGetRect(layout, `${key}.titleUnderline`);
        if (!uRect) {
            return;
        }
        let underlineWidthPt = 0;
        if (preCalculatedWidthPt !== null && preCalculatedWidthPt > 0) {
            underlineWidthPt = preCalculatedWidthPt;
        }
        else {
            underlineWidthPt = estimateTextWidthPt(title, initialFontSize);
        }
        const desiredWidthPt = underlineWidthPt + 10;
        const maxUnderlineWidth = layout.pageW_pt - uRect.left - layout.pxToPt(25);
        const finalWidth = Math.min(desiredWidthPt, maxUnderlineWidth);
        createPillShapeUnderline(slide, uRect.left, uRect.top, finalWidth, uRect.height, settings);
    }
}

export function drawSubheadIfAny(slide: GoogleAppsScript.Slides.Slide, layout: any, key: string, subhead: string, preCalculatedWidthPt: number | null = null) {
    if (!subhead) return 0;
    const theme = getThemeFromLayout(layout);
    const rect = safeGetRect(layout, `${key}.subhead`);
    if (!rect) {
        return 0;
    }
    const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left, rect.top, rect.width, rect.height);
    setStyledText(box, subhead, {
        size: theme.fonts.sizes.subhead,
        fontType: 'large'
    }, theme);
    return layout.pxToPt(36);
}

export function createContentCushion(slide: GoogleAppsScript.Slides.Slide, area: any, settings: any, layout: any) {
    const theme = getThemeFromLayout(layout);
    const cushion = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, area.width, area.height);
    cushion.getFill().setSolidFill(theme.colors.cardBg);
    try {
        const border = cushion.getBorder();
        border.setTransparent();
    } catch (e) { }
}

export function offsetRect(rect: any, dx: number, dy: number) {
    return {
        left: rect.left + dx,
        top: rect.top + dy,
        width: rect.width,
        height: rect.height
    };
}

export function adjustAreaForSubhead(area: any, subhead: string, layout: any) {
    return area;
}

export function setBoldTextSize(shapeOrTextRange: GoogleAppsScript.Slides.Shape | GoogleAppsScript.Slides.TextRange, size = 16) {
    let textRange;
    try {
        // @ts-ignore
        if (shapeOrTextRange && typeof shapeOrTextRange.getText === 'function') {
            // @ts-ignore
            textRange = shapeOrTextRange.getText();
        } else if (shapeOrTextRange && typeof (shapeOrTextRange as any).getRuns === 'function') {
            textRange = shapeOrTextRange as GoogleAppsScript.Slides.TextRange;
        } else {
            return;
        }
        if (!textRange || textRange.isEmpty()) {
            return;
        }
        const runs = textRange.getRuns();
        runs.forEach(run => {
            const style = run.getTextStyle();
            if (style.isBold()) {
                style.setFontSize(size);
            }
        });
    } catch (e) {
    }
}


