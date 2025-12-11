import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { SlideTheme } from '../../../common/config/SlideTheme';
import { DEFAULT_THEME } from '../../../common/config/DefaultTheme';
import {
    setStyledText,
    offsetRect,
    addFooter,
    drawArrowBetweenRects,
    setBoldTextSize,
    insertImageFromUrlOrFileId
} from '../../../common/utils/SlideUtils';
import {
    generateProcessColors,
    generateTimelineCardColors,
    generatePyramidColors,
    generateCompareColors,
    generateTintedGray
} from '../../../common/utils/ColorUtils';

export class GasDiagramSlideGenerator implements ISlideGenerator {
    constructor(private creditImageBlob: GoogleAppsScript.Base.BlobSource | null) { }

    generate(slide: GoogleAppsScript.Slides.Slide, data: any, layout: LayoutManager, pageNum: number, settings: any, imageUpdateOption: string = 'update') {
        Logger.log(`Generating Diagram Slide: ${data.layout || data.type}`);

        // Set Title
        const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
        if (titlePlaceholder) {
            try {
                titlePlaceholder.asShape().getText().setText(data.title || '');
            } catch (e) {
                Logger.log(`Warning: Title placeholder found but text could not be set. ${e}`);
            }
        }

        // Priority: JSON type field > layout field
        const type = (data.type || data.layout || '').toLowerCase();
        Logger.log('Generating Diagram Slide: ' + type);

        // Identify the target placeholder to use as the drawing canvas
        // Priority: BODY -> OBJECT -> PICTURE
        // This allows users to define the "Diagram Area" using standard placeholders in their Master slides.
        const placeholders = slide.getPlaceholders();

        // Helper to safely get placeholder type from PageElement
        const getPlaceholderTypeSafe = (p: GoogleAppsScript.Slides.PageElement): GoogleAppsScript.Slides.PlaceholderType | null => {
            try {
                const shape = p.asShape();
                if (shape) {
                    return shape.getPlaceholderType();
                }
            } catch (e) {
                // Not a shape or error getting type
            }
            return null;
        };

        const targetPlaceholder = placeholders.find(p => getPlaceholderTypeSafe(p) === SlidesApp.PlaceholderType.BODY)
            || placeholders.find(p => getPlaceholderTypeSafe(p) === SlidesApp.PlaceholderType.OBJECT)
            || placeholders.find(p => getPlaceholderTypeSafe(p) === SlidesApp.PlaceholderType.PICTURE);

        const workArea = targetPlaceholder ?
            { left: targetPlaceholder.getLeft(), top: targetPlaceholder.getTop(), width: targetPlaceholder.getWidth(), height: targetPlaceholder.getHeight() } :
            layout.getRect('contentSlide.body');

        // Remove the target placeholder to clear the stage for the diagram
        if (targetPlaceholder) {
            try {
                targetPlaceholder.remove();
            } catch (e) {
                Logger.log('Warning: Could not remove target placeholder: ' + e);
            }
        }

        // Get elements before drawing (to identify new ones later)
        const elementsBefore = slide.getPageElements().map(e => e.getObjectId());

        // switch logic for diagram types - wrapped in try-catch for robustness
        try {
            if (type.includes('timeline')) {
                this.drawTimeline(slide, data, workArea, settings, layout);
            } else if (type.includes('process')) {
                this.drawProcess(slide, data, workArea, settings, layout);
            } else if (type.includes('cycle')) {
                this.drawCycle(slide, data, workArea, settings, layout);
            } else if (type.includes('pyramid')) {
                this.drawPyramid(slide, data, workArea, settings, layout);
            } else if (type.includes('triangle')) {
                this.drawTriangle(slide, data, workArea, settings, layout);
            } else if (type.includes('statscompare')) {
                this.drawStatsCompare(slide, data, workArea, settings, layout);
            } else if (type.includes('barcompare')) {
                this.drawBarCompare(slide, data, workArea, settings, layout);
            } else if (type.includes('compare') || type.includes('kaizen')) {
                this.drawComparison(slide, data, workArea, settings, layout);
            } else if (type.includes('stepup') || type.includes('stair')) {
                this.drawStepUp(slide, data, workArea, settings, layout);
            } else if (type.includes('flowchart')) {
                this.drawFlowChart(slide, data, workArea, settings, layout);
            } else if (type.includes('diagram')) { // Lane diagram
                this.drawLanes(slide, data, workArea, settings, layout);
            } else if (type.includes('cards') || type.includes('headercards')) {
                this.drawCards(slide, data, workArea, settings, layout);
            } else if (type.includes('kpi')) {
                this.drawKPI(slide, data, workArea, settings, layout);
            } else if (type.includes('table')) {
                this.drawTable(slide, data, workArea, settings, layout);
            } else if (type.includes('faq')) {
                this.drawFAQ(slide, data, workArea, settings, layout);
            } else if (type.includes('progress')) {
                this.drawProgress(slide, data, workArea, settings, layout);
            } else if (type.includes('quote')) {
                this.drawQuote(slide, data, workArea, settings, layout);
            } else if (type.includes('imagetext')) {
                this.drawImageText(slide, data, workArea, settings, layout);
            } else {
                // Fallback
                Logger.log('Diagram logic not implemented for type: ' + type);
            }
        } catch (e) {
            Logger.log(`ERROR in drawing ${type}: ${e}`);
        }

        // Get new elements created during drawing and group them
        // Exclude placeholders (title, subtitle) - only group content/diagram elements
        const newElements = slide.getPageElements().filter(e => {
            // Must be a new element (not existing before drawing)
            if (elementsBefore.includes(e.getObjectId())) return false;

            // Exclude placeholder shapes (title, subtitle)
            try {
                const shape = e.asShape();
                if (shape) {
                    const placeholderType = shape.getPlaceholderType();
                    if (placeholderType === SlidesApp.PlaceholderType.TITLE ||
                        placeholderType === SlidesApp.PlaceholderType.SUBTITLE ||
                        placeholderType === SlidesApp.PlaceholderType.CENTERED_TITLE) {
                        return false;
                    }
                }
            } catch (e) {
                // Not a shape or can't determine placeholder type - include it
            }
            return true;
        });

        let generatedGroup: GoogleAppsScript.Slides.PageElement | null = null;

        if (newElements.length > 1) {
            try {
                generatedGroup = slide.group(newElements) as any;
                Logger.log(`Grouped ${newElements.length} content elements for ${type}`);
            } catch (e) {
                Logger.log(`Warning: Could not group elements: ${e}`);
            }
        } else if (newElements.length === 1) {
            generatedGroup = newElements[0];
        }

        // Center the generated content within the work area
        if (generatedGroup) {
            const currentWidth = generatedGroup.getWidth();
            const currentHeight = generatedGroup.getHeight();

            const centerX = workArea.left + (workArea.width - currentWidth) / 2;
            const centerY = workArea.top + (workArea.height - currentHeight) / 2;

            generatedGroup.setLeft(centerX);
            generatedGroup.setTop(centerY);
        }

        addFooter(slide, layout, pageNum, settings, this.creditImageBlob);
    }

    private drawTimeline(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const milestones = data.milestones || data.items || [];
        if (!milestones.length) return;

        const inner = layout.pxToPt(80),
            baseY = area.top + area.height * 0.50;
        const leftX = area.left + inner,
            rightX = area.left + area.width - inner;

        // Draw Line
        const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, leftX, baseY, rightX, baseY);
        line.getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
        line.setWeight(2);

        const dotR = layout.pxToPt(10);
        const gap = (milestones.length > 1) ? (rightX - leftX) / (milestones.length - 1) : 0;
        const cardW_pt = layout.pxToPt(180);
        const vOffset = layout.pxToPt(40);
        const headerHeight = layout.pxToPt(28);
        const bodyHeight = layout.pxToPt(80);

        // Colors
        const timelineColors = generateTimelineCardColors(settings.primaryColor, milestones.length);

        milestones.forEach((m: any, i: number) => {
            const x = leftX + gap * i;
            const isAbove = i % 2 === 0;
            const dateText = String(m.date || '');
            const labelText = String(m.label || m.state || '');

            const cardH_pt = headerHeight + bodyHeight;
            const cardLeft = x - (cardW_pt / 2);
            const cardTop = isAbove ? (baseY - vOffset - cardH_pt) : (baseY + vOffset);

            // Header
            const headerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cardLeft, cardTop, cardW_pt, headerHeight);
            headerShape.getFill().setSolidFill(timelineColors[i]);
            headerShape.getBorder().getLineFill().setSolidFill(timelineColors[i]);

            // Body
            const bodyShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cardLeft, cardTop + headerHeight, cardW_pt, bodyHeight);
            bodyShape.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
            bodyShape.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);

            // Connector
            const connectorY_start = isAbove ? (cardTop + cardH_pt) : baseY;
            const connectorY_end = isAbove ? baseY : cardTop;
            const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, connectorY_start, x, connectorY_end);
            connector.getLineFill().setSolidFill(DEFAULT_THEME.colors.neutralGray);
            connector.setWeight(1);

            // Dot
            const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - dotR / 2, baseY - dotR / 2, dotR, dotR);
            dot.getFill().setSolidFill(timelineColors[i]);
            dot.getBorder().setTransparent();

            // Header Text
            const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cardLeft, cardTop, cardW_pt, headerHeight);
            setStyledText(headerTextShape, dateText, {
                size: DEFAULT_THEME.fonts.sizes.body,
                bold: true,
                color: DEFAULT_THEME.colors.backgroundGray,
                align: SlidesApp.ParagraphAlignment.CENTER
            });
            try {
                headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) { }

            // Body Text
            let bodyFontSize = DEFAULT_THEME.fonts.sizes.body;
            const textLength = labelText.length;
            if (textLength > 40) bodyFontSize = 10;
            else if (textLength > 30) bodyFontSize = 11;
            else if (textLength > 20) bodyFontSize = 12;

            const bodyTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cardLeft, cardTop + headerHeight, cardW_pt, bodyHeight);
            setStyledText(bodyTextShape, labelText, {
                size: bodyFontSize,
                align: SlidesApp.ParagraphAlignment.CENTER
            });
            try {
                bodyTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) { }
        });
    }

    private drawProcess(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const steps = data.steps || data.items || [];
        if (!steps.length) return;

        const n = steps.length;
        let boxHPx, arrowHPx, fontSize;
        if (n <= 2) {
            boxHPx = 100; arrowHPx = 25; fontSize = 16;
        } else if (n === 3) {
            boxHPx = 80; arrowHPx = 20; fontSize = 16;
        } else {
            boxHPx = 65; arrowHPx = 15; fontSize = 14;
        }

        const processColors = generateProcessColors(settings.primaryColor, n);
        const startY = area.top + layout.pxToPt(10);
        let currentY = startY;
        const boxHPt = layout.pxToPt(boxHPx),
            arrowHPt = layout.pxToPt(arrowHPx);
        const headerWPt = layout.pxToPt(120);
        const bodyLeft = area.left + headerWPt;
        const bodyWPt = area.width - headerWPt;

        for (let i = 0; i < n; i++) {
            const cleanText = String(steps[i] || '').replace(/^\s*\d+[\.\s]*/, '');

            // Header
            const header = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, currentY, headerWPt, boxHPt);
            header.getFill().setSolidFill(processColors[i]);
            header.getBorder().setTransparent();
            setStyledText(header, `STEP ${i + 1}`, {
                size: fontSize,
                bold: true,
                color: DEFAULT_THEME.colors.backgroundGray,
                align: SlidesApp.ParagraphAlignment.CENTER
            });
            try { header.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Body
            const body = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, bodyLeft, currentY, bodyWPt, boxHPt);
            body.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
            body.getBorder().setTransparent();

            // Body Text
            const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, bodyLeft + layout.pxToPt(20), currentY, bodyWPt - layout.pxToPt(40), boxHPt);
            setStyledText(textShape, cleanText, { size: fontSize });
            try { textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            currentY += boxHPt;
            if (i < n - 1) {
                const arrowLeft = area.left + headerWPt / 2 - layout.pxToPt(8);
                const arrow = slide.insertShape(SlidesApp.ShapeType.DOWN_ARROW, arrowLeft, currentY, layout.pxToPt(16), arrowHPt);
                arrow.getFill().setSolidFill(DEFAULT_THEME.colors.processArrow || DEFAULT_THEME.colors.ghostGray);
                arrow.getBorder().setTransparent();
                currentY += arrowHPt;
            }
        }
    }

    private drawCycle(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const items = data.items || [];
        if (!items.length) return;

        const textLengths = items.map((item: any) => {
            const labelLength = (item.label || '').length;
            const subLabelLength = (item.subLabel || '').length;
            return labelLength + subLabelLength;
        });
        const maxLength = Math.max(...textLengths);
        const avgLength = textLengths.reduce((sum: number, len: number) => sum + len, 0) / textLengths.length;

        const centerX = area.left + area.width / 2;
        const centerY = area.top + area.height / 2;
        const radiusX = area.width / 3.2;
        const radiusY = area.height / 2.6;

        const maxCardW = Math.min(layout.pxToPt(220), radiusX * 0.8);
        const maxCardH = Math.min(layout.pxToPt(100), radiusY * 0.6);

        let cardW, cardH, fontSize;
        if (maxLength > 25 || avgLength > 18) {
            cardW = Math.min(layout.pxToPt(230), maxCardW); cardH = Math.min(layout.pxToPt(105), maxCardH); fontSize = 13;
        } else if (maxLength > 15 || avgLength > 10) {
            cardW = Math.min(layout.pxToPt(215), maxCardW); cardH = Math.min(layout.pxToPt(95), maxCardH); fontSize = 14;
        } else {
            cardW = layout.pxToPt(200); cardH = layout.pxToPt(90); fontSize = 16;
        }

        if (data.centerText) {
            const centerTextBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, centerX - layout.pxToPt(100), centerY - layout.pxToPt(50), layout.pxToPt(200), layout.pxToPt(100));
            setStyledText(centerTextBox, data.centerText, { size: 20, bold: true, align: SlidesApp.ParagraphAlignment.CENTER, color: DEFAULT_THEME.colors.textPrimary });
            try { centerTextBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
        }

        const positions = [
            { x: centerX + radiusX, y: centerY },
            { x: centerX, y: centerY + radiusY },
            { x: centerX - radiusX, y: centerY },
            { x: centerX, y: centerY - radiusY }
        ];

        // Ensure we only draw as many items as we have (up to 4)
        const itemsToDraw = items.slice(0, 4);

        itemsToDraw.forEach((item: any, i: number) => {
            const pos = positions[i];
            const cardX = pos.x - cardW / 2;
            const cardY = pos.y - cardH / 2;

            const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardW, cardH);
            card.getFill().setSolidFill(settings.primaryColor);
            card.getBorder().setTransparent();

            const subLabelText = item.subLabel || `${i + 1}番目`;
            const labelText = item.label || '';
            setStyledText(card, `${subLabelText}\n${labelText}`, { size: fontSize, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });

            try {
                card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
                const textRange = card.getText();
                const subLabelEnd = subLabelText.length;
                if (textRange.asString().length > subLabelEnd) {
                    textRange.getRange(0, subLabelEnd).getTextStyle().setFontSize(Math.max(10, fontSize - 2));
                }
            } catch (e) { }
        });

        // Bent Arrows
        const arrowRadiusX = radiusX * 0.75;
        const arrowRadiusY = radiusY * 0.80;
        const arrowSize = layout.pxToPt(80);
        const arrowPositions = [
            { left: centerX + arrowRadiusX, top: centerY - arrowRadiusY, rotation: 90 },
            { left: centerX + arrowRadiusX, top: centerY + arrowRadiusY, rotation: 180 },
            { left: centerX - arrowRadiusX, top: centerY + arrowRadiusY, rotation: 270 },
            { left: centerX - arrowRadiusX, top: centerY - arrowRadiusY, rotation: 0 }
        ];

        arrowPositions.slice(0, itemsToDraw.length).forEach(pos => {
            const arrow = slide.insertShape(SlidesApp.ShapeType.BENT_ARROW, pos.left - arrowSize / 2, pos.top - arrowSize / 2, arrowSize, arrowSize);
            arrow.getFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
            arrow.getBorder().setTransparent();
            arrow.setRotation(pos.rotation);
        });
    }

    private drawPyramid(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const levels = data.levels || data.items || [];
        if (!levels.length) return;

        // Limit to 4 for visual consistency as per reference, or relax? Reference says slice(0, 4)
        const levelsToDraw = levels.slice(0, 4);

        const levelHeight = layout.pxToPt(70);
        const levelGap = layout.pxToPt(2);
        const totalHeight = (levelHeight * levelsToDraw.length) + (levelGap * (levelsToDraw.length - 1));
        const startY = area.top + (area.height - totalHeight) / 2;
        const pyramidWidth = layout.pxToPt(480);
        const textColumnWidth = layout.pxToPt(400);
        const gap = layout.pxToPt(30);
        const pyramidLeft = area.left;
        const textColumnLeft = pyramidLeft + pyramidWidth + gap;

        const pyramidColors = generatePyramidColors(settings.primaryColor, levelsToDraw.length);
        const baseWidth = pyramidWidth;
        const widthIncrement = baseWidth / levelsToDraw.length;
        const centerX = pyramidLeft + pyramidWidth / 2;

        levelsToDraw.forEach((level: any, index: number) => {
            const levelWidth = baseWidth - (widthIncrement * (levelsToDraw.length - 1 - index));
            const levelX = centerX - levelWidth / 2;
            const levelY = startY + index * (levelHeight + levelGap);

            // Shape
            const levelBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, levelX, levelY, levelWidth, levelHeight);
            levelBox.getFill().setSolidFill(pyramidColors[index]);
            levelBox.getBorder().setTransparent();

            // Inner Text (Title)
            const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, levelX, levelY, levelWidth, levelHeight);
            titleShape.getFill().setTransparent();
            titleShape.getBorder().setTransparent();
            const levelTitle = level.title || `レベル${index + 1}`;
            setStyledText(titleShape, levelTitle, {
                size: DEFAULT_THEME.fonts.sizes.body,
                bold: true,
                color: DEFAULT_THEME.colors.backgroundGray,
                align: SlidesApp.ParagraphAlignment.CENTER
            });
            try { titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Connecting Line
            const connectionStartX = levelX + levelWidth;
            const connectionEndX = textColumnLeft;
            const connectionY = levelY + levelHeight / 2;
            if (connectionEndX > connectionStartX) {
                const connectionLine = slide.insertLine(
                    SlidesApp.LineCategory.STRAIGHT,
                    connectionStartX, connectionY, connectionEndX, connectionY
                );
                connectionLine.getLineFill().setSolidFill('#D0D7DE');
                connectionLine.setWeight(1.5);
            }

            // Description Text
            const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textColumnLeft, levelY, textColumnWidth, levelHeight);
            textShape.getFill().setTransparent();
            textShape.getBorder().setTransparent();
            const levelDesc = level.description || '';
            let formattedText;
            if (levelDesc.includes('•') || levelDesc.includes('・')) {
                formattedText = levelDesc;
            } else if (levelDesc.includes('\n')) {
                formattedText = levelDesc.split('\n').filter((l: string) => l.trim()).slice(0, 2).map((l: string) => `• ${l.trim()}`).join('\n');
            } else {
                formattedText = levelDesc;
            }
            setStyledText(textShape, formattedText, {
                size: DEFAULT_THEME.fonts.sizes.body - 1,
                align: SlidesApp.ParagraphAlignment.START, // Fixed from LEFT
                color: DEFAULT_THEME.colors.textPrimary,
                bold: true
            });
            try { textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
        });
    }

    private drawComparison(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const leftTitle = data.leftTitle || 'プランA';
        const rightTitle = data.rightTitle || 'プランB';
        const leftItems = data.leftItems || [];
        const rightItems = data.rightItems || [];

        const gap = 20;
        const colWidth = (area.width - gap) / 2;

        const compareColors = generateCompareColors(settings.primaryColor);
        const headerH = layout.pxToPt(40);

        // Left
        const leftHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, colWidth, headerH);
        leftHeader.getFill().setSolidFill(compareColors.left);
        leftHeader.getBorder().setTransparent();
        setStyledText(leftHeader, leftTitle, { size: 14, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });
        try { leftHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        const leftBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top + headerH, colWidth, area.height - headerH);
        leftBox.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray); // Use light gray instead of F0F0F0 for consistency
        leftBox.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);
        setStyledText(leftBox, leftItems.join('\n\n'), { size: DEFAULT_THEME.fonts.sizes.body });

        // Right
        const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left + colWidth + gap, area.top, colWidth, headerH);
        rightHeader.getFill().setSolidFill(compareColors.right);
        rightHeader.getBorder().setTransparent();
        setStyledText(rightHeader, rightTitle, { size: 14, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });
        try { rightHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        const rightBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left + colWidth + gap, area.top + headerH, colWidth, area.height - headerH);
        rightBox.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
        rightBox.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);
        setStyledText(rightBox, rightItems.join('\n\n'), { size: DEFAULT_THEME.fonts.sizes.body });
    }

    private drawStatsCompare(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const leftTitle = data.leftTitle || '導入前';
        const rightTitle = data.rightTitle || '導入後';
        const stats = data.stats || [];
        if (!stats.length) return;

        const compareColors = generateCompareColors(settings.primaryColor);

        // Header row
        const headerH = layout.pxToPt(45);
        const labelColW = area.width * 0.35;  // Label column width
        const valueColW = (area.width - labelColW) / 2;  // Each value column width

        // Left Title Header
        const leftHeaderX = area.left + labelColW;
        const leftHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftHeaderX, area.top, valueColW, headerH);
        leftHeader.getFill().setSolidFill(compareColors.left);
        leftHeader.getBorder().setTransparent();
        setStyledText(leftHeader, leftTitle, { size: 14, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });
        try { leftHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        // Right Title Header
        const rightHeaderX = area.left + labelColW + valueColW;
        const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightHeaderX, area.top, valueColW, headerH);
        rightHeader.getFill().setSolidFill(compareColors.right);
        rightHeader.getBorder().setTransparent();
        setStyledText(rightHeader, rightTitle, { size: 14, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });
        try { rightHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        // Data rows
        const availableHeight = area.height - headerH;
        const rowHeight = Math.min(layout.pxToPt(60), availableHeight / stats.length);
        let currentY = area.top + headerH;

        stats.forEach((stat: any, index: number) => {
            const label = stat.label || '';
            const leftValue = stat.leftValue || '';
            const rightValue = stat.rightValue || '';
            const trend = stat.trend || null;

            // Alternate row background
            const rowBg = index % 2 === 0 ? DEFAULT_THEME.colors.backgroundGray : '#FFFFFF';

            // Label cell
            const labelCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, currentY, labelColW, rowHeight);
            labelCell.getFill().setSolidFill(rowBg);
            labelCell.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
            setStyledText(labelCell, label, { size: DEFAULT_THEME.fonts.sizes.body, bold: true, align: SlidesApp.ParagraphAlignment.START });
            try { labelCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Left value cell
            const leftCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftHeaderX, currentY, valueColW, rowHeight);
            leftCell.getFill().setSolidFill(rowBg);
            leftCell.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
            setStyledText(leftCell, leftValue, { size: DEFAULT_THEME.fonts.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER, color: compareColors.left });
            try { leftCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Right value cell
            const rightCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightHeaderX, currentY, valueColW - (trend ? layout.pxToPt(40) : 0), rowHeight);
            rightCell.getFill().setSolidFill(rowBg);
            rightCell.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
            setStyledText(rightCell, rightValue, { size: DEFAULT_THEME.fonts.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER, color: compareColors.right });
            try { rightCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Trend indicator (optional)
            if (trend) {
                const trendX = rightHeaderX + valueColW - layout.pxToPt(35);
                const trendShape = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, trendX, currentY + rowHeight / 4, layout.pxToPt(25), layout.pxToPt(25));
                const isUp = trend.toLowerCase() === 'up';
                const trendColor = isUp ? '#28a745' : '#dc3545';  // Green for up, Red for down
                trendShape.getFill().setSolidFill(trendColor);
                trendShape.getBorder().setTransparent();
                setStyledText(trendShape, isUp ? '↑' : '↓', { size: 12, color: '#FFFFFF', bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
                try { trendShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
            }

            currentY += rowHeight;
        });
    }

    private drawBarCompare(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const leftTitle = data.leftTitle || '導入前';
        const rightTitle = data.rightTitle || '導入後';
        const stats = data.stats || [];
        if (!stats.length) return;

        const compareColors = generateCompareColors(settings.primaryColor);

        // Find max value for scaling
        let maxValue = 0;
        stats.forEach((stat: any) => {
            const leftNum = parseFloat(String(stat.leftValue || '0').replace(/[^0-9.]/g, '')) || 0;
            const rightNum = parseFloat(String(stat.rightValue || '0').replace(/[^0-9.]/g, '')) || 0;
            maxValue = Math.max(maxValue, leftNum, rightNum);
        });
        if (maxValue === 0) maxValue = 100;  // Fallback

        // Layout
        const labelColW = area.width * 0.2;
        const barAreaW = area.width * 0.6;
        const valueColW = area.width * 0.1;
        const trendColW = area.width * 0.1;

        const rowHeight = Math.min(layout.pxToPt(80), area.height / stats.length);
        const barHeight = layout.pxToPt(18);
        const barGap = layout.pxToPt(4);
        let currentY = area.top;

        stats.forEach((stat: any, index: number) => {
            const label = stat.label || '';
            const leftValue = stat.leftValue || '';
            const rightValue = stat.rightValue || '';
            const trend = stat.trend || null;

            const leftNum = parseFloat(String(leftValue).replace(/[^0-9.]/g, '')) || 0;
            const rightNum = parseFloat(String(rightValue).replace(/[^0-9.]/g, '')) || 0;

            // Label
            const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, currentY, labelColW, rowHeight);
            setStyledText(labelShape, label, { size: DEFAULT_THEME.fonts.sizes.body, bold: true, align: SlidesApp.ParagraphAlignment.START });
            try { labelShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Bar area
            const barLeft = area.left + labelColW;
            const barTop = currentY + (rowHeight - (barHeight * 2 + barGap)) / 2;

            // Left bar (Before)
            const leftBarWidth = (leftNum / maxValue) * barAreaW;
            if (leftBarWidth > 0) {
                const leftBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barTop, leftBarWidth, barHeight);
                leftBar.getFill().setSolidFill(compareColors.left);
                leftBar.getBorder().setTransparent();
            }
            // Left label
            const leftLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + leftBarWidth + layout.pxToPt(5), barTop, layout.pxToPt(60), barHeight);
            setStyledText(leftLabel, leftValue, { size: 10, color: compareColors.left });

            // Right bar (After)
            const rightBarWidth = (rightNum / maxValue) * barAreaW;
            if (rightBarWidth > 0) {
                const rightBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barTop + barHeight + barGap, rightBarWidth, barHeight);
                rightBar.getFill().setSolidFill(compareColors.right);
                rightBar.getBorder().setTransparent();
            }
            // Right label
            const rightLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + rightBarWidth + layout.pxToPt(5), barTop + barHeight + barGap, layout.pxToPt(60), barHeight);
            setStyledText(rightLabel, rightValue, { size: 10, color: compareColors.right });

            // Trend indicator (optional)
            if (trend) {
                const trendX = area.left + area.width - trendColW;
                const trendShape = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, trendX, currentY + rowHeight / 2 - layout.pxToPt(12), layout.pxToPt(24), layout.pxToPt(24));
                const isUp = trend.toLowerCase() === 'up';
                const trendColor = isUp ? '#28a745' : '#dc3545';
                trendShape.getFill().setSolidFill(trendColor);
                trendShape.getBorder().setTransparent();
                setStyledText(trendShape, isUp ? '↑' : '↓', { size: 12, color: '#FFFFFF', bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
                try { trendShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
            }

            // Separator line
            if (index < stats.length - 1) {
                const lineY = currentY + rowHeight;
                const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, area.left, lineY, area.left + area.width, lineY);
                line.getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
                line.setWeight(1);
            }

            currentY += rowHeight;
        });
    }

    private drawTriangle(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const items = data.items || [];
        if (!items.length) return;

        const itemsToDraw = items.slice(0, 3);

        const centerX = area.left + area.width / 2;
        const centerY = area.top + area.height / 2;
        // 少し小さめに
        const radius = Math.min(area.width, area.height) / 3.2;

        // 三角形の頂点配置 (上、右下、左下)
        const positions = [
            { x: centerX, y: centerY - radius },
            { x: centerX + radius * 0.866, y: centerY + radius * 0.5 },
            { x: centerX - radius * 0.866, y: centerY + radius * 0.5 }
        ];

        // 各要素の円のサイズ
        const circleSize = layout.pxToPt(160); // 直径

        // 中央の三角形（飾り）
        const trianglePath = slide.insertShape(SlidesApp.ShapeType.TRIANGLE, centerX - radius / 2, centerY - radius / 2, radius, radius);
        trianglePath.setRotation(0); // 正三角形の向きに
        // 実際にはシェイプの頂点は矩形内配置なので微調整が必要だが、一旦シンプルに
        // 背景の三角形は薄く
        trianglePath.getFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
        trianglePath.getBorder().setTransparent();

        // Send manually to back (GAS doesn't have z-index easily, relying on insertion order - insert first = back)
        // Adjust: insert shape first

        positions.slice(0, itemsToDraw.length).forEach((pos, i) => {
            const item = itemsToDraw[i];
            const title = item.title || item.label || '';
            const desc = item.desc || item.subLabel || '';

            const x = pos.x - circleSize / 2;
            const y = pos.y - circleSize / 2;

            // Circle Shape
            const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, circleSize, circleSize);
            circle.getFill().setSolidFill(settings.primaryColor);
            circle.getBorder().setTransparent();

            // Text
            setStyledText(circle, `${title}\n${desc}`, {
                size: 14,
                bold: true,
                color: DEFAULT_THEME.colors.backgroundGray,
                align: SlidesApp.ParagraphAlignment.CENTER
            });
            try {
                circle.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
                // タイトルと説明のフォントサイズを変える
                const textRange = circle.getText();
                if (title.length > 0 && desc.length > 0) {
                    // 1行目がタイトルと仮定
                    // 改行位置を探すのは手間なので一律サイズ設定でシンプルに
                }
            } catch (e) { }
        });
    }

    private drawStepUp(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const items = data.items || [];
        if (!items.length) return;
        const count = items.length;
        // Logic from reference createStepUpSlide - seems it was stubbed but reference has implementation?
        // Actually reference implementation of createStepUpSlide (not fully pasted in previous turn but inferred):
        // It uses stairs logic.

        const stepWidth = area.width / count;
        const stepHeight = area.height / count;

        items.forEach((item: any, i: number) => {
            const h = (i + 1) * stepHeight;
            const x = area.left + (i * stepWidth);
            const y = area.top + (area.height - h);

            const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, stepWidth - layout.pxToPt(5), h);
            // Use opacity/alpha if possible or gradients
            // GAS setSolidFill support overload? setSolidFill(color, alpha)
            shape.getFill().setSolidFill(settings.primaryColor, 0.5 + (i * 0.1));
            shape.getBorder().setTransparent();

            const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y, stepWidth - layout.pxToPt(5), layout.pxToPt(50));
            setStyledText(textBox, item.title || '', { color: '#FFFFFF', bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
        });
    }

    private drawLanes(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const lanes = data.lanes || [];
        const n = Math.max(1, lanes.length);
        const { laneGapPx, lanePadPx, laneTitleHeightPx, cardGapPx, cardMinHeightPx, cardMaxHeightPx, arrowHeightPx, arrowGapPx } = DEFAULT_THEME.diagram;
        const px = (p: number) => layout.pxToPt(p);

        const laneW = (area.width - px(laneGapPx) * (n - 1)) / n;
        const cardBoxes: any[] = [];

        for (let j = 0; j < n; j++) {
            const lane = lanes[j] || { title: '', items: [] };
            const left = area.left + j * (laneW + px(laneGapPx));

            // Lane Header
            const lt = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, area.top, laneW, px(laneTitleHeightPx));
            lt.getFill().setSolidFill(settings.primaryColor);
            lt.getBorder().getLineFill().setSolidFill(settings.primaryColor);
            setStyledText(lt, lane.title || '', { size: DEFAULT_THEME.fonts.sizes.laneTitle, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });
            try { lt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Lane Body
            const laneBodyTop = area.top + px(laneTitleHeightPx);
            const laneBodyHeight = area.height - px(laneTitleHeightPx);
            const laneBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, laneBodyTop, laneW, laneBodyHeight);
            laneBg.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
            laneBg.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.laneBorder);

            const items = Array.isArray(lane.items) ? lane.items : [];
            const rows = Math.max(1, items.length);
            const availH = laneBodyHeight - px(lanePadPx) * 2;
            const idealH = (availH - px(cardGapPx) * (rows - 1)) / rows;
            const cardH = Math.max(px(cardMinHeightPx), Math.min(px(cardMaxHeightPx), idealH));
            const firstTop = laneBodyTop + px(lanePadPx) + Math.max(0, (availH - (cardH * rows + px(cardGapPx) * (rows - 1))) / 2);

            cardBoxes[j] = [];
            for (let i = 0; i < rows; i++) {
                const cardTop = firstTop + i * (cardH + px(cardGapPx));
                const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, left + px(lanePadPx), cardTop, laneW - px(lanePadPx) * 2, cardH);
                card.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
                card.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);
                setStyledText(card, items[i] || '', { size: DEFAULT_THEME.fonts.sizes.body });
                try { card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

                cardBoxes[j][i] = {
                    left: left + px(lanePadPx),
                    top: cardTop,
                    width: laneW - px(lanePadPx) * 2,
                    height: cardH
                };
            }
        }

        // Draw Arrows
        const maxRows = Math.max(0, ...cardBoxes.map(a => a ? a.length : 0));
        for (let j = 0; j < n - 1; j++) {
            for (let i = 0; i < maxRows; i++) {
                if (cardBoxes[j] && cardBoxes[j][i] && cardBoxes[j + 1] && cardBoxes[j + 1][i]) {
                    drawArrowBetweenRects(slide, cardBoxes[j][i], cardBoxes[j + 1][i], px(arrowHeightPx), px(arrowGapPx), settings);
                }
            }
        }
    }

    private drawFlowChart(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        // TBD: Logic for FlowChart based on available data structure
        // If simple linear flow:
        const steps = data.steps || data.items || [];
        if (!steps.length) return;

        // Similar to Process but maybe boxes with arrows connected
        // Using a simple left-to-right flow for now
        const count = steps.length;
        const gap = 30;
        const boxWidth = (area.width - (gap * (count - 1))) / count;
        const boxHeight = 80;
        const y = area.top + (area.height - boxHeight) / 2;

        steps.forEach((step: any, i: number) => {
            const x = area.left + i * (boxWidth + gap);
            const shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, boxWidth, boxHeight);
            shape.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
            shape.getBorder().getLineFill().setSolidFill(settings.primaryColor);
            shape.getBorder().setWeight(2);
            setStyledText(shape, typeof step === 'string' ? step : step.label || '', { size: DEFAULT_THEME.fonts.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER });
            try { shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            if (i < count - 1) {
                // Arrow
                const ax = x + boxWidth;
                const ay = y + boxHeight / 2;
                const bx = x + boxWidth + gap;
                const by = ay;
                const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, ax, ay, bx, by);
                line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
                line.getLineFill().setSolidFill(DEFAULT_THEME.colors.neutralGray);
            }
        });
    }

    private drawCards(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const items = data.items || [];
        if (!items.length) return;
        const type = (data.type || '').toLowerCase();
        const hasHeader = type.includes('headercards');

        const cols = data.columns || 3;
        const rows = Math.ceil(items.length / cols);

        const gap = layout.pxToPt(20);
        const cardW = (area.width - (gap * (cols - 1))) / cols;
        const cardH = (area.height - (gap * (rows - 1))) / rows;

        items.forEach((item: any, i: number) => {
            const r = Math.floor(i / cols);
            const c = i % cols;
            const x = area.left + c * (cardW + gap);
            const y = area.top + r * (cardH + gap);

            // Handle both string items and object items
            let title = '';
            let desc = '';
            if (typeof item === 'string') {
                // Parse "Title\nDescription" or "Title: Description\nMore..." format
                const lines = item.split('\n');
                title = lines[0] || '';
                desc = lines.slice(1).join('\n') || '';
            } else {
                title = item.title || item.label || '';
                desc = item.desc || item.description || item.text || '';
            }

            if (hasHeader) {
                // Headered Card
                const headerH = layout.pxToPt(36);
                const header = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, cardW, headerH);
                header.getFill().setSolidFill(settings.primaryColor);
                header.getBorder().setTransparent();
                setStyledText(header, title, { size: 14, bold: true, color: '#FFFFFF', align: SlidesApp.ParagraphAlignment.CENTER });
                try { header.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

                const body = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y + headerH, cardW, cardH - headerH);
                body.getFill().setSolidFill('#FFFFFF');
                body.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);

                const textArea = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + layout.pxToPt(10), y + headerH + layout.pxToPt(10), cardW - layout.pxToPt(20), cardH - headerH - layout.pxToPt(20));
                setStyledText(textArea, desc, { size: 12, align: SlidesApp.ParagraphAlignment.START, color: DEFAULT_THEME.colors.textSmallFont });
            } else {
                // Modern Clean Card
                const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, cardW, cardH);
                card.getFill().setSolidFill('#FFFFFF');
                card.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);

                // Accent Strip on Left
                const strip = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y + cardH * 0.1, layout.pxToPt(6), cardH * 0.8);
                strip.getFill().setSolidFill(settings.primaryColor);
                strip.getBorder().setTransparent();

                // Content
                const contentX = x + layout.pxToPt(20);
                const contentW = cardW - layout.pxToPt(30);

                // Title
                const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y + layout.pxToPt(10), contentW, layout.pxToPt(30));
                setStyledText(titleBox, title, { size: 16, bold: true, color: DEFAULT_THEME.colors.textPrimary });

                // Desc
                const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y + layout.pxToPt(40), contentW, cardH - layout.pxToPt(50));
                setStyledText(descBox, desc, { size: 12, color: DEFAULT_THEME.colors.textSmallFont });
                try {
                    descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
                } catch (e) { }
            }
        });
    }

    private drawKPI(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const items = data.items || [];
        if (!items.length) return;

        const cols = items.length > 4 ? 4 : (items.length || 1);
        const gap = layout.pxToPt(20);
        const cardW = (area.width - (gap * (cols - 1))) / cols;
        const cardH = layout.pxToPt(160);
        const y = area.top + (area.height - cardH) / 2;

        items.forEach((item: any, i: number) => {
            const x = area.left + i * (cardW + gap);

            // Card BG
            const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, cardW, cardH);
            card.getFill().setSolidFill('#FFFFFF');
            card.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);
            // card.setShadow equivalent not avail, but we can add a bottom lip if we want depth?
            // Simulating shadow with offset shape underneath is possible but maybe too heavy. 
            // Stick to clean flat design.

            const padding = layout.pxToPt(10);

            // Label (Top)
            const labelH = layout.pxToPt(30);
            const labelBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, y + padding, cardW - padding * 2, labelH);
            setStyledText(labelBox, item.label || 'Metric', { size: 14, color: DEFAULT_THEME.colors.neutralGray, align: SlidesApp.ParagraphAlignment.CENTER });

            // Value (Middle - Large)
            const valueH = layout.pxToPt(70);
            const valueBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, y + labelH + padding, cardW - padding * 2, valueH);
            const valStr = String(item.value || '0');
            // Adaptive font size
            let fontSize = 48;
            if (valStr.length > 4) fontSize = 36;
            if (valStr.length > 6) fontSize = 28;
            if (valStr.length > 10) fontSize = 24;

            setStyledText(valueBox, valStr, { size: fontSize, bold: true, color: settings.primaryColor, align: SlidesApp.ParagraphAlignment.CENTER });

            // Change/Status (Bottom)
            if (item.change || item.status) {
                const statusH = layout.pxToPt(30);
                const statusBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, y + labelH + valueH + padding, cardW - padding * 2, statusH);

                let color = DEFAULT_THEME.colors.neutralGray;
                let prefix = '';
                if (item.status === 'good') { color = '#28a745'; prefix = '▲ '; }
                if (item.status === 'bad') { color = '#dc3545'; prefix = '▼ '; }

                setStyledText(statusBox, prefix + (item.change || ''), { size: 14, bold: true, color: color, align: SlidesApp.ParagraphAlignment.CENTER });
            }
        });
    }

    private drawTable(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const headers = data.headers || [];
        const rows = data.rows || [];
        const numRows = rows.length + (headers.length ? 1 : 0);
        const numCols = headers.length || (rows[0] ? rows[0].length : 1);
        if (numRows === 0 || numCols === 0) return;

        const table = slide.insertTable(numRows, numCols);
        table.setLeft(area.left);
        table.setTop(area.top);
        table.setWidth(area.width);

        let rowIndex = 0;
        // Header
        if (headers.length) {
            for (let c = 0; c < numCols; c++) {
                const cell = table.getCell(0, c);
                cell.getFill().setSolidFill(settings.primaryColor); // Header Color
                cell.getText().setText(headers[c] || '');
                const style = cell.getText().getTextStyle();
                style.setBold(true);
                style.setFontSize(14);
                style.setForegroundColor('#FFFFFF');
                try { cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
            }
            rowIndex++;
        }

        // Rows
        rows.forEach((row: any[], rIdx: number) => {
            const isAlt = rIdx % 2 !== 0; // Alternating
            const rowColor = isAlt ? DEFAULT_THEME.colors.faintGray : '#FFFFFF';

            for (let c = 0; c < numCols; c++) {
                const cell = table.getCell(rowIndex, c);
                cell.getFill().setSolidFill(rowColor);
                cell.getText().setText(String(row[c] || ''));
                const rowStyle = cell.getText().getTextStyle();
                rowStyle.setFontSize(12);
                rowStyle.setForegroundColor(DEFAULT_THEME.colors.textPrimary);
                try { cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
                // const borderColor = DEFAULT_THEME.colors.laneBorder;
                // cell.getBorderBottom().getLineFill().setSolidFill(borderColor);
                // cell.getBorderTop().getLineFill().setSolidFill(borderColor);
                // cell.getBorderLeft().setTransparent();
                // cell.getBorderRight().setTransparent();
            }
            rowIndex++;
        });
    }

    private drawFAQ(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const items = data.items || data.points || [];
        // Support points as strings ("Q:...", "A:...") or objects
        const parsedItems: any[] = [];
        if (items.length && typeof items[0] === 'string') {
            let currentQ = '';
            items.forEach((str: string) => {
                if (str.startsWith('Q:') || str.startsWith('Q.')) currentQ = str;
                else if (str.startsWith('A:') || str.startsWith('A.')) parsedItems.push({ q: currentQ, a: str });
            });
        } else {
            items.forEach((it: any) => parsedItems.push(it));
        }
        if (!parsedItems.length) return;

        const gap = layout.pxToPt(20);
        const itemH = (area.height - (gap * (parsedItems.length - 1))) / parsedItems.length;

        parsedItems.forEach((item, i) => {
            const y = area.top + i * (itemH + gap);
            const iconSize = layout.pxToPt(40);

            // Q Circle
            const qCircle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, area.left, y + (itemH - iconSize) / 2, iconSize, iconSize);
            qCircle.getFill().setSolidFill(settings.primaryColor);
            qCircle.getBorder().setTransparent();
            setStyledText(qCircle, 'Q', { size: 18, bold: true, color: '#FFFFFF', align: SlidesApp.ParagraphAlignment.CENTER });
            try { qCircle.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Content Box
            const boxLeft = area.left + iconSize + layout.pxToPt(15);
            const boxW = area.width - (iconSize + layout.pxToPt(15));
            const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, boxLeft, y, boxW, itemH);
            box.getFill().setSolidFill('#FFFFFF');
            box.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.cardBorder);

            const qText = (item.q || '').replace(/^[QA][:. ]+/, '');
            const aText = (item.a || '').replace(/^[QA][:. ]+/, '');

            setStyledText(box, `Q. ${qText}\n\nA. ${aText}`, { size: 12, color: DEFAULT_THEME.colors.textPrimary });

            // Style the A part? simple styling only.
            // A visual separator line inside?
            // Actually, let's keep it simple.
        });
    }

    private drawQuote(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const text = data.text || (data.points && data.points[0]) || '';
        const author = data.author || (data.points && data.points[1]) || '';

        // Background Cushion
        const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, area.width, area.height);
        bg.getFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
        bg.getBorder().setTransparent();

        // Big Quote Marks
        const quoteMark = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, area.top - layout.pxToPt(20), layout.pxToPt(100), layout.pxToPt(100));
        setStyledText(quoteMark, '“', { size: 120, color: DEFAULT_THEME.colors.ghostGray, font: 'Georgia' });

        const contentW = area.width * 0.8;
        const contentX = area.left + (area.width - contentW) / 2;

        const quoteBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, area.top, contentW, area.height - layout.pxToPt(60));
        setStyledText(quoteBox, text, { size: 28, bold: true, color: settings.primaryColor, align: SlidesApp.ParagraphAlignment.CENTER, font: 'Serif' }); // Usage of Serif if available, else standard
        try { quoteBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        if (author) {
            const authorBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, area.top + area.height - layout.pxToPt(60), contentW, layout.pxToPt(40));
            setStyledText(authorBox, `— ${author}`, { size: 16, align: SlidesApp.ParagraphAlignment.END, color: DEFAULT_THEME.colors.neutralGray });
        }
    }

    private drawProgress(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const items = data.items || [];
        if (!items.length) return;

        const rowH = layout.pxToPt(50);
        const gap = layout.pxToPt(15);
        const startY = area.top + (area.height - (items.length * (rowH + gap))) / 2;

        items.forEach((item: any, i: number) => {
            const y = startY + i * (rowH + gap);
            const labelW = layout.pxToPt(150);
            const barAreaW = area.width - labelW - layout.pxToPt(60);

            // Label
            const labelBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, y, labelW, rowH);
            setStyledText(labelBox, item.label || '', { size: 14, bold: true, align: SlidesApp.ParagraphAlignment.END });
            try { labelBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Bar BG
            const barBg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left + labelW + layout.pxToPt(20), y + rowH / 3, barAreaW, rowH / 3);
            barBg.getFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
            barBg.getBorder().setTransparent();

            // Bar FG
            const percent = Math.min(100, Math.max(0, parseInt(item.percent || 0)));
            if (percent > 0) {
                const barFg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left + labelW + layout.pxToPt(20), y + rowH / 3, barAreaW * (percent / 100), rowH / 3);
                barFg.getFill().setSolidFill(settings.primaryColor);
                barFg.getBorder().setTransparent();
            }

            // Value
            const valBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + labelW + barAreaW + layout.pxToPt(30), y, layout.pxToPt(50), rowH);
            setStyledText(valBox, `${percent}%`, { size: 14, color: DEFAULT_THEME.colors.neutralGray });
            try { valBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
        });
    }

    private drawImageText(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const imageUrl = data.image;
        const points = data.points || []; // Text content

        // Define areas
        const gap = layout.pxToPt(20);
        const halfW = (area.width - gap) / 2;

        const isImageLeft = data.imagePosition !== 'right';
        const imgX = isImageLeft ? area.left : area.left + halfW + gap;
        const txtX = isImageLeft ? area.left + halfW + gap : area.left;

        // Draw Image
        if (imageUrl) {
            // We use a helper from SlideUtils. But wait, existing code didn't import Utils fully.
            // I added standard imports at top.
            try {
                // using helper to get blob or null
                const blob = insertImageFromUrlOrFileId(imageUrl);
                let img;
                if (blob) {
                    img = slide.insertImage(blob);
                } else if (imageUrl.startsWith('http')) {
                    // Fallback to URL direct insert
                    img = slide.insertImage(imageUrl);
                }

                if (img) {
                    img.setLeft(imgX).setTop(area.top).setWidth(halfW).setHeight(area.height);
                    // crop or Aspect fit? 
                    // Simple aspect fit in center of rect usually better, but for now Stretch or Fit.
                    // Let's do Center Fit logic:
                    /*
                    const ratio = img.getWidth() / img.getHeight();
                    const targetRatio = halfW / area.height;
                    if (ratio > targetRatio) {
                        // Image is wider, fit to width
                        img.setWidth(halfW);
                        const h = halfW / ratio;
                        img.setHeight(h);
                        img.setTop(area.top + (area.height - h)/2);
                    } else {
                        // Image is taller, fit to height
                         img.setHeight(area.height);
                         const w = area.height * ratio;
                         img.setWidth(w);
                         img.setLeft(imgX + (halfW - w)/2);
                    }
                    */
                    // But user might want it to fill the side. Let's just setWidth/Height for now (stretch) or use the logic from SlideUtils.renderImagesInArea
                    // For now, simple stretch to fill layout area as requested by layout
                    // img.setWidth(halfW).setHeight(area.height); -- this stretches.
                    // Let's do fit logic simply:
                    const scale = Math.min(halfW / img.getWidth(), area.height / img.getHeight());
                    const w = img.getWidth() * scale;
                    const h = img.getHeight() * scale;
                    img.setWidth(w).setHeight(h).setLeft(imgX + (halfW - w) / 2).setTop(area.top + (area.height - h) / 2);
                } else {
                    throw new Error("Image insert failed");
                }
            } catch (e) {
                Logger.log('Image insert failed: ' + e);
                // Placeholder
                const ph = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, imgX, area.top, halfW, area.height);
                ph.getFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
                setStyledText(ph, 'Image Placeholder', { align: SlidesApp.ParagraphAlignment.CENTER });
            }
        }

        // Draw Text
        const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, txtX, area.top, halfW, area.height);
        const textContent = points.join('\n');
        setStyledText(textBox, textContent, { size: DEFAULT_THEME.fonts.sizes.body });
    }
}

