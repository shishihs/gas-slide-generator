import { ISlideGenerator } from '../../../domain/services/ISlideGenerator';
import { LayoutManager } from '../../../common/utils/LayoutManager';
import { CONFIG } from '../../../common/config/SlideConfig';
import {
    setStyledText,
    offsetRect,
    addCucFooter,
    drawArrowBetweenRects,
    setBoldTextSize
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
            titlePlaceholder.asShape().getText().setText(data.title || '');
        }

        const type = (data.layout || data.type || '').toLowerCase();

        const bodyPlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
        const workArea = bodyPlaceholder ?
            { left: bodyPlaceholder.getLeft(), top: bodyPlaceholder.getTop(), width: bodyPlaceholder.getWidth(), height: bodyPlaceholder.getHeight() } :
            layout.getRect('contentSlide.body');

        // switch logic for diagram types
        if (type.includes('timeline')) {
            this.drawTimeline(slide, data, workArea, settings, layout);
        } else if (type.includes('process')) {
            this.drawProcess(slide, data, workArea, settings, layout);
        } else if (type.includes('cycle')) {
            this.drawCycle(slide, data, workArea, settings, layout);
        } else if (type.includes('triangle') || type.includes('pyramid')) {
            this.drawPyramid(slide, data, workArea, settings, layout);
        } else if (type.includes('compare') || type.includes('kaizen')) {
            this.drawComparison(slide, data, workArea, settings, layout);
        } else if (type.includes('stepup') || type.includes('stair')) {
            this.drawStepUp(slide, data, workArea, settings, layout);
        } else if (type.includes('flowchart')) {
            this.drawFlowChart(slide, data, workArea, settings, layout);
        } else if (type.includes('diagram')) { // Lane diagram
            this.drawLanes(slide, data, workArea, settings, layout);
        } else {
            // Fallback
            Logger.log('Diagram logic not implemented for type: ' + type);
        }

        addCucFooter(slide, layout, pageNum, settings, this.creditImageBlob);
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
        line.getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
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
            bodyShape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
            bodyShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);

            // Connector
            const connectorY_start = isAbove ? (cardTop + cardH_pt) : baseY;
            const connectorY_end = isAbove ? baseY : cardTop;
            const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, connectorY_start, x, connectorY_end);
            connector.getLineFill().setSolidFill(CONFIG.COLORS.neutral_gray);
            connector.setWeight(1);

            // Dot
            const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - dotR / 2, baseY - dotR / 2, dotR, dotR);
            dot.getFill().setSolidFill(timelineColors[i]);
            dot.getBorder().setTransparent();

            // Header Text
            const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cardLeft, cardTop, cardW_pt, headerHeight);
            setStyledText(headerTextShape, dateText, {
                size: CONFIG.FONTS.sizes.body,
                bold: true,
                color: CONFIG.COLORS.background_gray,
                align: SlidesApp.ParagraphAlignment.CENTER
            });
            try {
                headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) { }

            // Body Text
            let bodyFontSize = CONFIG.FONTS.sizes.body;
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
                color: CONFIG.COLORS.background_gray,
                align: SlidesApp.ParagraphAlignment.CENTER
            });
            try { header.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Body
            const body = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, bodyLeft, currentY, bodyWPt, boxHPt);
            body.getFill().setSolidFill(CONFIG.COLORS.background_gray);
            body.getBorder().setTransparent();

            // Body Text
            const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, bodyLeft + layout.pxToPt(20), currentY, bodyWPt - layout.pxToPt(40), boxHPt);
            setStyledText(textShape, cleanText, { size: fontSize });
            try { textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            currentY += boxHPt;
            if (i < n - 1) {
                const arrowLeft = area.left + headerWPt / 2 - layout.pxToPt(8);
                const arrow = slide.insertShape(SlidesApp.ShapeType.DOWN_ARROW, arrowLeft, currentY, layout.pxToPt(16), arrowHPt);
                arrow.getFill().setSolidFill(CONFIG.COLORS.process_arrow || CONFIG.COLORS.ghost_gray);
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
            setStyledText(centerTextBox, data.centerText, { size: 20, bold: true, align: SlidesApp.ParagraphAlignment.CENTER, color: CONFIG.COLORS.text_primary });
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
            setStyledText(card, `${subLabelText}\n${labelText}`, { size: fontSize, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });

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
            arrow.getFill().setSolidFill(CONFIG.COLORS.ghost_gray);
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
                size: CONFIG.FONTS.sizes.body,
                bold: true,
                color: CONFIG.COLORS.background_gray,
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
                size: CONFIG.FONTS.sizes.body - 1,
                align: SlidesApp.ParagraphAlignment.START, // Fixed from LEFT
                color: CONFIG.COLORS.text_primary,
                bold: true
            });
            try { textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
        });
    }

    private drawComparison(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager) {
        const leftTitle = data.leftTitle || 'Plan A';
        const rightTitle = data.rightTitle || 'Plan B';
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
        setStyledText(leftHeader, leftTitle, { size: 14, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
        try { leftHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        const leftBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top + headerH, colWidth, area.height - headerH);
        leftBox.getFill().setSolidFill(CONFIG.COLORS.background_gray); // Use light gray instead of F0F0F0 for consistency
        leftBox.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
        setStyledText(leftBox, leftItems.join('\n\n'), { size: CONFIG.FONTS.sizes.body });

        // Right
        const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left + colWidth + gap, area.top, colWidth, headerH);
        rightHeader.getFill().setSolidFill(compareColors.right);
        rightHeader.getBorder().setTransparent();
        setStyledText(rightHeader, rightTitle, { size: 14, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
        try { rightHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

        const rightBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left + colWidth + gap, area.top + headerH, colWidth, area.height - headerH);
        rightBox.getFill().setSolidFill(CONFIG.COLORS.background_gray);
        rightBox.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
        setStyledText(rightBox, rightItems.join('\n\n'), { size: CONFIG.FONTS.sizes.body });
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
        const { laneGap_px, lanePad_px, laneTitle_h_px, cardGap_px, cardMin_h_px, cardMax_h_px, arrow_h_px, arrowGap_px } = CONFIG.DIAGRAM;
        const px = (p: number) => layout.pxToPt(p);

        const laneW = (area.width - px(laneGap_px) * (n - 1)) / n;
        const cardBoxes: any[] = [];

        for (let j = 0; j < n; j++) {
            const lane = lanes[j] || { title: '', items: [] };
            const left = area.left + j * (laneW + px(laneGap_px));

            // Lane Header
            const lt = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, area.top, laneW, px(laneTitle_h_px));
            lt.getFill().setSolidFill(settings.primaryColor);
            lt.getBorder().getLineFill().setSolidFill(settings.primaryColor);
            setStyledText(lt, lane.title || '', { size: CONFIG.FONTS.sizes.laneTitle, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
            try { lt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            // Lane Body
            const laneBodyTop = area.top + px(laneTitle_h_px);
            const laneBodyHeight = area.height - px(laneTitle_h_px);
            const laneBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, laneBodyTop, laneW, laneBodyHeight);
            laneBg.getFill().setSolidFill(CONFIG.COLORS.background_gray);
            laneBg.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);

            const items = Array.isArray(lane.items) ? lane.items : [];
            const rows = Math.max(1, items.length);
            const availH = laneBodyHeight - px(lanePad_px) * 2;
            const idealH = (availH - px(cardGap_px) * (rows - 1)) / rows;
            const cardH = Math.max(px(cardMin_h_px), Math.min(px(cardMax_h_px), idealH));
            const firstTop = laneBodyTop + px(lanePad_px) + Math.max(0, (availH - (cardH * rows + px(cardGap_px) * (rows - 1))) / 2);

            cardBoxes[j] = [];
            for (let i = 0; i < rows; i++) {
                const cardTop = firstTop + i * (cardH + px(cardGap_px));
                const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, left + px(lanePad_px), cardTop, laneW - px(lanePad_px) * 2, cardH);
                card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
                card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
                setStyledText(card, items[i] || '', { size: CONFIG.FONTS.sizes.body });
                try { card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

                cardBoxes[j][i] = {
                    left: left + px(lanePad_px),
                    top: cardTop,
                    width: laneW - px(lanePad_px) * 2,
                    height: cardH
                };
            }
        }

        // Draw Arrows
        const maxRows = Math.max(0, ...cardBoxes.map(a => a ? a.length : 0));
        for (let j = 0; j < n - 1; j++) {
            for (let i = 0; i < maxRows; i++) {
                if (cardBoxes[j] && cardBoxes[j][i] && cardBoxes[j + 1] && cardBoxes[j + 1][i]) {
                    drawArrowBetweenRects(slide, cardBoxes[j][i], cardBoxes[j + 1][i], px(arrow_h_px), px(arrowGap_px), settings);
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
            shape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
            shape.getBorder().getLineFill().setSolidFill(settings.primaryColor);
            shape.getBorder().setWeight(2);
            setStyledText(shape, typeof step === 'string' ? step : step.label || '', { size: CONFIG.FONTS.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER });
            try { shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

            if (i < count - 1) {
                // Arrow
                const ax = x + boxWidth;
                const ay = y + boxHeight / 2;
                const bx = x + boxWidth + gap;
                const by = ay;
                const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, ax, ay, bx, by);
                line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
                line.getLineFill().setSolidFill(CONFIG.COLORS.neutral_gray);
            }
        });
    }
}
