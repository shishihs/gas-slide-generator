import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { DEFAULT_THEME } from '../../../../common/config/DefaultTheme';
import { setStyledText } from '../../../../common/utils/SlideUtils';
import { generateTimelineCardColors } from '../../../../common/utils/ColorUtils';

export class TimelineDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
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
}
