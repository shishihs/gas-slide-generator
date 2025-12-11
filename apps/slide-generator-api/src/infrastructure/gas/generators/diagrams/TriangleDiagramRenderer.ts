import { IDiagramRenderer } from './IDiagramRenderer';
import { LayoutManager } from '../../../../common/utils/LayoutManager';
import { setStyledText } from '../../../../common/utils/SlideUtils';

export class TriangleDiagramRenderer implements IDiagramRenderer {
    render(slide: GoogleAppsScript.Slides.Slide, data: any, area: any, settings: any, layout: LayoutManager): void {
        const theme = layout.getTheme();
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
        trianglePath.getFill().setSolidFill(theme.colors.faintGray);
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
                color: theme.colors.backgroundGray,
                align: SlidesApp.ParagraphAlignment.CENTER
            }, theme);
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
}
