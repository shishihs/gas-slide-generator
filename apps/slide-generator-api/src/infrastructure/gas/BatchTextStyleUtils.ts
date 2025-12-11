import { RequestFactory } from './RequestFactory';

export class BatchTextStyleUtils {
    /**
     * Replicates the behavior of setStyledText but triggers Requests.
     * Supports basic Markdown '**bold**' parsing.
     */
    static setText(slideId: string, objectId: string, text: string, baseStyle: any, theme: any): GoogleAppsScript.Slides.Schema.Request[] {
        const requests: GoogleAppsScript.Slides.Schema.Request[] = [];

        // Simple parser: Split by ** to detect bold ranges
        // e.g. "Title **Bold** Suffix" -> ["Title ", "Bold", " Suffix"]
        const parts = (text || '').split('**');
        let cleanText = '';
        const boldRanges: { start: number; length: number }[] = [];

        let isBold = false;
        parts.forEach(part => {
            if (isBold && part.length > 0) {
                boldRanges.push({ start: cleanText.length, length: part.length });
            }
            cleanText += part;
            // Toggle bold state for next part (unless it was empty split which implies consecutive markers? 
            // 'Eq **A**' -> ['Eq ', 'A', ''] -> 3 parts. 
            // Part 0: 'Eq ' (not bold). Next isBold=true.
            // Part 1: 'A' (bold). Next isBold=false.
            // Part 2: '' (not bold). Next isBold=true. 
            // Logic holds for balanced markers.
            isBold = !isBold;
        });

        // 1. Insert Clean Text
        requests.push(RequestFactory.insertText(objectId, cleanText));

        // 2. Base Style
        // Map common keys to RequestFactory format
        const styleReq = {
            bold: baseStyle.bold || false,
            fontSize: baseStyle.size || baseStyle.fontSize || theme.fonts.sizes.body,
            color: baseStyle.color || theme.colors.textPrimary, // Default color
            fontFamily: theme.fonts.family
        };
        requests.push(RequestFactory.updateTextStyle(objectId, styleReq, 0, cleanText.length));

        // 3. Bold Ranges
        boldRanges.forEach(r => {
            requests.push(RequestFactory.updateTextStyle(objectId, { bold: true }, r.start, r.length));
        });

        // 4. Alignment
        if (baseStyle.align) {
            // Map SlidesApp alignment enums or strings to API strings
            let alignStr = 'START';
            const align = String(baseStyle.align).toUpperCase();
            if (align.includes('CENTER') || align.includes('MIDDLE')) alignStr = 'CENTER';
            else if (align.includes('END') || align.includes('RIGHT')) alignStr = 'END';
            else if (align.includes('JUSTIFY')) alignStr = 'JUSTIFIED';

            requests.push(RequestFactory.updateParagraphStyle(objectId, alignStr));
        }

        return requests;
    }
}
