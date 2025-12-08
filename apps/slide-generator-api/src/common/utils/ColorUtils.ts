
export function hexToRgb(hex: string): { r: number; g: number; b: number } | null {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    } : null;
}

export function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
    r /= 255;
    g /= 255;
    b /= 255;
    const max = Math.max(r, g, b),
        min = Math.min(r, g, b);
    let h = 0, s, l = (max + min) / 2;
    if (max === min) {
        h = s = 0;
    } else {
        const d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        switch (max) {
            case r:
                h = (g - b) / d + (g < b ? 6 : 0);
                break;
            case g:
                h = (b - r) / d + 2;
                break;
            case b:
                h = (r - g) / d + 4;
                break;
        }
        h /= 6;
    }
    return {
        h: h * 360,
        s: s * 100,
        l: l * 100
    };
}

export function hslToHex(h: number, s: number, l: number): string {
    s /= 100;
    l /= 100;
    const c = (1 - Math.abs(2 * l - 1)) * s,
        x = c * (1 - Math.abs((h / 60) % 2 - 1)),
        m = l - c / 2;
    let r = 0,
        g = 0,
        b = 0;
    if (0 <= h && h < 60) {
        r = c;
        g = x;
        b = 0;
    } else if (60 <= h && h < 120) {
        r = x;
        g = c;
        b = 0;
    } else if (120 <= h && h < 180) {
        r = 0;
        g = c;
        b = x;
    } else if (180 <= h && h < 240) {
        r = 0;
        g = x;
        b = c;
    } else if (240 <= h && h < 300) {
        r = x;
        g = 0;
        b = c;
    } else if (300 <= h && h < 360) {
        r = c;
        g = 0;
        b = x;
    }
    r = Math.round((r + m) * 255);
    g = Math.round((g + m) * 255);
    b = Math.round((b + m) * 255);
    return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

export function generateTintedGray(tintColorHex: string, saturation: number, lightness: number): string {
    const rgb = hexToRgb(tintColorHex);
    if (!rgb) return '#F8F9FA';
    const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
    return hslToHex(hsl.h, saturation, lightness);
}

export function lightenColor(color: string, amount: number): string {
    const rgb = hexToRgb(color);
    if (!rgb) return color;
    const lighten = (c: number) => Math.min(255, Math.round(c + (255 - c) * amount));
    const newR = lighten(rgb.r);
    const newG = lighten(rgb.g);
    const newB = lighten(rgb.b);
    return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}

export function darkenColor(color: string, amount: number): string {
    const rgb = hexToRgb(color);
    if (!rgb) return color;
    const darken = (c: number) => Math.max(0, Math.round(c * (1 - amount)));
    const newR = darken(rgb.r);
    const newG = darken(rgb.g);
    const newB = darken(rgb.b);
    return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}

export function generatePyramidColors(baseColor: string, levels: number): string[] {
    const colors = [];
    for (let i = 0; i < levels; i++) {
        const lightenAmount = (i / Math.max(1, levels - 1)) * 0.6;
        colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
}

export function generateStepUpColors(baseColor: string, steps: number): string[] {
    const colors = [];
    for (let i = 0; i < steps; i++) {
        const lightenAmount = 0.6 * (1 - (i / Math.max(1, steps - 1)));
        colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
}

export function generateProcessColors(baseColor: string, steps: number): string[] {
    const colors = [];
    for (let i = 0; i < steps; i++) {
        const lightenAmount = 0.5 * (1 - (i / Math.max(1, steps - 1)));
        colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
}

export function generateTimelineCardColors(baseColor: string, milestones: number): string[] {
    const colors = [];
    for (let i = 0; i < milestones; i++) {
        const lightenAmount = 0.4 * (1 - (i / Math.max(1, milestones - 1)));
        colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
}

export function generateCompareColors(baseColor: string): { left: string; right: string } {
    return {
        left: darkenColor(baseColor, 0.3),
        right: baseColor
    };
}
