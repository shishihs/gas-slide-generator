/**
 * SlideTheme - Interface for slide styling configuration
 * 
 * This replaces the global CONFIG object with a dependency-injectable theme system.
 * All styling, colors, fonts, and layout information flows through this interface.
 */

// ============================================================
// Font Configuration
// ============================================================

export interface FontSizes {
    title: number;
    date: number;
    sectionTitle: number;
    contentTitle: number;
    subhead: number;
    body: number;
    footer: number;
    chip: number;
    laneTitle: number;
    small: number;
    processStep: number;
    axis: number;
    ghostNum: number;
}

export interface FontConfig {
    family: string;
    sizes: FontSizes;
}

// ============================================================
// Color Configuration
// ============================================================

export interface ThemeColors {
    primary: string;
    deepPrimary: string;
    textPrimary: string;
    textSmallFont: string;
    backgroundWhite: string;
    cardBg: string;
    backgroundGray: string;
    faintGray: string;
    ghostGray: string;
    tableHeaderBg: string;
    laneBorder: string;
    cardBorder: string;
    neutralGray: string;
    processArrow: string;
}

// ============================================================
// Layout Configuration
// ============================================================

export interface BaseDimensions {
    width: number;
    height: number;
}

export interface DiagramConfig {
    laneGapPx: number;
    lanePadPx: number;
    laneTitleHeightPx: number;
    cardGapPx: number;
    cardMinHeightPx: number;
    cardMaxHeightPx: number;
    arrowHeightPx: number;
    arrowGapPx: number;
}

export interface LogoConfig {
    header: string;
    closing: string;
}

// ============================================================
// Position Configuration (for fallback when placeholders missing)
// ============================================================

export interface RectPosition {
    left?: number;
    right?: number;
    top: number;
    width: number;
    height?: number;
}

export interface SlidePositions {
    [slideType: string]: {
        [elementName: string]: RectPosition;
    };
}

// ============================================================
// Main Theme Interface
// ============================================================

export interface SlideTheme {
    /** Base dimensions in pixels (960x540 standard) */
    basePx: BaseDimensions;

    /** Font configuration */
    fonts: FontConfig;

    /** Color palette */
    colors: ThemeColors;

    /** Diagram-specific settings */
    diagram: DiagramConfig;

    /** Logo URLs/IDs */
    logos: LogoConfig;

    /** Footer text */
    footerText: string;

    /** Position definitions for fallback layouts */
    positions: SlidePositions;

    /** Background image URLs (optional) */
    backgroundImages?: {
        title?: string;
        closing?: string;
        section?: string;
        main?: string;
    };
}

// ============================================================
// Theme Creation Helper
// ============================================================

/**
 * Create a partial theme that overrides specific values from a base theme
 */
export type PartialTheme = Partial<{
    basePx: Partial<BaseDimensions>;
    fonts: Partial<FontConfig>;
    colors: Partial<ThemeColors>;
    diagram: Partial<DiagramConfig>;
    logos: Partial<LogoConfig>;
    footerText: string;
    positions: SlidePositions;
    backgroundImages: SlideTheme['backgroundImages'];
}>;

/**
 * Merge a partial theme with a base theme
 */
export function mergeTheme(base: SlideTheme, overrides: PartialTheme): SlideTheme {
    return {
        basePx: { ...base.basePx, ...overrides.basePx },
        fonts: {
            family: overrides.fonts?.family ?? base.fonts.family,
            sizes: { ...base.fonts.sizes, ...overrides.fonts?.sizes }
        },
        colors: { ...base.colors, ...overrides.colors },
        diagram: { ...base.diagram, ...overrides.diagram },
        logos: { ...base.logos, ...overrides.logos },
        footerText: overrides.footerText ?? base.footerText,
        positions: overrides.positions ?? base.positions,
        backgroundImages: { ...base.backgroundImages, ...overrides.backgroundImages }
    };
}
