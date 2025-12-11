/**
 * DefaultTheme - Default theme configuration
 * 
 * This contains the same values as the legacy CONFIG object,
 * structured according to the SlideTheme interface.
 */

import { SlideTheme, SlidePositions } from './SlideTheme';

// ============================================================
// Default Positions (Legacy POS_PX)
// ============================================================

const DEFAULT_POSITIONS: SlidePositions = {
    titleSlide: {
        logo: { left: 55, top: 60, width: 135 },
        title: { left: 50, top: 200, width: 830, height: 90 },
        date: { left: 50, top: 450, width: 250, height: 40 }
    },
    contentSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        body: { left: 25, top: 132, width: 910, height: 330 },
        twoColLeft: { left: 25, top: 132, width: 440, height: 330 },
        twoColRight: { left: 495, top: 132, width: 440, height: 330 }
    },
    compareSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        leftBox: { left: 25, top: 112, width: 445, height: 350 },
        rightBox: { left: 490, top: 112, width: 445, height: 350 }
    },
    processSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
    },
    timelineSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
    },
    diagramSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        lanesArea: { left: 25, top: 132, width: 910, height: 330 }
    },
    cardsSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        gridArea: { left: 25, top: 120, width: 910, height: 340 }
    },
    tableSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 130, width: 910, height: 330 }
    },
    progressSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
    },
    kpiSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 130, width: 910, height: 330 }
    },
    agendaSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        body: { left: 25, top: 130, width: 910, height: 350 }
    },
    sectionSlide: {
        ghostNum: { left: 680, top: 320, width: 280, height: 200 },
        title: { left: 40, top: 200, width: 700, height: 100 }
    },
    closingSlide: {
        logo: { left: 380, top: 150, width: 200 },
        message: { left: 40, top: 350, width: 880, height: 80 }
    },
    quoteSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        quoteBox: { left: 50, top: 140, width: 860, height: 280 },
        authorBox: { left: 550, top: 430, width: 360, height: 40 }
    },
    faqSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 130, width: 910, height: 350 }
    },
    imageTextSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        imageArea: { left: 25, top: 130, width: 440, height: 340 },
        textArea: { left: 485, top: 130, width: 450, height: 340 }
    },
    cycleSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
    },
    triangleSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
    },
    pyramidSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
    },
    stepUpSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
    },
    flowChartSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
    },
    headerCardsSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        gridArea: { left: 25, top: 120, width: 910, height: 340 }
    },
    statsCompareSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        leftHeader: { left: 25, top: 132, width: 440, height: 35 },
        rightHeader: { left: 495, top: 132, width: 440, height: 35 },
        leftBox: { left: 25, top: 172, width: 440, height: 290 },
        rightBox: { left: 495, top: 172, width: 440, height: 290 }
    },
    barCompareSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
    },
    footer: {
        leftText: { left: 15, top: 511, width: 500, height: 20 },
        creditImage: { left: 430, top: 514, width: 100, height: 16 },
        rightPage: { right: 15, top: 511, width: 50, height: 20 }
    },
    bottomBar: {
        bar: { left: 0, top: 534, width: 960, height: 6 }
    }
};

// ============================================================
// Default Theme
// ============================================================

// ============================================================
// Shared Base Configuration (Editorial Layout)
// ============================================================

const BASE_POSITIONS: SlidePositions = {
    // Updated positions for "Editorial" look - wider margins (60px side margins)
    titleSlide: {
        logo: { left: 60, top: 60, width: 150 },
        title: { left: 60, top: 200, width: 840, height: 120 },
        date: { left: 60, top: 480, width: 300, height: 40 }
    },
    contentSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 }, // Short elegant underline
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        body: { left: 60, top: 160, width: 840, height: 320 },
        twoColLeft: { left: 60, top: 160, width: 400, height: 320 }, // 40px gap
        twoColRight: { left: 500, top: 160, width: 400, height: 320 }
    },
    compareSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        leftBox: { left: 60, top: 160, width: 400, height: 320 },
        rightBox: { left: 500, top: 160, width: 400, height: 320 }
    },
    processSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    timelineSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    diagramSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        lanesArea: { left: 60, top: 160, width: 840, height: 320 }
    },
    cardsSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        gridArea: { left: 60, top: 160, width: 840, height: 320 }
    },
    tableSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    progressSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    kpiSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    agendaSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        body: { left: 60, top: 160, width: 840, height: 320 }
    },
    sectionSlide: {
        ghostNum: { left: 400, top: 100, width: 550, height: 350 }, // Adjusted
        title: { left: 60, top: 180, width: 700, height: 140 }
    },
    closingSlide: {
        logo: { left: 380, top: 150, width: 200 },
        message: { left: 60, top: 350, width: 840, height: 80 }
    },
    quoteSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        quoteBox: { left: 100, top: 160, width: 760, height: 260 },
        authorBox: { left: 500, top: 440, width: 400, height: 40 }
    },
    faqSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    imageTextSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        imageArea: { left: 60, top: 160, width: 400, height: 320 },
        textArea: { left: 500, top: 160, width: 400, height: 320 }
    },
    cycleSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    triangleSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    pyramidSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    stepUpSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    flowChartSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    headerCardsSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        gridArea: { left: 60, top: 160, width: 840, height: 320 }
    },
    statsCompareSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        leftHeader: { left: 60, top: 160, width: 400, height: 35 },
        rightHeader: { left: 500, top: 160, width: 400, height: 35 },
        leftBox: { left: 60, top: 200, width: 400, height: 280 },
        rightBox: { left: 500, top: 200, width: 400, height: 280 }
    },
    barCompareSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        area: { left: 60, top: 160, width: 840, height: 320 }
    },
    footer: {
        leftText: { left: 60, top: 510, width: 400, height: 20 },
        creditImage: { left: 430, top: 512, width: 100, height: 16 },
        rightPage: { right: 60, top: 510, width: 50, height: 20 }
    },
    bottomBar: {
        bar: { left: 0, top: 536, width: 960, height: 4 }
    }
};

const BASE_THEME_OPTS = {
    basePx: {
        width: 960,
        height: 540
    },
    fonts: {
        // Noto Sans JP is good, but let's assume we can use different weights via styles
        family: "Noto Sans JP",
        sizes: {
            title: 48,
            date: 14,
            sectionTitle: 52,
            contentTitle: 28,
            subhead: 20,        // Increased visibility
            body: 18,           // "18pt minimum" rule
            footer: 11,
            chip: 12,
            laneTitle: 16,
            small: 12,
            processStep: 16,
            axis: 12,
            ghostNum: 250
        }
    },
    diagram: {
        laneGapPx: 30,          // Wider gaps
        lanePadPx: 20,          // More padding
        laneTitleHeightPx: 40,
        cardGapPx: 20,          // Airy layout
        cardMinHeightPx: 60,
        cardMaxHeightPx: 90,
        arrowHeightPx: 8,       // Thinner, elegant arrows
        arrowGapPx: 12
    },
    logos: {
        header: "",
        closing: ""
    },
    footerText: "",
    positions: BASE_POSITIONS,
    backgroundImages: {
        title: "",
        closing: "",
        section: "",
        main: ""
    }
};

// ============================================================
// Themes
// ============================================================

export const THEME_GREEN: SlideTheme = {
    ...BASE_THEME_OPTS,
    colors: {
        primary: "#8FB130",     // Brand Main (Vibrant Leaf Green)
        deepPrimary: "#526717", // Brand Deep (Moss Green)
        textPrimary: "#333333", // Soft Black
        textSmallFont: "#545454", // Slightly warmer gray for text
        backgroundWhite: "#FFFFFF",
        cardBg: "#FFFFFF",
        backgroundGray: "#F8F9FA", // Keep neutral background for logos
        faintGray: "#F8F9FA",
        ghostGray: "#E2E4E6",   // 
        tableHeaderBg: "#E8EAE6", // Hint of green in gray
        laneBorder: "#E0E0E0",
        cardBorder: "#E0E0E0",
        neutralGray: "#9E9E9E",
        processArrow: "#526717" // Use Deep color for contrast
    }
};

export const THEME_BLUE: SlideTheme = {
    ...BASE_THEME_OPTS,
    colors: {
        primary: "#3498DB",     // Standard Blue
        deepPrimary: "#2C3E50", // Dark Blue
        textPrimary: "#333333",
        textSmallFont: "#555555",
        backgroundWhite: "#FFFFFF",
        cardBg: "#FFFFFF",
        backgroundGray: "#F4F6F7",
        faintGray: "#ECF0F1",
        ghostGray: "#BDC3C7",
        tableHeaderBg: "#ECF0F1",
        laneBorder: "#BDC3C7",
        cardBorder: "#E0E0E0",
        neutralGray: "#95A5A6",
        processArrow: "#2C3E50"
    }
};
export const THEME_RED: SlideTheme = {
    ...BASE_THEME_OPTS,
    colors: {
        primary: "#E74C3C",     // Red
        deepPrimary: "#C0392B", // Dark Red
        textPrimary: "#333333",
        textSmallFont: "#555555",
        backgroundWhite: "#FFFFFF",
        cardBg: "#FFFFFF",
        backgroundGray: "#F4F6F7",
        faintGray: "#ECF0F1",
        ghostGray: "#BDC3C7",
        tableHeaderBg: "#ECF0F1",
        laneBorder: "#BDC3C7",
        cardBorder: "#E0E0E0",
        neutralGray: "#95A5A6",
        processArrow: "#C0392B"
    }
};

export const THEME_GRAYSCALE: SlideTheme = {
    ...BASE_THEME_OPTS,
    colors: {
        primary: "#666666",
        deepPrimary: "#333333",
        textPrimary: "#000000",
        textSmallFont: "#666666",
        backgroundWhite: "#FFFFFF",
        cardBg: "#FFFFFF",
        backgroundGray: "#EEEEEE",
        faintGray: "#EEEEEE",
        ghostGray: "#CCCCCC",
        tableHeaderBg: "#DDDDDD",
        laneBorder: "#BBBBBB",
        cardBorder: "#CCCCCC",
        neutralGray: "#999999",
        processArrow: "#333333"
    }
};

export const DEFAULT_THEME = THEME_GREEN;

export const AVAILABLE_THEMES = {
    "Green": THEME_GREEN,
    "Blue": THEME_BLUE,
    "Red": THEME_RED,
    "Grayscale": THEME_GRAYSCALE
};


