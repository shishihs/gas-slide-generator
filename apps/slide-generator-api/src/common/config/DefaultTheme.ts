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
        ghostNum: { left: 600, top: 100, width: 350, height: 350 },
        title: { left: 40, top: 220, width: 880, height: 100 }
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

export const DEFAULT_THEME: SlideTheme = {
    basePx: {
        width: 960,
        height: 540
    },

    fonts: {
        family: 'Noto Sans JP',
        sizes: {
            title: 41,
            date: 16,
            sectionTitle: 38,
            contentTitle: 24,
            subhead: 16,
            body: 14,
            footer: 9,
            chip: 11,
            laneTitle: 13,
            small: 10,
            processStep: 14,
            axis: 12,
            ghostNum: 180
        }
    },

    colors: {
        primary: '#8FB130',
        deepPrimary: '#526717',
        textPrimary: '#333333',
        textSmallFont: '#1F2937',
        backgroundWhite: '#FFFFFF',
        cardBg: '#f6e9f0',
        backgroundGray: '#F1F3F4',
        faintGray: '#FAFAFA',
        ghostGray: '#E0E0E0',
        tableHeaderBg: '#E8EAED',
        laneBorder: '#DADCE0',
        cardBorder: '#DADCE0',
        neutralGray: '#9AA0A6',
        processArrow: '#8FB130'
    },

    diagram: {
        laneGapPx: 24,
        lanePadPx: 10,
        laneTitleHeightPx: 30,
        cardGapPx: 12,
        cardMinHeightPx: 48,
        cardMaxHeightPx: 70,
        arrowHeightPx: 10,
        arrowGapPx: 8
    },

    logos: {
        header: '',
        closing: ''
    },

    footerText: '',

    positions: DEFAULT_POSITIONS,

    backgroundImages: {
        title: '',
        closing: '',
        section: '',
        main: ''
    }
};
