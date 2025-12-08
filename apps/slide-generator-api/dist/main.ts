/**
* - pako.js (https://github.com/nodeca/pako) : MIT License
* - UPNG.js (https://github.com/photopea/UPNG.js) : MIT License
*/

function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : null;
}
function rgbToHsl(r, g, b) {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  let h, s, l = (max + min) / 2;
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
function hslToHex(h, s, l) {
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
function generateTintedGray(tintColorHex, saturation, lightness) {
  const rgb = hexToRgb(tintColorHex);
  if (!rgb) return '#F8F9FA';
  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  return hslToHex(hsl.h, saturation, lightness);
}
function generatePyramidColors(baseColor, levels) {
  const colors = [];
  for (let i = 0; i < levels; i++) {
    const lightenAmount = (i / Math.max(1, levels - 1)) * 0.6;
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}
function lightenColor(color, amount) {
  const rgb = hexToRgb(color);
  if (!rgb) return color;
  const lighten = (c) => Math.min(255, Math.round(c + (255 - c) * amount));
  const newR = lighten(rgb.r);
  const newG = lighten(rgb.g);
  const newB = lighten(rgb.b);
  return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}
function darkenColor(color, amount) {
  const rgb = hexToRgb(color);
  if (!rgb) return color;
  const darken = (c) => Math.max(0, Math.round(c * (1 - amount)));
  const newR = darken(rgb.r);
  const newG = darken(rgb.g);
  const newB = darken(rgb.b);
  return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}
function generateStepUpColors(baseColor, steps) {
  const colors = [];
  for (let i = 0; i < steps; i++) {
    const lightenAmount = 0.6 * (1 - (i / Math.max(1, steps - 1)));
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}
function generateProcessColors(baseColor, steps) {
  const colors = [];
  for (let i = 0; i < steps; i++) {
    const lightenAmount = 0.5 * (1 - (i / Math.max(1, steps - 1)));
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}
function generateTimelineCardColors(baseColor, milestones) {
  const colors = [];
  for (let i = 0; i < milestones; i++) {
    const lightenAmount = 0.4 * (1 - (i / Math.max(1, milestones - 1)));
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}
function generateCompareColors(baseColor) {
  return {
    left: darkenColor(baseColor, 0.3),
    right: baseColor
  };
}
const CONFIG = {
  BASE_PX: {
    W: 960,
    H: 540
  },
  BACKGROUND_IMAGES: {
    title: '',
    closing: '',
    section: '',
    main: ''
  },
  POS_PX: {
    titleSlide: {
      logo: {
        left: 55,
        top: 60,
        width: 135
      },
      title: {
        left: 50,
        top: 200,
        width: 830,
        height: 90
      },
      date: {
        left: 50,
        top: 450,
        width: 250,
        height: 40
      }
    },
    contentSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      body: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      },
      twoColLeft: {
        left: 25,
        top: 132,
        width: 440,
        height: 330
      },
      twoColRight: {
        left: 495,
        top: 132,
        width: 440,
        height: 330
      }
    },
    compareSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      leftBox: {
        left: 25,
        top: 112,
        width: 445,
        height: 350
      },
      rightBox: {
        left: 490,
        top: 112,
        width: 445,
        height: 350
      }
    },
    processSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    timelineSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    diagramSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      lanesArea: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    cardsSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      gridArea: {
        left: 25,
        top: 120,
        width: 910,
        height: 340
      }
    },
    tableSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 130,
        width: 910,
        height: 330
      }
    },
    progressSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    quoteSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 100,
        width: 910,
        height: 40
      }
    },
    kpiSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      gridArea: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    triangleSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 110,
        width: 910,
        height: 350
      }
    },
    flowChartSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      singleRow: {
        left: 25,
        top: 160,
        width: 910,
        height: 180
      },
      upperRow: {
        left: 25,
        top: 150,
        width: 910,
        height: 120
      },
      lowerRow: {
        left: 25,
        top: 290,
        width: 910,
        height: 120
      }
    },
    stepUpSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      stepArea: {
        left: 25,
        top: 130,
        width: 910,
        height: 330
      }
    },
    imageTextSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      leftImage: {
        left: 25,
        top: 140,
        width: 440,
        height: 310
      },
      leftImageCaption: {
        left: 25,
        top: 470,
        width: 440,
        height: 30
      },
      rightText: {
        left: 485,
        top: 140,
        width: 450,
        height: 310
      },
      leftText: {
        left: 25,
        top: 140,
        width: 450,
        height: 310
      },
      rightImage: {
        left: 495,
        top: 140,
        width: 440,
        height: 310
      },
      rightImageCaption: {
        left: 495,
        top: 430,
        width: 470,
        height: 30
      }
    },
      pyramidSlide: {
        headerLogo: {
          right: 20,
          top: 20,
          width: 75
        },
        title: {
          left: 25,
          top: 20,
          width: 830,
          height: 65
        },
        titleUnderline: {
          left: 25,
          top: 80,
          width: 260,
          height: 4
        },
        subhead: {
          left: 25,
          top: 90,
          width: 910,
          height: 40
        },
        pyramidArea: {
          left: 25,
          top: 120,
          width: 910,
          height: 360
        }
      },
    sectionSlide: {
      title: {
        left: 55,
        top: 230,
        width: 840,
        height: 80
      },
      ghostNum: {
        left: 35,
        top: 120,
        width: 400,
        height: 200
      }
    },
    footer: {
      leftText: {
        left: 15,
        top: 511,
        width: 250,
        height: 20
      },
      creditImage: {
        right: 55,
        top: 518,
        width: 150,
        height: 10
      },
      rightPage: {
        right: 15,
        top: 511,
        width: 50,
        height: 20
      }
    },
    bottomBar: {
      left: 0,
      top: 534,
      width: 960,
      height: 6
    }
  },
  FONTS: {
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
  COLORS: {
    primary_color: '#4285F4',
    text_primary: '#333333',
    text_small_font: '#1F2937',
    background_white: '#FFFFFF',
    card_bg: '#f6e9f0',
    background_gray: '',
    faint_gray: '',
    ghost_gray: '',
    table_header_bg: '',
    lane_border: '',
    card_border: '',
    neutral_gray: '',
    process_arrow: ''
  },
  DIAGRAM: {
    laneGap_px: 24,
    lanePad_px: 10,
    laneTitle_h_px: 30,
    cardGap_px: 12,
    cardMin_h_px: 48,
    cardMax_h_px: 70,
    arrow_h_px: 10,
    arrowGap_px: 8
  },
  LOGOS: {
    header: '',
    closing: ''
  },
  FOOTER_TEXT: `© ${new Date().getFullYear()} Your Company`
};
function doGet(e) {
  const activationData = checkUserActivation();
  const htmlTemplate = HtmlService.createTemplateFromFile('無題.html');
  const baseSettings = loadSettings();
  htmlTemplate.settings = baseSettings;
  htmlTemplate.activationData = activationData;
  return htmlTemplate.evaluate().setTitle('Google Slide Generator').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}
function saveSettings(settings) {
  try {
    const storableSettings = Object.assign({}, settings);
    storableSettings.showTitleUnderline = String(storableSettings.showTitleUnderline);
    storableSettings.showBottomBar = String(storableSettings.showBottomBar);
    storableSettings.showDateColumn = String(storableSettings.showDateColumn);
    storableSettings.enableGradient = String(storableSettings.enableGradient);
    storableSettings.showPageNumber = String(storableSettings.showPageNumber);
    PropertiesService.getUserProperties().setProperties(storableSettings, false);
    return {
      status: 'success',
      message: '設定を保存しました。'
    };
  } catch (e) {
    return {
      status: 'error',
      message: `設定の保存中にエラーが発生しました: ${e.message}`
    };
  }
}
function saveSelectedPreset(presetName) {
  try {
    PropertiesService.getUserProperties().setProperty('selectedPreset', presetName);
    return {
      status: 'success',
      message: 'プリセット選択を保存しました。'
    };
  } catch (e) {
    return {
      status: 'error',
      message: `プリセットの保存中にエラーが発生しました: ${e.message}`
    };
  }
}
function loadSettings() {
  const properties = PropertiesService.getUserProperties().getProperties();
  const scriptProperties = PropertiesService.getScriptProperties().getProperties();
  const DEFAULT_DRIVE_URL = "https://drive.google.com/drive/my-drive";
  const driveFolderUrl = (properties.driveFolderUrl === undefined || properties.driveFolderUrl === "") ?
    DEFAULT_DRIVE_URL : properties.driveFolderUrl;
  return {
    primaryColor: properties.primaryColor || '#4285F4',
    largeFontColor: properties.largeFontColor || '#333333',
    smallFontColor: properties.smallFontColor || '#1F2937',
    backgroundColor: properties.backgroundColor || '#FFFFFF',
    gradientStart: properties.gradientStart || '#4285F4',
    gradientEnd: properties.gradientEnd || '#ff52df',
    fontFamily: properties.fontFamily || 'Noto Sans JP',
    showTitleUnderline: properties.showTitleUnderline === 'false' ? false : true,
    showBottomBar: properties.showBottomBar === 'false' ? false : true,
    showDateColumn: properties.showDateColumn === 'false' ? false : true,
    showPageNumber: properties.showPageNumber === 'false' ? false : true,
    enableGradient: properties.enableGradient === 'true' ? true : false,
    footerText: (properties.footerText === undefined) ? '© Google Inc.' : properties.footerText,
  headerLogoUrl: (properties.headerLogoUrl === undefined) ? 'https://upload.wikimedia.org/wikipedia/commons/thumb/8/8a/Google_Gemini_logo.svg/2560px-Google_Gemini_logo.svg.png' : properties.headerLogoUrl,
    closingLogoUrl: (properties.closingLogoUrl === undefined) ? 'https://upload.wikimedia.org/wikipedia/commons/thumb/8/8a/Google_Gemini_logo.svg/2560px-Google_Gemini_logo.svg.png' : properties.closingLogoUrl,
    titleBgUrl: properties.titleBgUrl || '',
    sectionBgUrl: properties.sectionBgUrl || '',
    mainBgUrl: properties.mainBgUrl || '',
    closingBgUrl: properties.closingBgUrl || '',
    driveFolderUrl: properties.driveFolderUrl || driveFolderUrl,
    selectedPreset: properties.selectedPreset || 'default',
    geminiGemUrl: scriptProperties.geminiGemUrl || null,
    graphColorTheme: properties.graphColorTheme || 'primary',
  };
}
function createBlankPresentation(slideData, settings) {
  try {
    const rawTitle = (slideData[0] && slideData[0].type === 'title' ? String(slideData[0].title || '') : 'Google Slide Generator Presentation');
    const singleLineTitle = rawTitle.replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
    let finalName;
    if (settings.showDateColumn) {
      const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy.MM.dd');
      finalName = singleLineTitle ? (singleLineTitle + ' ' + dateStr) : ('Google Slide Generator Presentation ' + dateStr);
    } else {
      finalName = singleLineTitle || 'Google Slide Generator Presentation';
    }
    const presentation = SlidesApp.create(finalName);
    const defaultSlide = presentation.getSlides()[0];
    if (defaultSlide) {
      defaultSlide.remove();
    }
    let extractedId = null;
    const url = settings.driveFolderUrl;
    if (url) {
      const drivePatterns = [
        /\/folders\/([a-zA-Z0-9_-]+)/,
        /\?id=([a-zA-Z0-9_-]+)/
      ];
      for (let i = 0; i < drivePatterns.length; i++) {
        const match = url.match(drivePatterns[i]);
        if (match && match[1]) {
          extractedId = match[1];
          break;
        }
      }
      if (!extractedId && /^[a-zA-Z0-9_-]{25,}$/.test(url)) {
        extractedId = url;
      }
    }
    if (extractedId && extractedId.trim()) {
      try {
        DriveApp.getFileById(presentation.getId()).moveTo(DriveApp.getFolderById(extractedId.trim()));
      } catch (e) {
      }
    }
    return presentation.getUrl();
  } catch (e) {
    throw new Error(`Server error in createBlankPresentation: ${e.message}`);
  }
}
function updateDynamicColors(settings) {
  const primary = settings.primaryColor;
  CONFIG.COLORS.background_gray = generateTintedGray(primary, 10, 99.5); 
  CONFIG.COLORS.faint_gray = generateTintedGray(primary, 10, 93);
  CONFIG.COLORS.ghost_gray = generateTintedGray(primary, 38, 88);
  CONFIG.COLORS.table_header_bg = generateTintedGray(primary, 20, 94);
  CONFIG.COLORS.lane_border = generateTintedGray(primary, 15, 92);
  CONFIG.COLORS.card_border = generateTintedGray(primary, 15, 88);
  CONFIG.COLORS.neutral_gray = generateTintedGray(primary, 5, 62);
  CONFIG.COLORS.process_arrow = CONFIG.COLORS.ghost_gray;
}
function generateSlidesFromWebApp(slideDataString, settings, presentationId = null, imageUpdateOption = 'update') {
  try {
    const slideData = JSON.parse(slideDataString);
    return createPresentation(slideData, settings, presentationId, imageUpdateOption);
  } catch (e) {
    throw new Error(`Server error: ${e.message}`);
  }
}
let __SECTION_COUNTER = 0;
let __SLIDE_DATA_FOR_AGENDA = [];
let __CREDIT_IMAGE_BLOB = null;
const CREDIT_IMAGE_LINK = 'https://note.com/majin_108';
function createPresentation(slideData, settings, presentationId = null, imageUpdateOption = 'update') {
  updateDynamicColors(settings);
  CONFIG.COLORS.primary_color = settings.primaryColor || CONFIG.COLORS.primary_color;
  CONFIG.COLORS.text_primary = settings.largeFontColor || '#333333';
  CONFIG.COLORS.text_small_font = settings.smallFontColor || '#1F2937';
  if (settings.backgroundColor && settings.backgroundColor.toUpperCase() !== '#FFFFFF') {
    CONFIG.COLORS.background_white = settings.backgroundColor;
  } else {
    CONFIG.COLORS.background_white = '#FFFFFF';
  }
  CONFIG.FOOTER_TEXT = settings.footerText;
  CONFIG.FONTS.family = settings.fontFamily || CONFIG.FONTS.family;
  CONFIG.LOGOS.header = settings.headerLogoUrl;
  CONFIG.LOGOS.closing = settings.closingLogoUrl;
  CONFIG.BACKGROUND_IMAGES.title = settings.titleBgUrl;
  CONFIG.BACKGROUND_IMAGES.closing = settings.closingBgUrl;
  CONFIG.BACKGROUND_IMAGES.section = settings.sectionBgUrl;
  CONFIG.BACKGROUND_IMAGES.main = settings.mainBgUrl;
  __SLIDE_DATA_FOR_AGENDA = slideData;
  __CREDIT_IMAGE_BLOB = null;
  if (settings.creditImageBase64) {
    try {
      const parts = settings.creditImageBase64.split(',');
      if (parts.length !== 2) throw new Error('Invalid Base64 format for credit image.');
      const mimeType = parts[0].match(/:(.*?);/)[1];
      const decodedData = Utilities.base64Decode(parts[1]);
      __CREDIT_IMAGE_BLOB = Utilities.newBlob(decodedData, mimeType, 'credit.png');
    } catch (e) {
    }
  } else {
  }
  const rawTitle = (slideData[0] && slideData[0].type === 'title' ? String(slideData[0].title || '') : 'Google Slide Generator Presentation');
  const singleLineTitle = rawTitle.replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
  let finalName;
  if (settings.showDateColumn) {
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy.MM.dd');
    finalName = singleLineTitle ? (singleLineTitle + ' ' + dateStr) : ('Google Slide Generator Presentation ' + dateStr);
  } else {
    finalName = singleLineTitle || 'Google Slide Generator Presentation';
  }
  let presentation;
  let isUpdating = false;
  if (presentationId) {
    try {
      const file = DriveApp.getFileById(presentationId);
      presentation = SlidesApp.openById(file.getId());
      presentation.setName(finalName);
      if (settings.driveFolderId && settings.driveFolderId.trim()) {
        try {
          file.moveTo(DriveApp.getFolderById(settings.driveFolderId.trim()));
        } catch (moveError) {
        }
      } else {
        try {
          file.moveTo(DriveApp.getRootFolder());
        } catch (moveError) {
        }
      }
      isUpdating = true;
    } catch (e) {
      presentation = SlidesApp.create(finalName);
      const defaultSlide = presentation.getSlides()[0];
      if (defaultSlide) {
        defaultSlide.remove();
      }
      if (settings.driveFolderId && settings.driveFolderId.trim()) {
        try {
          DriveApp.getFileById(presentation.getId()).moveTo(DriveApp.getFolderById(settings.driveFolderId.trim()));
        } catch (moveError) {
        }
      }
    }
  } else {
    presentation = SlidesApp.create(finalName);
    const defaultSlide = presentation.getSlides()[0];
    if (defaultSlide) {
      defaultSlide.remove();
    }
    if (settings.driveFolderId && settings.driveFolderId.trim()) {
      try {
        DriveApp.getFileById(presentation.getId()).moveTo(DriveApp.getFolderById(settings.driveFolderId.trim()));
      } catch (e) {
      }
    }
  }
  __SECTION_COUNTER = 0;
  const layout = createLayoutManager(presentation.getPageWidth(), presentation.getPageHeight());
  let pageCounter = 0;
  const existingSlides = presentation.getSlides();
  const numExistingSlides = existingSlides.length;
  const numNewSlides = slideData.length;
  for (let i = 0; i < numNewSlides; i++) {
    const data = slideData[i];
    let slide;
    let slideAction = '';
    if (data.type !== 'title' && data.type !== 'closing') {
      pageCounter++;
    }
    const generator = slideGenerators[data.type];
    if (!generator) {
      slideAction = 'skipped';
      continue;
    }
    try {
      if (isUpdating && i < numExistingSlides) {
        slide = existingSlides[i];
        clearSlideContents(slide, imageUpdateOption);
        slideAction = 'updated';
      } else {
        slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        slideAction = 'created';
      }
      generator(slide, data, layout, pageCounter, settings, imageUpdateOption);
      if (data.notes) {
        const cleanedNotes = cleanSpeakerNotes(data.notes);
        slide.getNotesPage().getSpeakerNotesShape().getText().setText(cleanedNotes);
      }
    } catch (e) {
      if (slide && slideAction === 'created') {
        try { slide.remove(); } catch (removeError) {}
      } else if (slide && slideAction === 'updated') {
        try {
          clearSlideContents(slide, 'update');
          const errorBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 50, 50, 400, 100);
          errorBox.getText().setText(`Error generating this slide:\n${e.message}`);
        } catch (errorBoxError) {}
      }
    }
  }
  if (isUpdating && numExistingSlides > numNewSlides) {
    for (let i = numExistingSlides - 1; i >= numNewSlides; i--) {
      try {
        existingSlides[i].remove();
      } catch (removeError) {
      }
    }
  }
  return presentation.getUrl();
}
function clearSlideContents(slide, imageUpdateOption) {
  const pageElements = slide.getPageElements();
  pageElements.forEach(element => {
    try {
      if (imageUpdateOption === 'keep' && element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
      } else {
        element.remove();
      }
    } catch (e) {
    }
  });
  try {
    slide.getBackground().setSolidFill('#FFFFFF');
  } catch(e) {
  }
}
function updateSingleSlide(presentationId, slideIndex, singleSlideDataString, settings, imageUpdateOption) {
  if (!presentationId || slideIndex == null || slideIndex < 0 || !singleSlideDataString) {
    return { status: 'error', message: '無効なパラメータが指定されました。' };
  }
  try {
    __CREDIT_IMAGE_BLOB = null;
    if (settings.creditImageBase64) {
      try {
        const parts = settings.creditImageBase64.split(',');
        if (parts.length !== 2) throw new Error('Invalid Base64 format for credit image.');
        const mimeType = parts[0].match(/:(.*?);/)[1];
        const decodedData = Utilities.base64Decode(parts[1]);
        __CREDIT_IMAGE_BLOB = Utilities.newBlob(decodedData, mimeType, 'credit.png');
      } catch (e) {
      }
    } else {
    }
    let presentation;
    try {
      presentation = SlidesApp.openById(presentationId);
    } catch (e) {
      return { status: 'error', message: `プレゼンテーション (ID: ${presentationId}) を開けませんでした: ${e.message}` };
    }
    const slides = presentation.getSlides();
    if (slideIndex >= slides.length) {
      return { status: 'error', message: `指定されたページ番号 (${slideIndex + 1}) のスライドが存在しません。` };
    }
    const slide = slides[slideIndex];
    let data;
    try {
      data = JSON.parse(singleSlideDataString);
    } catch (e) {
      return { status: 'error', message: `スライドデータのJSON形式が不正です: ${e.message}` };
    }
    if (!data || !data.type) {
       return { status: 'error', message: 'スライドデータに type プロパティが含まれていません。' };
    }
    updateDynamicColors(settings);
    CONFIG.COLORS.primary_color = settings.primaryColor || CONFIG.COLORS.primary_color;
    CONFIG.COLORS.text_primary = settings.largeFontColor || '#333333';
    CONFIG.COLORS.text_small_font = settings.smallFontColor || '#1F2937';
    if (settings.backgroundColor && settings.backgroundColor.toUpperCase() !== '#FFFFFF') {
      CONFIG.COLORS.background_white = settings.backgroundColor;
    } else {
      CONFIG.COLORS.background_white = '#FFFFFF';
    }
    CONFIG.FOOTER_TEXT = settings.footerText;
    CONFIG.FONTS.family = settings.fontFamily || CONFIG.FONTS.family;
    CONFIG.LOGOS.header = settings.headerLogoUrl;
    CONFIG.LOGOS.closing = settings.closingLogoUrl;
    CONFIG.BACKGROUND_IMAGES.title = settings.titleBgUrl;
    CONFIG.BACKGROUND_IMAGES.closing = settings.closingBgUrl;
    CONFIG.BACKGROUND_IMAGES.section = settings.sectionBgUrl;
    CONFIG.BACKGROUND_IMAGES.main = settings.mainBgUrl;
    const layout = createLayoutManager(presentation.getPageWidth(), presentation.getPageHeight());
    clearSlideContents(slide, imageUpdateOption);
    const generator = slideGenerators[data.type];
    if (!generator) {
      return { status: 'error', message: `スライドタイプ "${data.type}" に対応する生成関数が見つかりません。` };
    }
    let pageCounterForFooter = 0;
    let tempCounter = 0;
    if (data.type !== 'title' && data.type !== 'closing') {
        pageCounterForFooter = slideIndex + 1;
    }
    generator(slide, data, layout, pageCounterForFooter, settings, imageUpdateOption);
    if (data.notes) {
      const cleanedNotes = cleanSpeakerNotes(data.notes);
      slide.getNotesPage().getSpeakerNotesShape().getText().setText(cleanedNotes);
    } else {
      slide.getNotesPage().getSpeakerNotesShape().getText().setText('');
    }
    return { status: 'success', message: `ページ ${slideIndex + 1} を更新しました。` };
  } catch (e) {
    return { status: 'error', message: `ページ ${slideIndex + 1} の更新中にエラーが発生しました: ${e.message}` };
  }
}
const slideGenerators = {
  title: createTitleSlide,
  section: createSectionSlide,
  content: createContentSlide,
  agenda: createAgendaSlide,
  compare: createCompareSlide,
  process: createProcessSlide,
  processList: createProcessListSlide,
  timeline: createTimelineSlide,
  diagram: createDiagramSlide,
  cycle: createCycleSlide,
  cards: createCardsSlide,
  headerCards: createHeaderCardsSlide,
  table: createTableSlide,
  progress: createProgressSlide,
  quote: createQuoteSlide,
  kpi: createKpiSlide,
  closing: createClosingSlide,
  bulletCards: createBulletCardsSlide,
  faq: createFaqSlide,
  statsCompare: createStatsCompareSlide,
  barCompare: createBarCompareSlide,
  triangle: createTriangleSlide,
  pyramid: createPyramidSlide,
  flowChart: createFlowChartSlide,
  stepUp: createStepUpSlide,
  imageText: createImageTextSlide
};
function createLayoutManager(pageW_pt, pageH_pt) {
  const pxToPt = (px) => px * 0.75;
  const baseW_pt = pxToPt(CONFIG.BASE_PX.W),
    baseH_pt = pxToPt(CONFIG.BASE_PX.H);
  const scaleX = pageW_pt / baseW_pt,
    scaleY = pageH_pt / baseH_pt;
  const getPositionFromPath = (path) => path.split('.').reduce((obj, key) => obj[key], CONFIG.POS_PX);
  return {
    scaleX,
    scaleY,
    pageW_pt,
    pageH_pt,
    pxToPt,
    getRect: (spec) => {
      const pos = typeof spec === 'string' ? getPositionFromPath(spec) : spec;
      let left_px = pos.left;
      if (pos.right !== undefined && pos.left === undefined) {
        left_px = CONFIG.BASE_PX.W - pos.right - pos.width;
      }
      if (left_px === undefined && pos.right === undefined) {
        left_px = 0;
      }
      return {
        left: left_px !== undefined ? pxToPt(left_px) * scaleX : 0,
        top: pos.top !== undefined ? pxToPt(pos.top) * scaleY : 0,
        width: pos.width !== undefined ? pxToPt(pos.width) * scaleX : 0,
        height: pos.height !== undefined ? pxToPt(pos.height) * scaleY : 0,
      };
    }
  };
}
function adjustAreaForSubhead(area, subhead, layout) {
  return area;
}
function createContentCushion(slide, area, settings, layout) {
  if (!area || !area.width || !area.height || area.width <= 0 || area.height <= 0) {
    return;
  }
  const cushionColor = CONFIG.COLORS.background_gray;
  const cushion = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 
    area.left, area.top, area.width, area.height);
  cushion.getFill().setSolidFill(cushionColor, 1.0);
  const border = cushion.getBorder();
  border.setTransparent();
}
function createTitleSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.title, CONFIG.COLORS.background_white, imageUpdateOption);
  if (imageUpdateOption === 'update') {
    const logoRect = layout.getRect('titleSlide.logo');
    try {
      if (CONFIG.LOGOS.header) {
        const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.header);
        if (imageData) {
          const logo = slide.insertImage(imageData);
          const aspect = logo.getHeight() / logo.getWidth();
          logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * aspect);
        }
      }
    } catch (e) {
    }
  }
  const titleRect = layout.getRect('titleSlide.title');
  const newTop = (layout.pageH_pt - titleRect.height) / 2;
  const newWidth = titleRect.width + layout.pxToPt(60);
  const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, newTop, newWidth, titleRect.height);
  setStyledText(titleShape, data.title, {
    size: CONFIG.FONTS.sizes.title,
    bold: true,
    fontType: 'large'
  });
  try {
    titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {
  }
  try {
    const titleText = data.title || '';
    if (titleText.indexOf('\n') === -1) {
      const preCalculatedWidth = (data && typeof data._title_widthPt === 'number') ?
  data._title_widthPt : null;
      if (preCalculatedWidth !== null && preCalculatedWidth < 900) {
        const adjustmentLog = adjustShapeText_External(titleShape, preCalculatedWidth);
      } else if (preCalculatedWidth !== null) {
      } else {
        const adjustmentLog = adjustShapeText_External(titleShape, null);
  }
    } else {
    }
  } catch (e) {
  }
  try {
    const titleTextRange = titleShape.getText();
    if (!titleTextRange.isEmpty()) {
      const firstRun = titleTextRange.getRuns()[0];
      if (firstRun) {
        const currentFontSize = firstRun.getTextStyle().getFontSize();
        if (currentFontSize === 41) {
          titleTextRange.getTextStyle().setFontSize(40);
        }
      }
    }
  } catch (e) {
  }
  if (settings.showDateColumn) {
    const dateRect = layout.getRect('titleSlide.date');
    const dateShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, dateRect.left, dateRect.top, dateRect.width, dateRect.height);
    dateShape.getText().setText(data.date || '');
    applyTextStyle(dateShape.getText(), {
      size: CONFIG.FONTS.sizes.date,
      fontType: 'large'
    });
  }
  if (settings.showBottomBar) {
    drawBottomBar(slide, layout, settings);
  }
  if (__CREDIT_IMAGE_BLOB) {
    drawCreditImage(slide, layout, __CREDIT_IMAGE_BLOB, CREDIT_IMAGE_LINK);
  }
}
function createSectionSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  const imageUrl = CONFIG.BACKGROUND_IMAGES.section || CONFIG.BACKGROUND_IMAGES.main || '';
  let fallbackColor;
  if (imageUrl) {
    fallbackColor = CONFIG.COLORS.background_white;
  } else {
    if (CONFIG.COLORS.background_white.toUpperCase() !== '#FFFFFF') {
      fallbackColor = CONFIG.COLORS.background_white;
    } else {
      fallbackColor = CONFIG.COLORS.background_gray;
    }
  }
  setBackgroundImageFromUrl(slide, layout, imageUrl, fallbackColor, imageUpdateOption);
  __SECTION_COUNTER++;
  const parsedNum = (() => {
    if (Number.isFinite(data.sectionNo)) {
      return Number(data.sectionNo);
    }
    const m = String(data.title || '').match(/^\s*(\d+)[\.．]/);
    return m ? Number(m[1]) : __SECTION_COUNTER;
  })();
  const num = String(parsedNum).padStart(2, '0');
  const ghostRect = layout.getRect('sectionSlide.ghostNum');
  let ghostImageInserted = false;
  if (imageUpdateOption === 'update') {
    if (data.ghostImageBase64 && ghostRect) {
      try {
        const imageData = insertImageFromUrlOrFileId(data.ghostImageBase64);
        if (imageData) {
          const ghostImage = slide.insertImage(imageData);
          const imgWidth = ghostImage.getWidth();
          const imgHeight = ghostImage.getHeight();
          const scale = Math.min(ghostRect.width / imgWidth, ghostRect.height / imgHeight);
          const w = imgWidth * scale;
          const h = imgHeight * scale;
          ghostImage.setWidth(w).setHeight(h)
            .setLeft(ghostRect.left + (ghostRect.width - w) / 2)
            .setTop(ghostRect.top + (ghostRect.height - h) / 2);
          ghostImageInserted = true;
        }
      } catch (e) {
      }
    }
  }
  if (!ghostImageInserted && ghostRect && imageUpdateOption === 'update') {
    const ghost = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, ghostRect.left, ghostRect.top, ghostRect.width, ghostRect.height);
    ghost.getText().setText(num);
    const ghostTextStyle = ghost.getText().getTextStyle();
    ghostTextStyle.setFontFamily(CONFIG.FONTS.family)
      .setFontSize(CONFIG.FONTS.sizes.ghostNum)
      .setBold(true);
    try {
      ghostTextStyle.setForegroundColorWithAlpha(CONFIG.COLORS.ghost_gray, 0.15);
    } catch (e) {
      ghostTextStyle.setForegroundColor(CONFIG.COLORS.ghost_gray);
    }
    try {
      ghost.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
  }
  const titleRect = layout.getRect('sectionSlide.title');
  const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, titleRect.top, titleRect.width, titleRect.height);
  titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  setStyledText(titleShape, data.title, {
    size: CONFIG.FONTS.sizes.sectionTitle,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    fontType: 'large'
  });
  try {
    const titleText = data.title ||
    '';
    if (titleText.indexOf('\n') === -1) {
      const preCalculatedWidthPt = (data && typeof data._title_widthPt === 'number') ?
      data._title_widthPt : null;
      const textAreaWidthPt = titleRect.width;
      if (textAreaWidthPt > 0 && (preCalculatedWidthPt === null || preCalculatedWidthPt < (textAreaWidthPt * 1.4))) {
        if (preCalculatedWidthPt !== null) {
        } else {
        }
        const adjustmentLog = adjustShapeText_External(titleShape, preCalculatedWidthPt);
      } else if (preCalculatedWidthPt !== null) {
      }
    } else {
    }
  } catch (e) {
  }
  addCucFooter(slide, layout, pageNum, settings);
}
function createClosingSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.closing, CONFIG.COLORS.background_white, imageUpdateOption);
  if (imageUpdateOption === 'update') {
    try {
      if (CONFIG.LOGOS.closing) {
        const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.closing);
        if (imageData) {
          const image = slide.insertImage(imageData);
          const imgW_pt = layout.pxToPt(450) * layout.scaleX;
          const aspect = image.getHeight() / image.getWidth();
          image.setWidth(imgW_pt).setHeight(imgW_pt * aspect);
          image.setLeft((layout.pageW_pt - imgW_pt) / 2).setTop((layout.pageH_pt - (imgW_pt * aspect)) / 2);
        }
      }
    } catch (e) {
    }
  }
  if (__CREDIT_IMAGE_BLOB) {
    drawCreditImage(slide, layout, __CREDIT_IMAGE_BLOB, CREDIT_IMAGE_LINK);
  }
}
function createContentSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead, subheadWidthPt); 
  const isAgenda = isAgendaTitle(data.title || '');
  let points = Array.isArray(data.points) ? data.points.slice(0) : [];
  if (isAgenda && points.length === 0) {
    points = buildAgendaFromSlideData();
    if (points.length === 0) {
      points = ['本日の目的', '進め方', '次のアクション'];
    }
  }
  const hasImages = Array.isArray(data.images) && data.images.length > 0;
  const isTwo = !!(data.twoColumn || data.columns);
  if ((isTwo && (data.columns || points)) || (!isTwo && points && points.length > 0)) {
    if (isTwo) {
      let L = [], R = [];
      if (Array.isArray(data.columns) && data.columns.length === 2) {
        L = data.columns[0] || [];
        R = data.columns[1] || [];
      } else {
        const mid = Math.ceil(points.length / 2);
        L = points.slice(0, mid);
        R = points.slice(mid);
      }
      const baseLeftRect = layout.getRect('contentSlide.twoColLeft');
      const baseRightRect = layout.getRect('contentSlide.twoColRight');
      const adjustedLeftRect = adjustAreaForSubhead(baseLeftRect, data.subhead, layout);
      const adjustedRightRect = adjustAreaForSubhead(baseRightRect, data.subhead, layout);
      const leftRect = offsetRect(adjustedLeftRect, 0, dy);
      const rightRect = offsetRect(adjustedRightRect, 0, dy);
      createContentCushion(slide, leftRect, settings, layout);
      createContentCushion(slide, rightRect, settings, layout);
      const padding = layout.pxToPt(20);
      const leftTextRect = { left: leftRect.left + padding, top: leftRect.top + padding, width: leftRect.width - (padding * 2), height: leftRect.height - (padding * 2) };
      const rightTextRect = { left: rightRect.left + padding, top: rightRect.top + padding, width: rightRect.width - (padding * 2), height: rightRect.height - (padding * 2) };
      const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftTextRect.left, leftTextRect.top, leftTextRect.width, leftTextRect.height);
      const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightTextRect.left, rightTextRect.top, rightTextRect.width, rightTextRect.height);
      setBulletsWithInlineStyles(leftShape, L);
      setBulletsWithInlineStyles(rightShape, R);
      setBoldTextSize(leftShape, 16);
      setBoldTextSize(rightShape, 16);
      try {
        const leftAdjustmentLog = adjustShapeText_External(leftShape, null);
      } catch(e) {
      }
      try {
        const rightAdjustmentLog = adjustShapeText_External(rightShape, null);
      } catch(e) {
      }
    } else {
      const baseBodyRect = layout.getRect('contentSlide.body');
      const adjustedBodyRect = adjustAreaForSubhead(baseBodyRect, data.subhead, layout);
      const bodyRect = offsetRect(adjustedBodyRect, 0, dy);
      createContentCushion(slide, bodyRect, settings, layout);
      if (isAgenda) {
        drawNumberedItems(slide, layout, bodyRect, points, settings);
      } else {
        const padding = layout.pxToPt(20);
        const textRect = { left: bodyRect.left + padding, top: bodyRect.top + padding, width: bodyRect.width - (padding * 2), height: bodyRect.height - (padding * 2) };
        const bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
        setBulletsWithInlineStyles(bodyShape, points);
        setBoldTextSize(bodyShape, 16);
        try {
          bodyShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch(e) {
        }
        try {
          const bodyAdjustmentLog = adjustShapeText_External(bodyShape, null);
        } catch(e) { 
        }
      }
    }
  }
  if (hasImages && !points.length && !isTwo) {
    const baseArea = layout.getRect('contentSlide.body');
    const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
    const area = offsetRect(adjustedArea, 0, dy);
    createContentCushion(slide, area, settings, layout);
    renderImagesInArea(slide, layout, area, normalizeImages(data.images), imageUpdateOption);
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createCompareSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead, subheadWidthPt);
  const baseArea = layout.getRect('cardsSlide.gridArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const cols = 2;
  const gap = layout.pxToPt(16);
  const cardW = (area.width - gap * (cols - 1)) / cols;
  const cardH = area.height;
  const leftBoxLeft = area.left;
  const leftBoxTop = area.top;
  const leftBoxWidth = cardW;
  const leftBoxHeight = cardH;
  const rightBoxLeft = area.left + cardW + gap;
  const rightBoxTop = area.top;
  const rightBoxWidth = cardW;
  const rightBoxHeight = cardH;
  drawCompareBox(slide, layout, leftBoxLeft, leftBoxTop, leftBoxWidth, leftBoxHeight, data.leftTitle || '選択肢A', data.leftItems || [], settings, true);
  drawCompareBox(slide, layout, rightBoxLeft, rightBoxTop, rightBoxWidth, rightBoxHeight, data.rightTitle || '選択肢B', data.rightItems || [], settings, false);
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createProcessSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout, imageUpdateOption);
  drawStandardTitleHeader(slide, layout, 'processSlide', data.title, settings, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead, subheadWidthPt);
  const baseArea = layout.getRect('processSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const steps = Array.isArray(data.steps) ? data.steps.slice(0, 4) : [];
  if (steps.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const processBodyBgColor = generateTintedGray(settings.primaryColor, 30, 94);
  const n = steps.length;
  let boxHPx, arrowHPx, fontSize;
  if (n <= 2) {
    boxHPx = 100;
    arrowHPx = 25;
    fontSize = 16;
  } else if (n === 3) {
    boxHPx = 80;
    arrowHPx = 20;
    fontSize = 16;
  } else {
    boxHPx = 65;
    arrowHPx = 15;
    fontSize = 14;
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
    const header = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, currentY, headerWPt, boxHPt);
    header.getFill().setSolidFill(processColors[i]);
    header.getBorder().setTransparent();
    setStyledText(header, `STEP ${i + 1}`, {
      size: fontSize,
      bold: true,
      color: CONFIG.COLORS.background_gray,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      header.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    const body = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, bodyLeft, currentY, bodyWPt, boxHPt);
    body.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    body.getBorder().setTransparent();
    const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, bodyLeft + layout.pxToPt(20), currentY, bodyWPt - layout.pxToPt(40), boxHPt);
    setStyledText(textShape, cleanText, {
      size: fontSize
    });
    try {
      textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    currentY += boxHPt;
    if (i < n - 1) {
      const arrowLeft = area.left + headerWPt / 2 - layout.pxToPt(8);
      const arrow = slide.insertShape(SlidesApp.ShapeType.DOWN_ARROW, arrowLeft, currentY, layout.pxToPt(16), arrowHPt);
      arrow.getFill().setSolidFill(CONFIG.COLORS.process_arrow);
      arrow.getBorder().setTransparent();
      currentY += arrowHPt;
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createProcessListSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'processSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead, subheadWidthPt); 
  const area = offsetRect(layout.getRect('processSlide.area'), 0, dy);
  const steps = Array.isArray(data.steps) ? data.steps : [];
  if (steps.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const n = Math.max(1, steps.length);
  const topPadding = layout.pxToPt(30);
  const bottomPadding = layout.pxToPt(10);
  const drawableHeight = area.height - topPadding - bottomPadding; 
  const maxGapY_pt = layout.pxToPt(70);
  let idealGapY = Infinity;
  if (n > 1) {
    idealGapY = drawableHeight / (n - 1);
  }
  const finalGapY = Math.min(idealGapY, maxGapY_pt);
  const cx = area.left + layout.pxToPt(44); 
  const top0 = area.top + topPadding;
  const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx - layout.pxToPt(1), top0 + layout.pxToPt(6), layout.pxToPt(2), finalGapY * (n - 1));
  line.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
  line.getBorder().setTransparent();
  for (let i = 0; i < n; i++) {
    const cy = top0 + finalGapY * i;
    const sz = layout.pxToPt(28);
    const numBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx - sz/2, cy - sz/2, sz, sz);
    numBox.getFill().setSolidFill(settings.primaryColor);
    numBox.getBorder().setTransparent();
    const num = numBox.getText(); num.setText(String(i + 1));
    applyTextStyle(num, { size: 12, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
    let cleanText = String(steps[i] || '');
    cleanText = cleanText.replace(/^\s*\d+[\.\s]*/, '');
    const txt = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32));
    setStyledText(txt, cleanText, { size: CONFIG.FONTS.sizes.processStep, fontType: 'large' });
    try { txt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createTimelineSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'timelineSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'timelineSlide', data.subhead, subheadWidthPt); 
  const baseArea = layout.getRect('timelineSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const milestones = Array.isArray(data.milestones) ? data.milestones : [];
  if (milestones.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const inner = layout.pxToPt(80),
    baseY = area.top + area.height * 0.50;
  const leftX = area.left + inner,
    rightX = area.left + area.width - inner;
  const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, leftX, baseY, rightX, baseY);
  line.getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
  line.setWeight(2);
  const dotR = layout.pxToPt(10);
  const gap = (milestones.length > 1) ? (rightX - leftX) / (milestones.length - 1) : 0;
  const cardW_pt = layout.pxToPt(180);
  const vOffset = layout.pxToPt(40);
  const headerHeight = layout.pxToPt(28);
  const bodyHeight = layout.pxToPt(80);
  const cardPadding = layout.pxToPt(8);
  milestones.forEach((m, i) => {
    const x = leftX + gap * i;
    const isAbove = i % 2 === 0;
    const dateText = String(m.date || '');
    const labelText = String(m.label || '');
    const cardH_pt = headerHeight + bodyHeight;
    const cardLeft = x - (cardW_pt / 2);
    const cardTop = isAbove ? (baseY - vOffset - cardH_pt) : (baseY + vOffset);
    const timelineColors = generateTimelineCardColors(settings.primaryColor, milestones.length);
    const headerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cardLeft, cardTop, cardW_pt, headerHeight);
    headerShape.getFill().setSolidFill(timelineColors[i]);
    headerShape.getBorder().getLineFill().setSolidFill(timelineColors[i]);
    const bodyShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cardLeft, cardTop + headerHeight, cardW_pt, bodyHeight);
    bodyShape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    bodyShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const connectorY_start = isAbove ? (cardTop + cardH_pt) : baseY;
    const connectorY_end = isAbove ? baseY : cardTop;
    const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, connectorY_start, x, connectorY_end);
    connector.getLineFill().setSolidFill(CONFIG.COLORS.neutral_gray);
    connector.setWeight(1);
    const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - dotR / 2, baseY - dotR / 2, dotR, dotR);
    dot.getFill().setSolidFill(timelineColors[i]);
    dot.getBorder().setTransparent();
    const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      cardLeft, cardTop, 
      cardW_pt, headerHeight);
    setStyledText(headerTextShape, dateText, {
      size: CONFIG.FONTS.sizes.body,
      bold: true,
      color: CONFIG.COLORS.background_gray,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    setParenthesizedTextSize(headerTextShape, 8);
    try {
      headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    let bodyFontSize = CONFIG.FONTS.sizes.body;
    const textLength = labelText.length;
    if (textLength > 40) bodyFontSize = 10;
    else if (textLength > 30) bodyFontSize = 11;
    else if (textLength > 20) bodyFontSize = 12;
    const bodyTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      cardLeft, cardTop + headerHeight, 
      cardW_pt, bodyHeight);
    setStyledText(bodyTextShape, labelText, {
      size: bodyFontSize,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      bodyTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
  });
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createDiagramSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'diagramSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'diagramSlide', data.subhead, subheadWidthPt); 
  const lanes = Array.isArray(data.lanes) ? data.lanes : [];
  const baseArea = layout.getRect('diagramSlide.lanesArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const px = (p) => layout.pxToPt(p);
  const {
    laneGap_px,
    lanePad_px,
    laneTitle_h_px,
    cardGap_px,
    cardMin_h_px,
    cardMax_h_px,
    arrow_h_px,
    arrowGap_px
  } = CONFIG.DIAGRAM;
  const n = Math.max(1, lanes.length);
  const laneW = (area.width - px(laneGap_px) * (n - 1)) / n;
  const cardBoxes = [];
  for (let j = 0; j < n; j++) {
    const lane = lanes[j] || {
      title: '',
      items: []
    };
    const left = area.left + j * (laneW + px(laneGap_px));
    const lt = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, area.top, laneW, px(laneTitle_h_px));
    lt.getFill().setSolidFill(settings.primaryColor);
    lt.getBorder().getLineFill().setSolidFill(settings.primaryColor);
    setStyledText(lt, lane.title || '', {
      size: CONFIG.FONTS.sizes.laneTitle,
      bold: true,
      color: CONFIG.COLORS.background_gray,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      lt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    const laneBodyTop = area.top + px(laneTitle_h_px),
      laneBodyHeight = area.height - px(laneTitle_h_px);
    const laneBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, laneBodyTop, laneW, laneBodyHeight);
    laneBg.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    laneBg.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);
    const items = Array.isArray(lane.items) ? lane.items : [];
    const availH = laneBodyHeight - px(lanePad_px) * 2,
      rows = Math.max(1, items.length);
    const idealH = (availH - px(cardGap_px) * (rows - 1)) / rows;
    const cardH = Math.max(px(cardMin_h_px), Math.min(px(cardMax_h_px), idealH));
    const firstTop = laneBodyTop + px(lanePad_px) + Math.max(0, (availH - (cardH * rows + px(cardGap_px) * (rows - 1))) / 2);
    cardBoxes[j] = [];
    for (let i = 0; i < rows; i++) {
      const cardTop = firstTop + i * (cardH + px(cardGap_px));
      const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, left + px(lanePad_px), cardTop, laneW - px(lanePad_px) * 2, cardH);
      card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
      card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
      setStyledText(card, items[i] || '', {
        size: CONFIG.FONTS.sizes.body
      });
      try {
        card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {}
      cardBoxes[j][i] = {
        left: left + px(lanePad_px),
        top: cardTop,
        width: laneW - px(lanePad_px) * 2,
        height: cardH
      };
    }
  }
  const maxRows = Math.max(0, ...cardBoxes.map(a => a ? a.length : 0));
  for (let j = 0; j < n - 1; j++) {
    for (let i = 0; i < maxRows; i++) {
      if (cardBoxes[j] && cardBoxes[j][i] && cardBoxes[j + 1] && cardBoxes[j + 1][i]) {
        drawArrowBetweenRects(slide, cardBoxes[j][i], cardBoxes[j + 1][i], px(arrow_h_px), px(arrowGap_px), settings);
      }
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createCycleSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead, subheadWidthPt); 
  const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
  const items = Array.isArray(data.items) && data.items.length === 4 ? data.items : [];
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const textLengths = items.map(item => {
    const labelLength = (item.label || '').length;
    const subLabelLength = (item.subLabel || '').length;
    return labelLength + subLabelLength;
  });
  const maxLength = Math.max(...textLengths);
  const avgLength = textLengths.reduce((sum, len) => sum + len, 0) / textLengths.length;
  const centerX = area.left + area.width / 2;
  const centerY = area.top + area.height / 2;
  const radiusX = area.width / 3.2;
  const radiusY = area.height / 2.6;
  const maxCardW = Math.min(layout.pxToPt(220), radiusX * 0.8);
  const maxCardH = Math.min(layout.pxToPt(100), radiusY * 0.6);
  let cardW, cardH, fontSize;
  if (maxLength > 25 || avgLength > 18) {
    cardW = Math.min(layout.pxToPt(230), maxCardW);
    cardH = Math.min(layout.pxToPt(105), maxCardH);
    fontSize = 13;
  } else if (maxLength > 15 || avgLength > 10) {
    cardW = Math.min(layout.pxToPt(215), maxCardW);
    cardH = Math.min(layout.pxToPt(95), maxCardH);
    fontSize = 14;
  } else {
    cardW = layout.pxToPt(200);
    cardH = layout.pxToPt(90);
    fontSize = 16;
  }
  if (data.centerText) {
    const centerTextBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, centerX - layout.pxToPt(100), centerY - layout.pxToPt(50), layout.pxToPt(200), layout.pxToPt(100));
    setStyledText(centerTextBox, data.centerText, { size: 20, bold: true, align: SlidesApp.ParagraphAlignment.CENTER, color: CONFIG.COLORS.text_primary });
    try { centerTextBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
  }
  const positions = [
    { x: centerX + radiusX, y: centerY },
    { x: centerX, y: centerY + radiusY },
    { x: centerX - radiusX, y: centerY },
    { x: centerX, y: centerY - radiusY }
  ];
  positions.forEach((pos, i) => {
    const cardX = pos.x - cardW / 2;
    const cardY = pos.y - cardH / 2;
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardW, cardH);
    card.getFill().setSolidFill(settings.primaryColor);
    card.getBorder().setTransparent();
    const item = items[i] || {};
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
    } catch (e) {}
  });
  const arrowRadiusX = radiusX * 0.75;
  const arrowRadiusY = radiusY * 0.80;
  const arrowSize = layout.pxToPt(80);
  const arrowPositions = [
    { left: centerX + arrowRadiusX, top: centerY - arrowRadiusY, rotation: 90 },
    { left: centerX + arrowRadiusX, top: centerY + arrowRadiusY, rotation: 180 },
    { left: centerX - arrowRadiusX, top: centerY + arrowRadiusY, rotation: 270 },
    { left: centerX - arrowRadiusX, top: centerY - arrowRadiusY, rotation: 0 }
  ];
  arrowPositions.forEach(pos => {
    const arrow = slide.insertShape(SlidesApp.ShapeType.BENT_ARROW, pos.left - arrowSize / 2, pos.top - arrowSize / 2, arrowSize, arrowSize);
    arrow.getFill().setSolidFill(CONFIG.COLORS.ghost_gray);
    arrow.getBorder().setTransparent();
    arrow.setRotation(pos.rotation);
  });
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createCardsSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout, imageUpdateOption);
  drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title, settings, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
 data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead, subheadWidthPt); 
  const baseArea = layout.getRect('cardsSlide.gridArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  const cols = Math.min(3, Math.max(2, Number(data.columns) || (items.length <= 4 ? 2 : 3)));
  const gap = layout.pxToPt(16),
    rows = Math.ceil(items.length / cols);
  const cardW = (area.width - gap * (cols - 1)) / cols,
    cardH = Math.max(layout.pxToPt(92), (area.height - gap * (rows - 1)) / rows);
  for (let idx = 0; idx < items.length; idx++) {
    const r = Math.floor(idx / cols),
      c = idx % cols;
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left + c * (cardW + gap), area.top + r * (cardH + gap), cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const obj = items[idx];
    if (typeof obj === 'string') {
      setStyledText(card, obj, {
        size: CONFIG.FONTS.sizes.body
      });
    } else {
      const title = String(obj.title || ''),
        desc = String(obj.desc || '');
      if (title && desc) {
        const combined = `${title}\n\n${desc}`;
        setStyledText(card, combined, {
          size: CONFIG.FONTS.sizes.body
        });
        try {
          card.getText().getRange(0, title.length).getTextStyle().setBold(true);
        } catch (e) {}
      } else if (title) {
        setStyledText(card, title, {
          size: CONFIG.FONTS.sizes.body,
          bold: true
        });
      } else {
        setStyledText(card, desc, {
          size: CONFIG.FONTS.sizes.body
        });
      }
    }
    try {
      card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createHeaderCardsSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead, subheadWidthPt); 
  const baseArea = layout.getRect('cardsSlide.gridArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  const cols = Math.min(3, Math.max(2, Number(data.columns) || (items.length <= 4 ? 2 : 3)));
  const gap = layout.pxToPt(16),
    rows = Math.ceil(items.length / cols);
  const cardW = (area.width - gap * (cols - 1)) / cols,
    cardH = Math.max(layout.pxToPt(92), (area.height - gap * (rows - 1)) / rows);
  for (let idx = 0; idx < items.length; idx++) {
    const r = Math.floor(idx / cols),
      c = idx % cols;
    const left = area.left + c * (cardW + gap),
      top = area.top + r * (cardH + gap);
    const titleText = String(items[idx].title || ''),
      descText = String(items[idx].desc || '');
    const headerHeight = layout.pxToPt(40);
    const bodyShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top + headerHeight, cardW, cardH - headerHeight);
    bodyShape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    bodyShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const headerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, cardW, headerHeight);
    headerShape.getFill().setSolidFill(settings.primaryColor);
    headerShape.getBorder().getLineFill().setSolidFill(settings.primaryColor);
    const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, cardW, headerHeight);
    setStyledText(headerTextShape, titleText, {
      size: CONFIG.FONTS.sizes.body,
      bold: true,
      color: CONFIG.COLORS.background_gray,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    const bodyTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + layout.pxToPt(12), top + headerHeight, cardW - layout.pxToPt(24), cardH - headerHeight);
    setStyledText(bodyTextShape, descText, {
      size: CONFIG.FONTS.sizes.body,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      setBoldTextSize(bodyTextShape, 16);
    } catch (e) {
    }
    try {
      bodyTextShape.getText().getParagraphStyle().setLineSpacing(115);
    } catch (e) {
    }
    try {
      bodyTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    try {
      const adjustmentLog = adjustShapeText_External(bodyTextShape, null);
    } catch(e) {
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createTableSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'tableSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'tableSlide', data.subhead, subheadWidthPt);
  const baseArea = layout.getRect('tableSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const headers = Array.isArray(data.headers) ? data.headers : [];
  const rows = Array.isArray(data.rows) ?
  data.rows : [];
  try {
    if (headers.length > 0) {
      const numRowsTotal = rows.length + 1;
      const table = slide.insertTable(numRowsTotal, headers.length, area.left, area.top, area.width, area.height);
      const estimatedRowHeight = area.height / numRowsTotal;
      for (let c = 0; c < headers.length; c++) {
        const cell = table.getCell(0, c);
        const cellWidth = table.getColumn(c).getWidth();
        cell.getFill().setSolidFill(CONFIG.COLORS.table_header_bg);
        setStyledText(cell, String(headers[c] || ''), {
          bold: true,
          color: CONFIG.COLORS.text_small_font,
          align: SlidesApp.ParagraphAlignment.CENTER
        });
        try {
          cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {}
        try {
          adjustShapeText_External(cell, null, cellWidth, estimatedRowHeight);
        } catch (adjustError) {
        }
      }
      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < headers.length; c++) {
          const cell = table.getCell(r + 1, c);
          const cellWidth = table.getColumn(c).getWidth();
          cell.getFill().setSolidFill(CONFIG.COLORS.background_gray);
          setStyledText(cell, String((rows[r] || [])[c] || ''), {
            align: SlidesApp.ParagraphAlignment.CENTER
          });
          try {
            cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {}
          try {
            adjustShapeText_External(cell, null, cellWidth, estimatedRowHeight);
          } catch (adjustError) {
          }
        }
      }
    }
  } catch (e) {
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createProgressSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'progressSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'progressSlide', data.subhead, subheadWidthPt); 
  const baseArea = layout.getRect('progressSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const items = Array.isArray(data.items) ? data.items.slice(0, 4) : [];
  const n = Math.max(1, items.length);
  const cardGap = layout.pxToPt(12);
  const cardHeight = Math.max(layout.pxToPt(50), (area.height - cardGap * (n - 1)) / n);
  const cardPadding = layout.pxToPt(15);
  const barHeight = layout.pxToPt(12);
  const percentHeight = layout.pxToPt(30);
  const percentWidth = layout.pxToPt(120);
  for (let i = 0; i < n; i++) {
    const cardTop = area.top + i * (cardHeight + cardGap);
    const p = Math.max(0, Math.min(100, Number(items[i].percent || 0)));
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 
      area.left, cardTop, area.width, cardHeight);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const labelHeight = layout.pxToPt(20);
    const labelWidth = area.width - percentWidth - cardPadding * 3;
    const label = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      area.left + cardPadding, cardTop + cardPadding, 
      labelWidth, labelHeight);
    setStyledText(label, String(items[i].label || ''), {
      size: CONFIG.FONTS.sizes.body,
      bold: true,
      align: SlidesApp.ParagraphAlignment.LEFT
    });
    try {
    label.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
    const pct = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      area.left + area.width - percentWidth - cardPadding, 
      cardTop + cardPadding - layout.pxToPt(2), 
      percentWidth, percentHeight);
    setStyledText(pct, `${p}%`, {
      size: 20,
      bold: true,
      color: settings.primaryColor,
      align: SlidesApp.ParagraphAlignment.RIGHT
    });
    try {
    pct.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
    const barTop = cardTop + cardHeight - cardPadding - barHeight;
    const barWidth = area.width - cardPadding * 2;
    const barBG = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 
      area.left + cardPadding, barTop, barWidth, barHeight);
    barBG.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
    barBG.getBorder().setTransparent();
    if (p > 0) {
      const filledBarWidth = Math.max(layout.pxToPt(6), barWidth * (p / 100));
      const barFG = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 
        area.left + cardPadding, barTop, filledBarWidth, barHeight);
      barFG.getFill().setSolidFill(settings.primaryColor);
      barFG.getBorder().setTransparent();
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createQuoteSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'quoteSlide', data.title || '引用', settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'quoteSlide', data.subhead, subheadWidthPt); 
  const baseTop = 120;
  const subheadHeight = data.subhead ? layout.pxToPt(40) : 0;
  const margin = layout.pxToPt(10);
  const area = offsetRect(layout.getRect({
    left: 40,
    top: baseTop + subheadHeight + margin,
    width: 880,
    height: 320 - subheadHeight - margin
  }), 0, dy);
  const bgCard = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left, area.top, area.width, area.height);
  bgCard.getFill().setSolidFill(CONFIG.COLORS.background_white);
  const border = bgCard.getBorder();
  border.getLineFill().setSolidFill(CONFIG.COLORS.card_border);
  border.setWeight(2);
  const padding = layout.pxToPt(40);
  const textLeft = area.left + padding,
    textTop = area.top + padding;
  const textWidth = area.width - (padding * 2),
    textHeight = area.height - (padding * 2);
  const quoteTextHeight = textHeight - layout.pxToPt(30);
  const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textLeft, textTop, textWidth, quoteTextHeight);
  setStyledText(textShape, data.text || '', {
    size: 24,
    align: SlidesApp.ParagraphAlignment.CENTER,
    //color: CONFIG.COLORS.text_primary
    fontType: 'large'
  });
  try {
    textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
  const authorTop = textTop + quoteTextHeight;
  const authorShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textLeft, authorTop, textWidth, layout.pxToPt(30));
  setStyledText(authorShape, `— ${data.author || ''}`, {
    size: 16,
    color: CONFIG.COLORS.neutral_gray,
    align: SlidesApp.ParagraphAlignment.END
  });
  try {
    authorShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createKpiSlide(slide, data, layout, pageNum, settings, imageUpdateOption) {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'kpiSlide', data.title || '主要指標', settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ? data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'kpiSlide', data.subhead, subheadWidthPt);
  const items = (Array.isArray(data.items) ? data.items.slice(0, 8) : []);
  const n = items.length;
  if (n === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const gap = layout.pxToPt(16);
  const pad = layout.pxToPt(15);
  const originalCardH_pt = layout.pxToPt(240);
  if (n <= 4) {
    const baseArea = layout.getRect('kpiSlide.gridArea');
const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
    let finalDy = dy;
    if (dy === 0) {
        finalDy = layout.pxToPt(36); 
    }
    const area = offsetRect(adjustedArea, 0, finalDy);
    const cols = n;
    const cardW = (area.width - gap * (cols - 1)) / cols;
    const cardH = originalCardH_pt;
    for (let idx = 0; idx < n; idx++) {
      const left = area.left + idx * (cardW + gap);
      const top = area.top;
      drawKpiCard(slide, layout, items[idx], left, top, cardW, cardH, pad, 1.0);
    }
  } else {
    const baseArea = layout.getRect('kpiSlide.gridArea');
    const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
    const area = offsetRect(adjustedArea, 0, 0); 
    const rows = 2;
    const cols = (n <= 6) ? 3 : 4; 
    const cardW = (area.width - gap * (cols - 1)) / cols;
    const cardH = (area.height - gap * (rows - 1)) / rows;
    const scaleY = cardH / originalCardH_pt; 
    for (let idx = 0; idx < n; idx++) {
      const c = idx % cols;
      const r = Math.floor(idx / cols);
      const left = area.left + c * (cardW + gap);
      const top = area.top + r * (cardH + gap);
      drawKpiCard(slide, layout, items[idx], left, top, cardW, cardH, pad, scaleY);
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function drawKpiCard(slide, layout, item, left, top, cardW, cardH, pad, scaleY) {
  const card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, cardW, cardH);
  card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
  card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
  const extraWidthPt = layout.pxToPt(30);
  const labelLeft = (left + pad) - (extraWidthPt / 2);
  const labelWidth = (cardW - pad * 2) + extraWidthPt;
  const labelTop = top + layout.pxToPt(25) * scaleY;
  const labelHeight = layout.pxToPt(35) * scaleY;
  const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, labelLeft, labelTop, labelWidth, labelHeight);
  setStyledText(labelShape, item.label || 'KPI', {
    size: 14,
    color: CONFIG.COLORS.neutral_gray,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  try {
    labelShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {
  }
  try {
    const adjustmentResult = adjustShapeText_External(labelShape, null);
  } catch(e) {
  }
  const valueLeft = left + pad;
  const valueWidth = cardW - pad * 2;
  const valueTop = top + layout.pxToPt(80) * scaleY;
  const valueHeight = layout.pxToPt(80) * scaleY;
  const valueShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, valueLeft, valueTop, valueWidth, valueHeight);
  const valueText = item.value || '0';
  const textRange = valueShape.getText().setText(valueText);
  applyTextStyle(textRange, { size: 32, bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
  const largeChars = '0123456789+-.,:';
  const smallerSize = 16;
  let containsNumericOrSymbol = false;
  for (let i = 0; i < valueText.length; i++) {
    if (largeChars.includes(valueText[i])) {
      containsNumericOrSymbol = true;
      break;
    }
  }
  if (containsNumericOrSymbol) {
    try {
      for (let i = 0; i < valueText.length; i++) {
        const char = valueText[i];
        if (!largeChars.includes(char)) {
          const charRange = textRange.getRange(i, i + 1);
          charRange.getTextStyle().setFontSize(smallerSize);
        }
      }
    } catch (e) {  }
  }
  try { valueShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
  try {
    const adjustmentResult = adjustShapeText_External(valueShape, null);
    const textRange = valueShape.getText();
    const valueText = textRange.asString();
    const containsSpecificChars = /[0-9+-]/.test(valueText);
    if (containsSpecificChars) {
      const currentStyle = textRange.getTextStyle();
      if (currentStyle && currentStyle.getFontSize() !== 32) {
        currentStyle.setFontSize(32);
        try {
          for (let i = 0; i < valueText.length; i++) {
            const char = valueText[i];
            if (!largeChars.includes(char)) {
              const charRange = textRange.getRange(i, i + 1);
              charRange.getTextStyle().setFontSize(smallerSize);
            }
          }
        } catch (e) {  }
      }
    } else {
      const firstRun = textRange.getRuns()[0];
      if (firstRun) {
        const currentSize = firstRun.getTextStyle().getFontSize();
        if (typeof currentSize === 'number' && currentSize < 16) {
          textRange.getTextStyle().setFontSize(16);
        }
      }
    }
  } catch(e) {
  }
  const changeLeft = left + pad;
  const changeWidth = cardW - pad * 2;
  const changeTop = top + layout.pxToPt(180) * scaleY;
  const changeHeight = layout.pxToPt(40) * scaleY;
  const changeShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, changeLeft, changeTop, changeWidth, changeHeight);
  let changeColor = CONFIG.COLORS.text_primary;
  if (item.status === 'bad') changeColor = '#d93025';
  if (item.status === 'good') changeColor = '#1e8e3e';
  if (item.status === 'neutral') changeColor = CONFIG.COLORS.neutral_gray;
  setStyledText(changeShape, item.change || '', {
    size: 14,
    color: changeColor,
    bold: true,
    align: SlidesApp.ParagraphAlignment.END
  });
  try {
    adjustShapeText_External(changeShape, null);
  } catch(e) {
  }
}
function createBulletCardsSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead, subheadWidthPt); 
  const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
  const items = Array.isArray(data.items) ? data.items.slice(0, 3) : [];
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const gap = layout.pxToPt(16);
  const cardHeight = (area.height - gap * (items.length - 1)) / items.length;
  for (let i = 0; i < items.length; i++) {
    const card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top + i * (cardHeight + gap), area.width, cardHeight);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const padding = layout.pxToPt(20);
    const title = String(items[i].title || '');
    const desc = String(items[i].desc || '');
    if (title && desc) {
      const titleHeight = layout.pxToPt(16 + 4);
      const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
        area.left + padding - layout.pxToPt(10),
        area.top + i * (cardHeight + gap) + layout.pxToPt(12),
        area.width - padding * 2,
        titleHeight
      );
      setStyledText(titleShape, title, {
        size: 16,
        bold: true
      });
      try {
        titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {}
      const descTop = area.top + i * (cardHeight + gap) + layout.pxToPt(12) + titleHeight + layout.pxToPt(8);
      const descHeight = cardHeight - layout.pxToPt(12) - titleHeight - layout.pxToPt(8) - layout.pxToPt(12);
      let baseDescFontSize = CONFIG.FONTS.sizes.body;
      if (desc.length > 100) {
        baseDescFontSize = 12;
      } else if (desc.length > 80) {
        baseDescFontSize = 13;
      }
      const descShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
        area.left + padding,
        descTop,
        area.width - padding * 2,
        descHeight
      );
      setStyledText(descShape, desc, {
        size: baseDescFontSize
      });
      setBoldTextSize(descShape, 16);
      try {
        descShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {}
    } else {
      const singleText = title || desc;
      const isTitleOnly = !!title;
      const singleFontSize = isTitleOnly ? 16 : CONFIG.FONTS.sizes.body;
      const singleHeight = layout.pxToPt(singleFontSize + 8);
      const shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
        area.left + padding - layout.pxToPt(10),
        area.top + i * (cardHeight + gap) + (cardHeight - singleHeight) / 2,
        area.width - padding * 2,
        singleHeight
      );
      setStyledText(shape, singleText, {
        size: singleFontSize,
        bold: isTitleOnly
      });
      if (!isTitleOnly) {
        setBoldTextSize(shape, 16);
      }
      try {
        shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {}
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createAgendaSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout, imageUpdateOption);
  drawStandardTitleHeader(slide, layout, 'processSlide', data.title, settings, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead, subheadWidthPt); 
  let effectiveDy = dy;
  if (dy > 0) {
    effectiveDy = dy - layout.pxToPt(20); 
  }
  const area = offsetRect(layout.getRect('processSlide.area'), 0, effectiveDy);
  let items = Array.isArray(data.items) ? data.items : [];
  if (items.length === 0) {
    items = buildAgendaFromSlideData();
    if (items.length === 0) {
      items = ['本日の目的', '進め方', '次のアクション'];
    }
  }
  const n = Math.max(1, items.length);
  const topPadding = layout.pxToPt(30);
  const bottomPadding = layout.pxToPt(10);
  const drawableHeight = area.height - topPadding - bottomPadding;
  const maxGapY_pt = layout.pxToPt(60);
  let idealGapY = Infinity;
  if (n > 1) {
    idealGapY = drawableHeight / (n - 1);
  }
  const finalGapY = Math.min(idealGapY, maxGapY_pt);
  const top0 = area.top + topPadding;
  const cx = area.left + layout.pxToPt(44);
  for (let i = 0; i < n; i++) {
    const cy = top0 + finalGapY * i;
    const sz = layout.pxToPt(28);
    const numBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx - sz/2, cy - sz/2, sz, sz);
    numBox.getFill().setSolidFill(settings.primaryColor);
    numBox.getBorder().setTransparent();
    const num = numBox.getText(); 
    num.setText(String(i + 1));
    applyTextStyle(num, { 
      size: 12, 
      bold: true, 
      color: CONFIG.COLORS.background_gray, 
      align: SlidesApp.ParagraphAlignment.CENTER 
    });
    let cleanText = String(items[i] || '');
    cleanText = cleanText.replace(/^\s*\d+[\.\s]*/, '');
    const txt = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32));
    setStyledText(txt, cleanText, {
      size: CONFIG.FONTS.sizes.processStep,
      bold: true,                        
      fontType: 'large'
    });
    try { 
      txt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); 
    } catch(e){}
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createFaqSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title || 'よくあるご質問', settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead, subheadWidthPt); 
  const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
  const items = Array.isArray(data.items) ? data.items.slice(0, 4) : [];
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  let currentY = area.top;
  const cardGap = layout.pxToPt(12);
  const totalGaps = cardGap * (items.length - 1);
  const availableHeight = area.height - totalGaps;
  const cardHeight = availableHeight / items.length;
  items.forEach((item, index) => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 
      area.left, currentY, area.width, cardHeight);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    card.getBorder().setWeight(1);
    let cardPadding, qAreaRatio, qAGap;
    if (items.length <= 2) {
      cardPadding = layout.pxToPt(16);
      qAreaRatio = 0.30;
      qAGap = layout.pxToPt(6);
    } else if (items.length === 3) {
      cardPadding = layout.pxToPt(12);
      qAreaRatio = 0.35;
      qAGap = layout.pxToPt(4);
    } else {
      cardPadding = layout.pxToPt(8);
      qAreaRatio = 0.40;
      qAGap = layout.pxToPt(2);
    }
    const baseFontSize = items.length >= 4 ? 12 : 14;
    const availableHeight = cardHeight - cardPadding * 2;
    const qAreaHeight = Math.floor(availableHeight * qAreaRatio);
    const aAreaHeight = availableHeight - qAreaHeight - qAGap;
    const qTop = currentY + cardPadding;
    const qText = item.q || '';
    const qParsed = parseInlineStyles(qText);
    const qFullText = `Q. ${qParsed.output}`;
    const qTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      area.left + cardPadding, qTop, 
      area.width - cardPadding * 2, qAreaHeight - layout.pxToPt(2));
    qTextShape.getFill().setTransparent();
    qTextShape.getBorder().setTransparent();
    const qTextRange = qTextShape.getText().setText(qFullText);
    applyTextStyle(qTextRange, {
      size: baseFontSize,
      color: CONFIG.COLORS.text_small_font,
      align: SlidesApp.ParagraphAlignment.LEFT
    });
    try {
      const qPrefixRange = qTextRange.getRange(0, 2);
      qPrefixRange.getTextStyle()
        .setBold(true)
        .setForegroundColor(settings.primaryColor);
      if (qFullText.length > 3) {
        const qContentRange = qTextRange.getRange(3, qFullText.length);
        qContentRange.getTextStyle().setBold(true);
      }
      qParsed.ranges.forEach(r => {
        const adjustedRange = qTextRange.getRange(r.start + 3, r.end + 3);
        if (r.bold) adjustedRange.getTextStyle().setBold(true);
        if (r.color) adjustedRange.getTextStyle().setForegroundColor(r.color);
      });
    } catch (e) {}
    try {
      qTextShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
    } catch (e) {}
    const aTop = qTop + qAreaHeight + qAGap;
    const aText = item.a || '';
    const aParsed = parseInlineStyles(aText);
    const aFullText = `A. ${aParsed.output}`;
    const aIndent = layout.pxToPt(16);
    const aTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      area.left + cardPadding + aIndent, aTop, 
      area.width - cardPadding * 2 - aIndent, aAreaHeight - layout.pxToPt(2));
    aTextShape.getFill().setTransparent();
    aTextShape.getBorder().setTransparent();
    const aTextRange = aTextShape.getText().setText(aFullText);
    applyTextStyle(aTextRange, {
      size: baseFontSize,
      color: CONFIG.COLORS.text_small_font,
      align: SlidesApp.ParagraphAlignment.LEFT
    });
    try {
      const aPrefixRange = aTextRange.getRange(0, 2);
      aPrefixRange.getTextStyle()
        .setBold(true)
        .setForegroundColor(generateTintedGray(settings.primaryColor, 15, 70));
      aParsed.ranges.forEach(r => {
        const adjustedRange = aTextRange.getRange(r.start + 3, r.end + 3);
        if (r.bold) adjustedRange.getTextStyle().setBold(true);
        if (r.color) adjustedRange.getTextStyle().setForegroundColor(r.color);
      });
    } catch (e) {}
    try {
      aTextShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
    } catch (e) {}
    currentY += cardHeight + cardGap;
  });
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function drawFaqItems(slide, items, layout, listArea, settings) {
  if (!items || !items.length) return;
  const px = v => layout.pxToPt(v);
  const GAP_ITEM = px(16);
  const PADDING = px(20);
  const totalCardHeight = listArea.height - (GAP_ITEM * (items.length - 1));
  const cardHeight = totalCardHeight / items.length;
  let currentY = listArea.top;
  items.forEach((qa) => {
    const card = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      listArea.left, currentY,
      listArea.width, cardHeight
    );
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const q = qa.q || '';
    const a = qa.a || '';
    const qIconWidth = px(30);
    const qTextLeft = listArea.left + PADDING + qIconWidth;
    const qTextWidth = listArea.width - PADDING * 2 - qIconWidth;
    const qIcon = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
      listArea.left + PADDING, currentY + PADDING, qIconWidth, px(24));
    setStyledText(qIcon, 'Q.', { size: 18, bold: true, color: settings.primaryColor });
    const qBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
      qTextLeft, currentY + PADDING, qTextWidth, px(40));
    setStyledText(qBox, q, { size: 14, bold: true, color: CONFIG.COLORS.text_primary });
    const aTop = currentY + PADDING + px(35);
    const aHeight = cardHeight - (PADDING * 2) - px(35);
    let answerFontSize = 14;
    if (a.length > 100) {
      answerFontSize = 11;
    } else if (a.length > 60) {
      answerFontSize = 12.5;
    }
    const aIcon = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
      listArea.left + PADDING, aTop, qIconWidth, aHeight);
    const tintedGrayColor = generateTintedGray(settings.primaryColor, 15, 70);
    setStyledText(aIcon, 'A.', { size: 18, bold: true, color: tintedGrayColor });
    const aBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
      qTextLeft, aTop, qTextWidth, aHeight);
    setStyledText(aBox, a, { size: answerFontSize, color: CONFIG.COLORS.text_primary });
    try {
      [qIcon, qBox, aIcon, aBox].forEach(s => {
        s.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        s.setAutofit(SlidesApp.AutofitType.SHRINK_ON_OVERFLOW);
      });
    } catch(e){}
    currentY += cardHeight + GAP_ITEM;
  });
}
function safeAlignTop(box){
  try { box.setContentAlignment(SlidesApp.ContentAlignment.TOP);
  } catch(e){}
}
function insertTrendIcon(slide, position, trend, settings) {
  const iconSize = 20;
  const icon = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
    position.left, position.top - iconSize/2, iconSize, iconSize);
  let iconText = '';
  let iconColor = CONFIG.COLORS.text_primary;
  switch (trend) {
    case 'up':
      iconText = '↑';
      iconColor = CONFIG.COLORS.success_green;
      break;
    case 'down':
      iconText = '↓';
      iconColor = CONFIG.COLORS.error_red;
      break;
    case 'neutral':
      iconText = '→';
      iconColor = CONFIG.COLORS.neutral_gray;
      break;
    default:
      iconText = '→';
      iconColor = CONFIG.COLORS.neutral_gray;
  }
  setStyledText(icon, iconText, {
    size: 16,
    bold: true,
    color: iconColor,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  try {
    icon.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
  return icon;
}
function parseNumericValue(str) {
  if (typeof str !== 'string') return 0;
  const match = str.match(/(\d+(\.\d+)?)/);
  return match ? parseFloat(match[1]) : 0;
}
function createStatsCompareSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'compareSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'compareSlide', data.subhead, subheadWidthPt); 
  const area = offsetRect(layout.getRect({
    left: 25,
    top: 130,
    width: 910,
    height: 330
  }), 0, dy);
  const stats = Array.isArray(data.stats) ? data.stats : [];
  if (stats.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const headerHeight = layout.pxToPt(40);
  const cushionTop = area.top + headerHeight;
  const cushionHeight = area.height - headerHeight;
  if (cushionHeight > 0) {
      const cushion = slide.insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          area.left,
          cushionTop,
          area.width,
          cushionHeight
      );
      cushion.getFill().setSolidFill(CONFIG.COLORS.background_gray);
      cushion.getBorder().setTransparent();
  }
  const totalContentWidth = area.width;
  const centerColWidth = totalContentWidth * 0.25;
  const sideColWidth = (totalContentWidth - centerColWidth) / 2;
  const leftValueColX = area.left;
  const centerLabelColX = leftValueColX + sideColWidth;
  const rightValueColX = centerLabelColX + centerColWidth;
  const labelColor = generateTintedGray(settings.primaryColor, 35, 70);
  const compareColors = generateCompareColors(settings.primaryColor);
  const leftHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftValueColX, area.top, sideColWidth, headerHeight);
  leftHeader.getFill().setSolidFill(compareColors.left);
  leftHeader.getBorder().setTransparent();
  setStyledText(leftHeader, data.leftTitle || '', {
    size: 14,
    bold: true,
    color: CONFIG.COLORS.background_gray,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  try { leftHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
  const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightValueColX, area.top, sideColWidth, headerHeight);
  rightHeader.getFill().setSolidFill(compareColors.right);
  rightHeader.getBorder().setTransparent();
  setStyledText(rightHeader, data.rightTitle || '', {
    size: 14,
    bold: true,
    color: CONFIG.COLORS.background_gray,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  try { rightHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
  const contentAreaHeight = area.height - headerHeight;
  const rowHeight = contentAreaHeight / stats.length;
  let currentY = area.top + headerHeight;
  stats.forEach((stat, index) => {
    const centerY = currentY + rowHeight / 2;
    const valueHeight = layout.pxToPt(40);
    const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, centerLabelColX, centerY - valueHeight / 2, centerColWidth, valueHeight);
    setStyledText(labelShape, stat.label || '', {
      size: 14,
      align: SlidesApp.ParagraphAlignment.CENTER,
      bold: true,
      //fontType: 'large'
    });
    try { labelShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
    const leftValueShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftValueColX, centerY - valueHeight / 2, sideColWidth, valueHeight);
    setStyledText(leftValueShape, stat.leftValue || '', {
      size: 22,
      bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER,
      color: compareColors.left
    });
    try { leftValueShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
    const rightValueShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightValueColX, centerY - valueHeight / 2, sideColWidth, valueHeight);
    setStyledText(rightValueShape, stat.rightValue || '', {
      size: 22,
      bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER,
      color: compareColors.right
    });
    try { rightValueShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
    if (index < stats.length - 1) {
      const lineY = currentY + rowHeight;
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, area.left + layout.pxToPt(15), lineY, area.left + area.width - layout.pxToPt(15), lineY);
      line.getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
      line.setWeight(1);
    }
    currentY += rowHeight;
  });
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createBarCompareSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'compareSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'compareSlide', data.subhead, subheadWidthPt); 
  const area = offsetRect(layout.getRect({
    left: 40,
    top: 130,
    width: 880,
    height: 340
  }), 0, dy);
  const stats = Array.isArray(data.stats) ? data.stats : [];
  if (stats.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const showTrends = !!data.showTrends;
  const blockMargin = layout.pxToPt(20);
  const totalContentHeight = area.height - (blockMargin * (stats.length - 1));
  const blockHeight = totalContentHeight / stats.length;
  let currentY = area.top;
  stats.forEach(stat => {
    const blockTop = currentY;
    const titleHeight = layout.pxToPt(35);
    const barAreaHeight = blockHeight - titleHeight;
    const barRowHeight = barAreaHeight / 2;
    const shouldShowTrend = showTrends && stat.trend;
    const statTitleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, blockTop, area.width, titleHeight);
    setStyledText(statTitleShape, stat.label || '', {
      size: 16,
      bold: true,
    });
    try { statTitleShape.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM); } catch(e){}
    const asIsY = blockTop + titleHeight;
    const toBeY = asIsY + barRowHeight;
    const labelWidth = layout.pxToPt(90);
    const valueWidth = layout.pxToPt(140);
    const barWidth = Math.max(layout.pxToPt(50), area.width - labelWidth - valueWidth - layout.pxToPt(10));
    const barLeft = area.left + labelWidth;
    const barHeight = layout.pxToPt(18);
    const val1 = parseNumericValue(stat.leftValue);
    const val2 = parseNumericValue(stat.rightValue);
    const maxValue = Math.max(val1, val2, 1);
    const asIsLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, asIsY, labelWidth, barRowHeight);
    setStyledText(asIsLabel, '現状', { size: 13, color: CONFIG.COLORS.neutral_gray });
    const asIsValue = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + barWidth, asIsY, valueWidth, barRowHeight);
    setStyledText(asIsValue, stat.leftValue || '', { size: 18, bold: true, align: SlidesApp.ParagraphAlignment.END });
    const asIsTrack = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, barLeft, asIsY + barRowHeight/2 - barHeight/2, barWidth, barHeight);
    asIsTrack.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
    asIsTrack.getBorder().setTransparent();
    const asIsFillWidth = Math.max(layout.pxToPt(2), barWidth * (val1 / maxValue));
    const asIsFill = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, barLeft, asIsY + barRowHeight/2 - barHeight/2, asIsFillWidth, barHeight);
    asIsFill.getFill().setSolidFill(CONFIG.COLORS.neutral_gray);
    asIsFill.getBorder().setTransparent();
    const toBeLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, toBeY, labelWidth, barRowHeight);
    setStyledText(toBeLabel, '導入後', { size: 13, color: settings.primaryColor, bold: true });
    const toBeValue = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + barWidth, toBeY, valueWidth, barRowHeight);
    setStyledText(toBeValue, stat.rightValue || '', { size: 18, bold: true, color: settings.primaryColor, align: SlidesApp.ParagraphAlignment.END });
    const toBeTrack = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, barLeft, toBeY + barRowHeight/2 - barHeight/2, barWidth, barHeight);
    toBeTrack.getFill().setSolidFill(generateTintedGray(settings.primaryColor, 20, 96));
    toBeTrack.getBorder().setTransparent();
    const toBeFillWidth = Math.max(layout.pxToPt(2), barWidth * (val2 / maxValue));
    const shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, barLeft, toBeY + barRowHeight/2 - barHeight/2, toBeFillWidth, barHeight);
    shape.getFill().setSolidFill(settings.primaryColor);
    shape.getBorder().setTransparent();
    try {
        [asIsLabel, asIsValue, toBeLabel, toBeValue].forEach(shape => shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE));
    } catch(e){}
    if (shouldShowTrend) {
      const trendIcon = insertTrendIcon(slide, { left: barLeft + barWidth + layout.pxToPt(10), top: toBeY + barRowHeight/2 }, stat.trend, settings);
    }
    currentY += blockHeight + blockMargin;
  });
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function smartFormatTriangleText(text) {
  if (!text || text.length <= 30) {
    return { text: text, isSimple: true, headerLength: 0 };
  }
  const separators = [
    { pattern: '：', priority: 1 },
    { pattern: ':', priority: 2 },
    { pattern: '。', priority: 3 },
    { pattern: 'について', priority: 4, keepSeparator: true },
    { pattern: 'における', priority: 5, keepSeparator: true }
  ];
  for (let sep of separators) {
    const index = text.indexOf(sep.pattern);
    if (index > 5 && index < text.length * 0.6) {
      const headerEnd = sep.keepSeparator ? index + sep.pattern.length : index;
      const header = text.substring(0, headerEnd).trim();
      const body = text.substring(index + sep.pattern.length).trim();
      if (header.length >= 3 && body.length >= 3) {
        return {
          text: `${header}\n${body}`,
          isSimple: false,
          headerLength: header.length
        };
      }
    }
  }
  if (text.length > 50) {
    const midPoint = Math.floor(text.length * 0.4);
    const header = text.substring(0, midPoint).trim();
    const body = text.substring(midPoint).trim();
    return {
      text: `${header}\n${body}`,
      isSimple: false,
      headerLength: header.length
    };
  }
  return { text: text, isSimple: true, headerLength: 0 };
}
function createTriangleSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'triangleSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'triangleSlide', data.subhead, subheadWidthPt); 
  const baseArea = layout.getRect('triangleSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const verticalShiftUpPt = layout.pxToPt(26.67);
  area.top -= verticalShiftUpPt;
  const items = Array.isArray(data.items) ? data.items.slice(0, 3) : [];
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const IMAGE_ADJUSTMENTS_LARGE = {
    topLeft:  { sizePt: layout.pxToPt(76), offsetXPt: layout.pxToPt(-105), offsetYPt: layout.pxToPt(-25) },
    topRight: { sizePt: layout.pxToPt(78), offsetXPt: layout.pxToPt(100),  offsetYPt: layout.pxToPt(-30) },
    bottom:   { sizePt: layout.pxToPt(95), offsetXPt: layout.pxToPt(0),    offsetYPt: layout.pxToPt(0)  }
  };
  const IMAGE_ADJUSTMENTS_MEDIUM = {
    topLeft:  { sizePt: layout.pxToPt(78), offsetXPt: layout.pxToPt(-95), offsetYPt: layout.pxToPt(-35) },
    topRight: { sizePt: layout.pxToPt(80), offsetXPt: layout.pxToPt(90),  offsetYPt: layout.pxToPt(-40) },
    bottom:   { sizePt: layout.pxToPt(100), offsetXPt: layout.pxToPt(0),   offsetYPt: layout.pxToPt(10)  }
  };
  const IMAGE_ADJUSTMENTS_SMALL = {
    topLeft:  { sizePt: layout.pxToPt(80), offsetXPt: layout.pxToPt(-85), offsetYPt: layout.pxToPt(-25) },
    topRight: { sizePt: layout.pxToPt(82), offsetXPt: layout.pxToPt(80),  offsetYPt: layout.pxToPt(-30) },
    bottom:   { sizePt: layout.pxToPt(105), offsetXPt: layout.pxToPt(0),   offsetYPt: layout.pxToPt(0)   }
  };
  const textLengths = items.map(item => (typeof item === 'string' ? item : (item.title || '') + (item.desc || '')).length);
  const maxLength = Math.max(0, ...textLengths);
  const avgLength = textLengths.length > 0 ? textLengths.reduce((sum, len) => sum + len, 0) / textLengths.length : 0;
  let cardW, cardH, fontSize;
  let selectedAdjustments;
  let isLargeCardSize = false;
  if (maxLength > 60 || avgLength > 40) {
    cardW = layout.pxToPt(340); cardH = layout.pxToPt(160); fontSize = 13;
    selectedAdjustments = IMAGE_ADJUSTMENTS_LARGE;
    isLargeCardSize = true;
  }
  else if (maxLength > 35 || avgLength > 25) {
    cardW = layout.pxToPt(290); cardH = layout.pxToPt(135); fontSize = 14;
    selectedAdjustments = IMAGE_ADJUSTMENTS_MEDIUM;
  }
  else {
    cardW = layout.pxToPt(250); cardH = layout.pxToPt(115); fontSize = 15;
    selectedAdjustments = IMAGE_ADJUSTMENTS_SMALL;
  }
  const maxCardW = (area.width - layout.pxToPt(160)) / 1.5;
  const maxCardH = (area.height - layout.pxToPt(80)) / 2;
  cardW = Math.min(cardW, maxCardW); cardH = Math.min(cardH, maxCardH);
  let bottomCardMarginPt;
  if (isLargeCardSize) {
    bottomCardMarginPt = layout.pxToPt(50);
  } else {
    bottomCardMarginPt = layout.pxToPt(80);
  }
  const positions = [
    { x: area.left + area.width / 2,                          y: area.top + layout.pxToPt(40) + cardH / 2 },
    { x: area.left + area.width - bottomCardMarginPt - cardW / 2, y: area.top + area.height - cardH / 2 },
    { x: area.left + bottomCardMarginPt + cardW / 2,             y: area.top + area.height - cardH / 2 }
  ];
  positions.forEach((pos, i) => {
    if (!items[i]) return;
    const cardX = pos.x - cardW / 2;
    const cardY = pos.y - cardH / 2;
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const item = items[i] || {};
    const itemTitle = typeof item === 'string' ? '' : (item.title || '');
    const itemDesc = typeof item === 'string' ? item : (item.desc || '');
    if (typeof item === 'string' || !itemTitle) {
        const itemText = typeof item === 'string' ? item : itemDesc;
        const processedText = smartFormatTriangleText(itemText);
        if (processedText.isSimple) {
            let appliedFontSize = fontSize;
            if ((processedText.text || '').length > 35) { appliedFontSize = Math.max(fontSize - 1, 12); }
            setStyledText(card, processedText.text, { size: appliedFontSize, bold: true, color: CONFIG.COLORS.text_small_font, align: SlidesApp.ParagraphAlignment.CENTER });
        } else {
            const lines = processedText.text.split('\n');
            const header = lines[0] || '';
            const body = lines.slice(1).join('\n') || '';
            const enhancedText = `${header}\n${body}`;
            const headerFontSize = Math.max(fontSize - 1, 13);
            const bodyFontSize = Math.max(fontSize - 3, 11);
            setStyledText(card, enhancedText, { size: bodyFontSize, bold: false, color: CONFIG.COLORS.text_small_font, align: SlidesApp.ParagraphAlignment.CENTER });
            try {
                const textRange = card.getText();
                const headerEndIndex = header.length;
                if (headerEndIndex > 0) {
                    const headerRange = textRange.getRange(0, headerEndIndex);
                    headerRange.getTextStyle().setBold(true).setFontSize(headerFontSize).setForegroundColor(settings.primaryColor);
                }
            } catch (e) { Logger.log(`Error styling triangle header (auto-split): ${e.message}`); }
        }
    } else {
        const enhancedText = itemDesc ? `${itemTitle}\n${itemDesc}` : itemTitle;
        const headerFontSize = Math.max(fontSize - 1, 13);
        const bodyFontSize = Math.max(fontSize - 3, 11);
        setStyledText(card, enhancedText, { size: bodyFontSize, bold: false, color: CONFIG.COLORS.text_small_font, align: SlidesApp.ParagraphAlignment.CENTER });
        try {
            const textRange = card.getText();
            const headerEndIndex = itemTitle.length;
            if (headerEndIndex > 0) {
                const headerRange = textRange.getRange(0, headerEndIndex);
                headerRange.getTextStyle().setBold(true).setFontSize(headerFontSize).setForegroundColor(settings.primaryColor);
            }
        } catch (e) { Logger.log(`Error styling triangle header (structured): ${e.message}`); }
    }
    try { setBoldTextSize(card, 16); } catch(e) { Logger.log(`Error applying setBoldTextSize in triangle slide: ${e.message}`); }
    try { card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
  });
  const triangleArrows = settings.triangleArrows;
  const hasValidImageData = triangleArrows &&
                             triangleArrows.topLeftArrowBase64 &&
                             triangleArrows.topRightArrowBase64 &&
                             triangleArrows.bottomArrowBase64;
  if (hasValidImageData && items.length === 3) {
    function insertTriangleArrowImage(arrowName, base64Data, left, top, size) {
    if (imageUpdateOption !== 'update') return;  
      if (!base64Data) {
        return;
      }
      try {
        const imageData = insertImageFromUrlOrFileId(base64Data);
        if (imageData) {
          const img = slide.insertImage(imageData);
          const originalWidth = img.getWidth();
          const originalHeight = img.getHeight();
          let targetWidth = size;
          let targetHeight = size;
          if (originalWidth > originalHeight) {
              targetHeight = size * (originalHeight / originalWidth);
          } else if (originalHeight > originalWidth) {
              targetWidth = size * (originalWidth / originalHeight);
          }
          img.setLeft(left + (size - targetWidth) / 2)
             .setTop(top + (size - targetHeight) / 2)
             .setWidth(targetWidth)
             .setHeight(targetHeight);
        }
      } catch (e) {
      }
    }
    const arrow1_baseX = (positions[2].x + positions[0].x) / 2;
    const arrow1_baseY = (positions[2].y + positions[0].y) / 2;
    const adj1 = selectedAdjustments.topLeft;
    insertTriangleArrowImage(
      'topLeft',
      triangleArrows.topLeftArrowBase64,
      arrow1_baseX - (adj1.sizePt / 2) + adj1.offsetXPt,
      arrow1_baseY - (adj1.sizePt / 2) + adj1.offsetYPt,
      adj1.sizePt
    );
    const arrow2_baseX = (positions[0].x + positions[1].x) / 2;
    const arrow2_baseY = (positions[0].y + positions[1].y) / 2;
    const adj2 = selectedAdjustments.topRight;
    insertTriangleArrowImage(
      'topRight',
      triangleArrows.topRightArrowBase64,
      arrow2_baseX - (adj2.sizePt / 2) + adj2.offsetXPt,
      arrow2_baseY - (adj2.sizePt / 2) + adj2.offsetYPt,
      adj2.sizePt
    );
    const arrow3_baseX = (positions[1].x + positions[2].x) / 2;
    const arrow3_baseY = (positions[1].y + positions[2].y) / 2;
    const adj3 = selectedAdjustments.bottom;
    insertTriangleArrowImage(
      'bottom',
      triangleArrows.bottomArrowBase64,
      arrow3_baseX - (adj3.sizePt / 2) + adj3.offsetXPt,
      arrow3_baseY - (adj3.sizePt / 2) + adj3.offsetYPt,
      adj3.sizePt
    );
  } else if (items.length === 3 && imageUpdateOption === 'update') {
    if (items.length === 3) {
      const arrowPadding = cardW > layout.pxToPt(300) ? layout.pxToPt(25) : layout.pxToPt(20);
      const cardEdges = [];
      cardEdges[0] = { rightCenter: { x: positions[0].x + cardW / 2, y: positions[0].y }, leftCenter: { x: positions[0].x - cardW / 2, y: positions[0].y }, bottomCenter: { x: positions[0].x, y: positions[0].y + cardH / 2 } };
      cardEdges[1] = { leftCenter: { x: positions[1].x - cardW / 2, y: positions[1].y }, topCenter: { x: positions[1].x, y: positions[1].y - cardH / 2 } };
      cardEdges[2] = { rightCenter: { x: positions[2].x + cardW / 2, y: positions[2].y }, topCenter: { x: positions[2].x, y: positions[2].y - cardH / 2 } };
      const arrowLines = [];
      arrowLines[0] = { startX: cardEdges[0].rightCenter.x + arrowPadding, startY: cardEdges[0].rightCenter.y, endX: cardEdges[1].topCenter.x + arrowPadding / 2 , endY: cardEdges[1].topCenter.y - arrowPadding};
      arrowLines[1] = { startX: cardEdges[1].leftCenter.x - arrowPadding, startY: cardEdges[1].leftCenter.y + arrowPadding / 2, endX: cardEdges[2].rightCenter.x + arrowPadding, endY: cardEdges[2].rightCenter.y + arrowPadding / 2};
      arrowLines[2] = { startX: cardEdges[2].topCenter.x - arrowPadding / 2, startY: cardEdges[2].topCenter.y - arrowPadding, endX: cardEdges[0].leftCenter.x - arrowPadding, endY: cardEdges[0].leftCenter.y};
      arrowLines.forEach(lineDef => {
        try {
          const line = slide.insertLine(
            SlidesApp.LineCategory.STRAIGHT,
            lineDef.startX, lineDef.startY,
            lineDef.endX, lineDef.endY
          );
          line.getLineFill().setSolidFill(CONFIG.COLORS.ghost_gray);
          line.setWeight(4);
          line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
        } catch (e) {
        }
      });
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createPyramidSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'pyramidSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'pyramidSlide', data.subhead, subheadWidthPt); 
  const baseArea = layout.getRect('pyramidSlide.pyramidArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  const levels = Array.isArray(data.levels) ? data.levels.slice(0, 4) : [];
  if (levels.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const levelHeight = layout.pxToPt(70);
  const levelGap = layout.pxToPt(2);
  const totalHeight = (levelHeight * levels.length) + (levelGap * (levels.length - 1));
  const startY = area.top + (area.height - totalHeight) / 2;
  const pyramidWidth = layout.pxToPt(480);
  const textColumnWidth = layout.pxToPt(400);
  const gap = layout.pxToPt(30);
  const pyramidLeft = area.left;
  const textColumnLeft = pyramidLeft + pyramidWidth + gap;
  const pyramidColors = generatePyramidColors(settings.primaryColor, levels.length);
  const baseWidth = pyramidWidth;
  const widthIncrement = baseWidth / levels.length;
  const centerX = pyramidLeft + pyramidWidth / 2;
  levels.forEach((level, index) => {
    const levelWidth = baseWidth - (widthIncrement * (levels.length - 1 - index));
    const levelX = centerX - levelWidth / 2;
    const levelY = startY + index * (levelHeight + levelGap);
    const levelBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, levelX, levelY, levelWidth, levelHeight);
    levelBox.getFill().setSolidFill(pyramidColors[index]);
    levelBox.getBorder().setTransparent();
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
    try {
      titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    const connectionStartX = levelX + levelWidth;
    const connectionEndX = textColumnLeft;
    const connectionY = levelY + levelHeight / 2;
    if (connectionEndX > connectionStartX) {
      const connectionLine = slide.insertLine(
        SlidesApp.LineCategory.STRAIGHT,
        connectionStartX,
        connectionY,
        connectionEndX,
        connectionY
      );
      connectionLine.getLineFill().setSolidFill('#D0D7DE');
      connectionLine.setWeight(1.5);
    }
    const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textColumnLeft, levelY, textColumnWidth, levelHeight);
    textShape.getFill().setTransparent();
    textShape.getBorder().setTransparent();
    const levelDesc = level.description || '';
    let formattedText;
    if (levelDesc.includes('•') || levelDesc.includes('・')) {
      formattedText = levelDesc;
    } else if (levelDesc.includes('\n')) {
      const lines = levelDesc.split('\n').filter(line => line.trim()).slice(0, 2);
      formattedText = lines.map(line => `• ${line.trim()}`).join('\n');
    } else {
      formattedText = levelDesc;
    }
    setStyledText(textShape, formattedText, {
      size: CONFIG.FONTS.sizes.body - 1,
      align: SlidesApp.ParagraphAlignment.LEFT,
      color: CONFIG.COLORS.text_primary,
      bold: true
    });
    try {
      textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
  });
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createFlowChartSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'flowChartSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'flowChartSlide', data.subhead, subheadWidthPt); 
  const flows = Array.isArray(data.flows) ? data.flows : [{ steps: data.steps || [] }];
  let isDouble = flows.length > 1;
  let upperFlow, lowerFlow, maxStepsPerRow;
  if (!isDouble && flows[0] && flows[0].steps && flows[0].steps.length >= 5) {
    isDouble = true;
    const allSteps = flows[0].steps;
    const midPoint = Math.ceil(allSteps.length / 2);
    upperFlow = { steps: allSteps.slice(0, midPoint) };
    lowerFlow = { steps: allSteps.slice(midPoint) };
    maxStepsPerRow = midPoint;
  } else {
    upperFlow = flows[0];
    lowerFlow = flows.length > 1 ? flows[1] : null;
    maxStepsPerRow = Math.max(
      upperFlow?.steps?.length || 0, 
      lowerFlow?.steps?.length || 0
    );
  }
  if (isDouble) {
    const upperArea = offsetRect(layout.getRect('flowChartSlide.upperRow'), 0, dy);
    const lowerArea = offsetRect(layout.getRect('flowChartSlide.lowerRow'), 0, dy);
    drawFlowRow(slide, upperFlow, upperArea, settings, layout, maxStepsPerRow);
    if (lowerFlow && lowerFlow.steps && lowerFlow.steps.length > 0) {
      drawFlowRow(slide, lowerFlow, lowerArea, settings, layout, maxStepsPerRow);
    }
  } else {
    const baseSingleArea = layout.getRect('flowChartSlide.singleRow');
    let singleArea = offsetRect(baseSingleArea, 0, dy);
    const verticalShiftDownPt = layout.pxToPt(26.67);
    singleArea.top += verticalShiftDownPt;
    drawFlowRow(slide, flows[0], singleArea, settings, layout);
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function drawFlowRow(slide, flow, area, settings, layout, maxStepsPerRow = null) {
  if (!flow || !flow.steps || !Array.isArray(flow.steps)) {
    return;
  }
  const steps = flow.steps.filter(step => step && String(step).trim());
  if (steps.length === 0) return;
  const actualSteps = maxStepsPerRow || steps.length;
  const baseArrowSpace = layout.pxToPt(25);
  const arrowSpace = Math.max(baseArrowSpace, area.width * 0.04);
  const totalArrowSpace = (actualSteps - 1) * arrowSpace;
  const cardW = (area.width - totalArrowSpace) / actualSteps;
  const cardH = area.height;
  const arrowHeight = Math.min(cardH * 0.3, layout.pxToPt(40));
  const arrowWidth = arrowSpace;
  steps.forEach((step, index) => {
    const cardX = area.left + index * (cardW + arrowSpace);
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, area.top, cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const stepText = String(step || '').trim() || 'ステップ';
    setStyledText(card, stepText, {
      size: CONFIG.FONTS.sizes.body,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    if (index < steps.length - 1) {
      const arrowStartX = cardX + cardW;
      const arrowCenterY = area.top + cardH / 2;
      const arrowTop = arrowCenterY - (arrowHeight / 2);
      const arrow = slide.insertShape(SlidesApp.ShapeType.RIGHT_ARROW, arrowStartX, arrowTop, arrowWidth, arrowHeight);
      arrow.getFill().setSolidFill(settings.primaryColor);
      arrow.getBorder().setTransparent();
    }
  });
}
function createStepUpSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'stepUpSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'stepUpSlide', data.subhead, subheadWidthPt); 
  const area = offsetRect(layout.getRect('stepUpSlide.stepArea'), 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const numSteps = Math.min(5, items.length);
  const gap = 0;
  const headerHeight = layout.pxToPt(40);
  const maxHeight = area.height * 0.95;
  let minHeightRatio;
  if (numSteps <= 2) {
    minHeightRatio = 0.70;
  } else if (numSteps === 3) {
    minHeightRatio = 0.60;
  } else {
    minHeightRatio = 0.50;
  }
  const minHeight = maxHeight * minHeightRatio;
  const totalWidth = area.width;
  const cardW = totalWidth / numSteps;
  const stepUpColors = generateStepUpColors(settings.primaryColor, numSteps);
  for (let idx = 0; idx < numSteps; idx++) {
    const item = items[idx] || {};
    const titleText = String(item.title || `STEP ${idx + 1}`);
    const descText = String(item.desc || '');
    const heightRatio = (idx / Math.max(1, numSteps - 1));
    const cardH = minHeight + (maxHeight - minHeight) * heightRatio;
    const left = area.left + idx * cardW;
    const top = area.top + area.height - cardH;
    const bodyShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top + headerHeight, cardW, cardH - headerHeight);
    bodyShape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    bodyShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const headerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, cardW, headerHeight);
    headerShape.getFill().setSolidFill(stepUpColors[idx]);
    headerShape.getBorder().getLineFill().setSolidFill(stepUpColors[idx]);
    const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, cardW, headerHeight);
    setStyledText(headerTextShape, titleText, {
      size: CONFIG.FONTS.sizes.body,
      bold: true,
      color: CONFIG.COLORS.background_gray,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      setParenthesizedTextSize(headerTextShape, 10);
    } catch(e) {
    }
    try {
      headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      headerTextShape.setAutofit(SlidesApp.AutofitType.SHRINK_ON_OVERFLOW);
    } catch (e) {}
    const bodyTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      left + layout.pxToPt(8), top + headerHeight, 
      cardW - layout.pxToPt(16), cardH - headerHeight);
    setStyledText(bodyTextShape, descText, {
      size: CONFIG.FONTS.sizes.body,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      bodyTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      bodyTextShape.setAutofit(SlidesApp.AutofitType.SHRINK_ON_OVERFLOW);
    } catch (e) {}
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function createImageTextSlide(slide, data, layout, pageNum, settings, imageUpdateOption = 'update') {
  setMainSlideBackground(slide, layout, imageUpdateOption);
  const titleWidthPt = (data && typeof data._title_widthPt === 'number') ? data._title_widthPt : null;
  drawStandardTitleHeader(slide, layout, 'imageTextSlide', data.title, settings, titleWidthPt, imageUpdateOption);
  const subheadWidthPt = (data && typeof data._subhead_widthPt === 'number') ?
  data._subhead_widthPt : null;
  const dy = drawSubheadIfAny(slide, layout, 'imageTextSlide', data.subhead, subheadWidthPt);
  let imageUrl = '';
  let imageInfo = null;
  const imageData = data.image;
  if (typeof imageData === 'string') {
    imageUrl = imageData;
  } else if (imageData && typeof imageData === 'object') {
    imageUrl = imageData.data || '';
    imageInfo = imageData.info || null;
  }
  const imageCaption = data.imageCaption || '';
  const points = Array.isArray(data.points) ? data.points : [];
  const imagePosition = data.imagePosition === 'right' ? 'right' : 'left';
  if (imagePosition === 'left') {
    const textArea = offsetRect(layout.getRect('imageTextSlide.rightText'), 0, dy);
    if (imageUrl) {
      let newMaxWidthPt;
      let newMaxHeightPt;
      if (imageInfo === 'chart') {
        newMaxWidthPt = layout.pxToPt(400);
        newMaxHeightPt = layout.pxToPt(400);
      } else {
        newMaxWidthPt = layout.pxToPt(400);
        newMaxHeightPt = layout.pxToPt(300);
      }
      renderImageMirrored(
        slide, 
        layout, 
        textArea,
        imageUrl, 
        newMaxWidthPt, 
        newMaxHeightPt, 
        imageUpdateOption
      );
    }
    if (points.length > 0) { 
      createContentCushion(slide, textArea, settings, layout);
      const padding = layout.pxToPt(20);
      const textRect = { 
        left: textArea.left + padding,
        top: textArea.top + padding,
        width: textArea.width - (padding * 2),
        height: textArea.height - (padding * 2)
      };
      const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);  
      setBulletsWithInlineStyles(textShape, points);
      try {
        setBoldTextSize(textShape, 16);
      } catch(e) {
      }
      try {
        adjustShapeText_External(textShape);
      } catch(e) {
      }
      try {
        setParenthesizedTextSize(textShape, 10);
      } catch(e) {
      }
      try {
        textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch(e) {
      }
    }
  } else {
    const textArea = offsetRect(layout.getRect('imageTextSlide.leftText'), 0, dy);
    const imageArea = offsetRect(layout.getRect('imageTextSlide.rightImage'), 0, dy);
    if (points.length > 0) { 
      createContentCushion(slide, textArea, settings, layout);
      const padding = layout.pxToPt(20);
      const textRect = { 
        left: textArea.left + padding,
        top: textArea.top + padding,
        width: textArea.width - (padding * 2),
        height: textArea.height - (padding * 2)
      };
      const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height); 
      setBulletsWithInlineStyles(textShape, points); 
      try {
      setBoldTextSize(textShape, 16);
      } catch(e) {
      }
      try {
        adjustShapeText_External(textShape);
      } catch(e) {
      }
      try {
      setParenthesizedTextSize(textShape, 10);
      } catch(e) {
      }
      try {
      textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch(e) {
      }
    }
    if (imageUrl) { 
      let newMaxWidthPt;
      let newMaxHeightPt;
      if (imageInfo === 'chart') {
        newMaxWidthPt = layout.pxToPt(400);
        newMaxHeightPt = layout.pxToPt(400);
      } else {
        newMaxWidthPt = layout.pxToPt(400);
        newMaxHeightPt = layout.pxToPt(300);
      }
      renderImageMirrored(
        slide, 
        layout, 
        textArea,
        imageUrl, 
        newMaxWidthPt, 
        newMaxHeightPt, 
        imageUpdateOption
      );
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
function renderSingleImageInArea(slide, layout, area, imageUrl, caption = '', position = 'left', imageInfo = null, imageUpdateOption = 'update') {
  if (!imageUrl) return;
  if (imageUpdateOption !== 'update') {
    return null;
  }
  try {
    const imageData = insertImageFromUrlOrFileId(imageUrl);
    if (!imageData) return null;
    const img = slide.insertImage(imageData);
    if (imageInfo === 'chart') {
      const imgAspect = img.getWidth() / img.getHeight();
      const baseTop = area.top;
      const baseCenterX = area.left + area.width / 2;
      const newBottom = layout.pageH_pt - layout.pxToPt(5);
      const newH = newBottom - baseTop;
      const newW = newH * imgAspect;
      const newLeft = baseCenterX - (newW / 2);
      img.setWidth(newW).setHeight(newH)
         .setLeft(newLeft)
         .setTop(baseTop);
    } else {
      const scale = Math.min(area.width / img.getWidth(), area.height / img.getHeight());
      const w = img.getWidth() * scale;
      const h = img.getHeight() * scale;
      img.setWidth(w).setHeight(h)
         .setLeft(area.left + (area.width - w) / 2)
         .setTop(area.top + (area.height - h) / 2);
    }
    if (caption && caption.trim()) { 
      const imgHeight = img.getHeight();
      const imgTop = img.getTop();
      const imageBottom = imgTop + imgHeight;
      const captionMargin = layout.pxToPt(8);
      const captionHeight = layout.pxToPt(30);
      let captionRect;
      if (imageInfo === 'chart') {
        let captionTop = imageBottom + captionMargin;
        if ((captionTop + captionHeight) > layout.pageH_pt) {
          captionTop = layout.pageH_pt - captionHeight - layout.pxToPt(2);
        }
        captionRect = { left: img.getLeft(), top: captionTop, width: img.getWidth(), height: captionHeight };
      } else {
        const rectPath = `imageTextSlide.${position}ImageCaption`;
        captionRect = layout.getRect(rectPath);
      }
      const captionShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
        captionRect.left, captionRect.top, captionRect.width, captionRect.height);
      captionShape.getFill().setTransparent();
      captionShape.getBorder().setTransparent();
      setStyledText(captionShape, caption.trim(), {
        size: CONFIG.FONTS.sizes.small,
        color: CONFIG.COLORS.neutral_gray,
        align: SlidesApp.ParagraphAlignment.CENTER
      });
      try {
        captionShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {}
    }
    return img;
  } catch (e) {
    return null;
  }
}
function renderImageMirrored(slide, layout, textAreaRect, imageUrl, maxWidthPt, maxHeightPt, imageUpdateOption) {
  if (imageUpdateOption !== 'update') {
    return null;
  }
  if (!imageUrl) {
    return null;
  }
  if (!textAreaRect || !maxWidthPt || !maxHeightPt) {
    return null;
  }
  try {
    const pageCenterX = layout.pageW_pt / 2;
    const textCenterX = textAreaRect.left + textAreaRect.width / 2;
    const textCenterY = textAreaRect.top + textAreaRect.height / 2;
    const mirroredCenterX = pageCenterX + (pageCenterX - textCenterX);
    const mirroredCenterY = textCenterY;
    const imageData = insertImageFromUrlOrFileId(imageUrl);
    if (!imageData) return null;
    const img = slide.insertImage(imageData);
    const scale = Math.min(maxWidthPt / img.getWidth(), maxHeightPt / img.getHeight());
    const w = img.getWidth() * scale;
    const h = img.getHeight() * scale;
    const finalLeft = mirroredCenterX - (w / 2);
    const finalTop = mirroredCenterY - (h / 2);
    img.setWidth(w).setHeight(h)
       .setLeft(finalLeft)
       .setTop(finalTop);
    return img;
  } catch (e) {
    return null;
  }
}
function estimateTextWidthPt(text, fontSizePt) {
  const multipliers = {
    space: 0.3,
    ascii: 0.62,
    uppercase: 0.8,
    japanese: 1.0,
    other: 0.85
  };
  const totalMultiplier = String(text || '').split('').reduce((acc, char) => {
    if (char === ' ') {
      return acc + multipliers.space;
    }
    else if (char === 'I' || char === 'i') {
      return acc + multipliers.space;
    }
    else if (char.match(/[A-HJ-Z]/)) {
      return acc + multipliers.uppercase;
    }
    else if (char.match(/[ -~]/)) {
      return acc + multipliers.ascii;
    }
    else if (char.match(/[\u3000-\u303F\u3040-\u309F\u30A0-\uFF00-\uFFEF]/)) {
      return acc + multipliers.japanese;
    }
    else if (char.match(/[\u4E00-\u9FAF]/)) {
        return acc + multipliers.japanese;
    }
    else {
      return acc + multipliers.other;
    }
  }, 0);
  const calculatedWidth = totalMultiplier * fontSizePt;
  return calculatedWidth + 10;
}
function offsetRect(rect, dx, dy) {
  return {
    left: rect.left + (dx || 0),
    top: rect.top + (dy || 0),
    width: rect.width,
    height: rect.height
  };
}
function drawStandardTitleHeader(slide, layout, key, title, settings, preCalculatedWidthPt = null, imageUpdateOption = 'update') {
  if (imageUpdateOption === 'update') {
    const logoRect = safeGetRect(layout, `${key}.headerLogo`);
    try {
      if (CONFIG.LOGOS.header && logoRect) {
        const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.header);
        if (imageData) {
          const logo = slide.insertImage(imageData);
          const asp = logo.getHeight() / logo.getWidth();
          logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * asp);
        }
      }
    } catch (e) {
    }
  }
  const titleRect = safeGetRect(layout, `${key}.title`);
  if (!titleRect) {
    return;
  }
  const initialFontSize = CONFIG.FONTS.sizes.contentTitle;
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
  });
  try {
    titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
  let finalFontSize = initialFontSize;
  let adjustmentLogResult = '';
  try {
    const adjustmentResult = adjustShapeText_External(titleShape, preCalculatedWidthPt);
    adjustmentLogResult = adjustmentResult.log;
    if (adjustmentResult.finalSize !== null) {
      finalFontSize = adjustmentResult.finalSize;
    }
  } catch (e) {
  }
  if (settings.showTitleUnderline && title) {
    const uRect = safeGetRect(layout, `${key}.titleUnderline`);
    if (!uRect) {
      return;
    }
    let underlineWidthPt;
    if (finalFontSize < initialFontSize && finalFontSize >= 10 && preCalculatedWidthPt !== null && preCalculatedWidthPt > 0 && initialFontSize > 0) {
        underlineWidthPt = preCalculatedWidthPt * (finalFontSize / initialFontSize);
    }
    else if (preCalculatedWidthPt !== null && preCalculatedWidthPt > 0) {
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
function estimateTextHeightPt(text, widthPt, fontSizePt, lineHeight) {
  var paragraphs = String(text).split(/\r?\n/);
  var charsPerLine = Math.max(1, Math.floor(widthPt / (fontSizePt * 0.95)));
  var lines = 0;
  for (var i = 0; i < paragraphs.length; i++) {
    var s = paragraphs[i].replace(/\s+/g, ' ').trim();
    var len = s.length || 1;
    lines += Math.ceil(len / charsPerLine);
  }
  var lineH = fontSizePt * (lineHeight || 1.2);
  return Math.max(lineH, lines * lineH);
}
function drawSubheadIfAny(slide, layout, key, subhead, preCalculatedWidthPt = null) {
  if (!subhead) return 0;
  const rect = safeGetRect(layout, `${key}.subhead`);
  if (!rect) {
    return 0;
  }
  const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left, rect.top, rect.width, rect.height);
  setStyledText(box, subhead, {
    size: CONFIG.FONTS.sizes.subhead,
    //color: CONFIG.COLORS.text_primary
    fontType: 'large'
  });
  try {
    const adjustmentResult = adjustShapeText_External(box, preCalculatedWidthPt);
  } catch (e) {
  }
  return layout.pxToPt(36);
}
function drawBottomBar(slide, layout, settings) {
  const barRect = layout.getRect('bottomBar');
  applyFill(slide, barRect.left, barRect.top, barRect.width, barRect.height, settings);
}
function addCucFooter(slide, layout, pageNum, settings) {
  if (CONFIG.FOOTER_TEXT && CONFIG.FOOTER_TEXT.trim() !== '') {
    const leftRect = layout.getRect('footer.leftText');
    const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftRect.left, leftRect.top, leftRect.width, leftRect.height);
  leftShape.getText().setText(CONFIG.FOOTER_TEXT);
  applyTextStyle(leftShape.getText(), {
      size: CONFIG.FONTS.sizes.footer,
      //color: CONFIG.COLORS.text_primary
      fontType: 'large'
    });
    try {
      leftShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch(e) {}
  }
  if (__CREDIT_IMAGE_BLOB) {
    drawCreditImage(slide, layout, __CREDIT_IMAGE_BLOB, CREDIT_IMAGE_LINK);
  }
  if (pageNum > 0 && settings && settings.showPageNumber) { 
    const rightRect = layout.getRect('footer.rightPage');
    const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightRect.left, rightRect.top, rightRect.width, rightRect.height);
    rightShape.getText().setText(String(pageNum)); 
    applyTextStyle(rightShape.getText(), {
      size: CONFIG.FONTS.sizes.footer,
      color: CONFIG.COLORS.primary_color,
      align: SlidesApp.ParagraphAlignment.END
    });
    try {
      rightShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch(e) {}
  }
}
function drawBottomBarAndFooter(slide, layout, pageNum, settings) {
  if (settings.showBottomBar) {
    drawBottomBar(slide, layout, settings);
  }
  addCucFooter(slide, layout, pageNum, settings);
}
function applyTextStyle(textRange, opt) {
  const style = textRange.getTextStyle();
  let defaultColor;
  if (opt.fontType === 'large') {
    defaultColor = CONFIG.COLORS.text_primary;
  } else {
    defaultColor = CONFIG.COLORS.text_small_font;
  }
  style.setFontFamily(CONFIG.FONTS.family)
       .setForegroundColor(opt.color || defaultColor)
       .setFontSize(opt.size || CONFIG.FONTS.sizes.body)
       .setBold(opt.bold || false);
  if (opt.align) {
    try {
      textRange.getParagraphs().forEach(p => {
        p.getRange().getParagraphStyle().setParagraphAlignment(opt.align);
      });
    } catch (e) {}
  }
}
function setStyledText(shapeOrCell, rawText, baseOpt) {
  const parsed = parseInlineStyles(rawText || '');
  const tr = shapeOrCell.getText().setText(parsed.output);
  applyTextStyle(tr, baseOpt || {});
  applyStyleRanges(tr, parsed.ranges);
}
function setBulletsWithInlineStyles(shape, points) {
  const joiner = '\n\n';
  let combined = '';
  const ranges = [];
  (points || []).forEach((pt, idx) => {
    const parsed = parseInlineStyles(String(pt || ''));
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
  const tr = shape.getText().setText(combined || '—');
  applyTextStyle(tr, {
    size: CONFIG.FONTS.sizes.body
  });
  try {
    tr.getParagraphs().forEach(p => {
      p.getRange().getParagraphStyle().setLineSpacing(100).setSpaceBelow(6);
    });
  } catch (e) {}
  applyStyleRanges(tr, ranges);
}
function checkSpacing(s, out, i, nextCharIndex) {
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
function parseInlineStyles(s) {
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
          color: CONFIG.COLORS.primary_color,
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
          color: CONFIG.COLORS.primary_color,
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
function cleanSpeakerNotes(notesText) {
  if (!notesText) return '';
  let cleaned = notesText;
  cleaned = cleaned.replace(/\*\*([^*]+)\*\*/g, '$1');
  cleaned = cleaned.replace(/\[\[([^\]]+)\]\]/g, '$1');
  cleaned = cleaned.replace(/\*([^*]+)\*/g, '$1');
  cleaned = cleaned.replace(/_([^_]+)_/g, '$1');
  cleaned = cleaned.replace(/~~([^~]+)~~/g, '$1');
  cleaned = cleaned.replace(/`([^`]+)`/g, '$1');
  return cleaned;
}
function applyStyleRanges(textRange, ranges) {
  ranges.forEach(r => {
    try {
      const sub = textRange.getRange(r.start, r.end);
      if (!sub) return;
      const st = sub.getTextStyle();
      if (r.bold) st.setBold(true);
      if (r.color) st.setForegroundColor(r.color);
      if (r.size) st.setFontSize(r.size);
    } catch (e) {}
  });
}
function isAgendaTitle(title) {
  return /(agenda|アジェンダ|目次|本日お伝えすること)/i.test(String(title || ''));
}
function buildAgendaFromSlideData() {
  return __SLIDE_DATA_FOR_AGENDA.filter(d => d && d.type === 'section' && d.title).map(d => d.title.trim());
}
function drawCompareBox(slide, layout, left, top, width, height, title, items, settings, isLeft = false) {
  const box = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, width, height);
  box.getFill().setSolidFill(CONFIG.COLORS.background_gray);
  box.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);
  const th = layout.pxToPt(40);
  const titleBarBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, width, th);
  const compareColors = generateCompareColors(settings.primaryColor);
  const headerColor = isLeft ? compareColors.left : compareColors.right;
  titleBarBg.getFill().setSolidFill(headerColor);
  titleBarBg.getBorder().getLineFill().setSolidFill(headerColor);
  const titleTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, width, th);
  titleTextShape.getFill().setTransparent();
  titleTextShape.getBorder().setTransparent();
  setStyledText(titleTextShape, title, {
    size: CONFIG.FONTS.sizes.laneTitle,
    bold: true,
    color: CONFIG.COLORS.background_gray,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  const pad = layout.pxToPt(12);
  const body = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
    left + pad,
    top + th + pad,
    width - pad * 2,
    height - th - pad * 2
  );
  setBulletsWithInlineStyles(body, items);
  try {
    const adjustmentLog = adjustShapeText_External(body, null); 
  } catch(e) {
  }
}
function drawArrowBetweenRects(slide, a, b, arrowH, arrowGap, settings) {
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
function drawNumberedItems(slide, layout, area, items, settings) {
  createContentCushion(slide, area, settings, layout);
  const n = Math.max(1, items.length);
  const topPadding = layout.pxToPt(30);
  const bottomPadding = layout.pxToPt(10);
  const drawableHeight = area.height - topPadding - bottomPadding;
  const gapY = drawableHeight / Math.max(1, n - 1);
  const cx = area.left + layout.pxToPt(44);
  const top0 = area.top + topPadding;
  const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx - layout.pxToPt(1), top0 + layout.pxToPt(6), layout.pxToPt(2), gapY * (n - 1));
  line.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
  line.getBorder().setTransparent();
  for (let i = 0; i < n; i++) {
    const cy = top0 + gapY * i;
    const sz = layout.pxToPt(28);
    const numBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx - sz/2, cy - sz/2, sz, sz);
    numBox.getFill().setSolidFill(settings.primaryColor);
    numBox.getBorder().setTransparent();
    const num = numBox.getText(); num.setText(String(i + 1));
    applyTextStyle(num, { size: 12, bold: true, color: CONFIG.COLORS.background_white, align: SlidesApp.ParagraphAlignment.CENTER });
    let cleanText = String(items[i] || '');
    cleanText = cleanText.replace(/^\s*\d+[\.\s]*/, '');
    const txt = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32));
    setStyledText(txt, cleanText, { size: CONFIG.FONTS.sizes.processStep });
    try { txt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
  }
}
function setBoldTextSize(shapeOrTextRange, size = 16) {
  let textRange;
  try {
    if (shapeOrTextRange && typeof shapeOrTextRange.getText === 'function') {
      textRange = shapeOrTextRange.getText();
    } else if (shapeOrTextRange && typeof shapeOrTextRange.getRuns === 'function') {
      textRange = shapeOrTextRange;
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
function setParenthesizedTextSize(shapeOrTextRange, size) {
  let textRange;
  try {
    if (shapeOrTextRange && typeof shapeOrTextRange.getText === 'function') {
      textRange = shapeOrTextRange.getText();
    } else if (shapeOrTextRange && typeof shapeOrTextRange.getRuns === 'function') {
      textRange = shapeOrTextRange;
    } else {
      return;
    }
    if (!textRange || textRange.isEmpty()) {
      return;
    }
    const fullText = textRange.asString();
    const matches = [];
    const regex = /(\(|（)([^()（）]+)(\)|）)/g;
    let match;
    while ((match = regex.exec(fullText)) !== null) {
        matches.push(match);
    }
    for (let i = matches.length - 1; i >= 0; i--) {
        match = matches[i];
        const startIndex = match.index;
        const endIndex = match.index + match[0].length;
        const originalText = match[0];
        let newText = originalText;
        if (match[1] === '（') {
            newText = ' (' + newText.substring(1);
        }
        if (match[3] === '）') {
            newText = newText.substring(0, newText.length - 1) + ') ';
        }
        if (startIndex < endIndex) {
            try {
                const subRange = textRange.getRange(startIndex, endIndex);
                subRange.setText(newText);
                const newEndIndex = startIndex + newText.length;
                const replacedRange = textRange.getRange(startIndex, newEndIndex);
                replacedRange.getTextStyle().setFontSize(size);
            } catch (rangeError) {
            }
        }
    }
  } catch (e) {
  }
}
function drawCreditImage(slide, layout, creditImageBlob, creditLink) {
  try {
    const creditPosPx = CONFIG.POS_PX.footer.creditImage;
    if (!creditPosPx || !creditImageBlob) {
      return;
    }
    const img = slide.insertImage(creditImageBlob);
    const newHeight = layout.pxToPt(creditPosPx.height) * layout.scaleY;
    const aspect = img.getWidth() / img.getHeight();
    const newWidth = newHeight * aspect;
    const rightMarginPt = layout.pxToPt(creditPosPx.right) * layout.scaleX;
    const newLeft = (layout.pageW_pt - rightMarginPt) - newWidth;
    const topPt = layout.pxToPt(creditPosPx.top) * layout.scaleY;
    img.setLeft(newLeft)
       .setTop(topPt)
       .setWidth(newWidth)
       .setHeight(newHeight)
       .setLinkUrl(creditLink);
  } catch (e) {
  }
}
function safeGetRect(layout, path) {
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
    if (!path.includes('headerLogo')) {
    }
    return null;
  } catch (e) {
    if (!path.includes('headerLogo')) {
    }
    return null;
  }
}
function findContentRect(layout, key) {
  const candidates = [
    'body',
    'area',
    'gridArea',
    'lanesArea',
    'pyramidArea',
    'stepArea',
    'singleRow',
    'twoColLeft',
    'leftBox',
    'leftText'
  ];
  for (const name of candidates) {
    const r = safeGetRect(layout, `${key}.${name}`);
    if (r && r.top != null) return r;
  }
  return null;
}
function adjustColorBrightness(hex, factor) {
  const c = hex.replace('#', '');
  const rgb = parseInt(c, 16);
  let r = (rgb >> 16) & 0xff,
    g = (rgb >> 8) & 0xff,
    b = (rgb >> 0) & 0xff;
  r = Math.min(255, Math.round(r * factor));
  g = Math.min(255, Math.round(g * factor));
  b = Math.min(255, Math.round(b * factor));
  return '#' + (0x1000000 + (r << 16) + (g << 8) + b).toString(16).slice(1);
}
function setMainSlideBackground(slide, layout, imageUpdateOption = 'update') {
  setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.main, CONFIG.COLORS.background_white, imageUpdateOption);
}
function setBackgroundImageFromUrl(slide, layout, imageUrl, fallbackColor, imageUpdateOption = 'update') {
    slide.getBackground().setSolidFill(fallbackColor);
  if (imageUpdateOption === 'update') {
    if (!imageUrl) return;
    try {
      const image = insertImageFromUrlOrFileId(imageUrl);
      if (!image) return;
      slide.insertImage(image).setWidth(layout.pageW_pt).setHeight(layout.pageH_pt).setLeft(0).setTop(0).sendToBack();
    } catch (e) {
    }
  }
}
function insertImageFromUrlOrFileId(urlOrFileId) {
  if (!urlOrFileId) return null;
  function extractFileIdFromUrl(url) {
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
      const mimeType = parts[0].match(/:(.*?);/)[1];
      const base64Data = parts[1];
      const decodedData = Utilities.base64Decode(base64Data);
      return Utilities.newBlob(decodedData, mimeType);
    } catch (e) {
      return null;
    }
  } else {
    return urlOrFileId;
  }
}
function normalizeImages(arr) {
  return (arr || []).map(v => typeof v === 'string' ? {
    url: v
  } : (v && v.url ? v : null)).filter(Boolean).slice(0, 6);
}
function renderImagesInArea(slide, layout, area, images, imageUpdateOption = 'update') {
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
    } catch (e) {}
  }
}
function createGradientRectangle(slide, x, y, width, height, colors) {
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
    return slide.group(shapes);
  }
  return shapes[0] || null;
}
function applyFill(slide, x, y, width, height, settings) {
  if (settings.enableGradient) {
    createGradientRectangle(slide, x, y, width, height, [settings.gradientStart, settings.gradientEnd]);
  } else {
    const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, width, height);
    shape.getFill().setSolidFill(settings.primaryColor);
    shape.getBorder().setTransparent();
  }
}
function createPillShapeUnderline(slide, x, y, width, height, settings) {
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
      if (mainShape) shapes.push(mainShape);
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
      slide.group(shapes);
    } catch (e) {
    }
  }
}
function clearLegacyUserProperties() {
  try {
    const properties = PropertiesService.getUserProperties().getProperties();
    const legacyKeys = [
      'primaryColor',
      'gradientStart', 
      'gradientEnd',
      'fontFamily',
      'showTitleUnderline',
      'showBottomBar',
      'enableGradient',
      'footerText',
      'headerLogoUrl',
      'closingLogoUrl',
      'titleBgUrl',
      'sectionBgUrl',
      'mainBgUrl',
      'closingBgUrl',
      'driveFolderUrl',
      'driveFolderId'
    ];
    const keysToDelete = [];
    legacyKeys.forEach(key => {
      if (properties.hasOwnProperty(key)) {
        keysToDelete.push(key);
      }
    });
    if (keysToDelete.length > 0) {
      const userProperties = PropertiesService.getUserProperties();
      keysToDelete.forEach(key => {
        userProperties.deleteProperty(key);
      });
      return {
        status: 'success',
        message: `${keysToDelete.length}個のレガシープロパティを削除しました。`,
        deletedKeys: keysToDelete
      };
    } else {
      return {
        status: 'info',
        message: '削除対象のレガシープロパティは見つかりませんでした。'
      };
    }
  } catch (e) {
    return {
      status: 'error',
      message: `レガシープロパティの削除中にエラーが発生しました: ${e.message}`
    };
  }
}
function saveGemUrl(url) {
  try {
    const trimmedUrl = url ? url.trim() : '';
    const gemBaseUrl = 'https://gemini.google.com/gem/';
    if (!trimmedUrl) {
      throw new Error("URLを入力してください。");
    }
    // [変更] 'https://' で始まらない場合はエラーをスローして保存を拒否
    if (!trimmedUrl.startsWith('https://')) {
      throw new Error("URLは https:// で始まる必要があります。");
    }
    PropertiesService.getScriptProperties().setProperty('geminiGemUrl', trimmedUrl);
    if (!trimmedUrl.startsWith(gemBaseUrl)) {
      // 'https://' で始まっているが 'gemini.google.com/gem/' ではない場合
      return { status: 'warning', message: 'URLは保存されましたが、Gemini Gem (https://gemini.google.com/gem/...) の形式ではありません。' };
    } else if (trimmedUrl.length === gemBaseUrl.length) {
      return { status: 'warning', message: 'URLが保存されましたが、特定のGemが指定されていません。' };
    } else {
      return { status: 'success', message: 'Gemini Gem URLを保存しました。' };
    }
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}
function adjustShapeText_External(shape, preCalculatedWidthPt = null, widthOverride = null, heightOverride = null) {
  const PADDING_TOP_BOTTOM = 7.5;
  const PADDING_LEFT_RIGHT = 10.0;
  const LINE_HEIGHT_FACTOR = 1.1;
  const AVG_CHAR_WIDTH_FACTOR = 1.0;
  function getEffectiveCharCount(text) {
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
  function _isShapeShortBox(shape, baseFontSize, heightOverride = null) {
    if (!baseFontSize || baseFontSize === 0) {
      return false;
    }
    const boxHeight = (heightOverride !== null) ? heightOverride : shape.getHeight();
    return (boxHeight <= (baseFontSize * 2));
  }
  function getNonEmptyFontSize(shape) {
    try {
      const textRange = shape.getText();
      if (!textRange || textRange.isEmpty()) return null;
      const paragraphs = textRange.getParagraphs();
      for (const paragraph of paragraphs) {
        if (paragraph.getRange().asString().trim().length > 0) {
          const run = paragraph.getRange().getRuns()[0];
          if (run && run.getTextStyle()) {
            const fontSize = run.getTextStyle().getFontSize();
            if (typeof fontSize === 'number') {
                 return Math.floor(fontSize * 2) / 2;
            }
          }
        }
      }
      const firstRun = textRange.getRuns()[0];
      if (firstRun && firstRun.getTextStyle()) {
        const fontSize = firstRun.getTextStyle().getFontSize();
        if (typeof fontSize === 'number') {
             return Math.floor(fontSize * 2) / 2;
        }
      }
      return null;
    } catch (e) {
      return null;
    }
  }
  function isTextOverflowing(shape, nonEmptyFontSizeOverride = null, preCalculatedWidthPt = null, originalBaseFontSize = null, widthOverride = null, heightOverride = null) {
    let details = "--- 簡易計算の詳細 ---\n\n";
    try {
      const textRange = shape.getText();
      const originalText = textRange.asString();
      if (originalText.length === 0) {
        return { isOverflow: false, details: "テキストが空です。" };
      }
      const paragraphs = textRange.getParagraphs();
      const paragraphCount = paragraphs.length;
      let currentBaseFontSize;
      if (nonEmptyFontSizeOverride !== null) {
        currentBaseFontSize = nonEmptyFontSizeOverride;
        details += `基準フォントサイズ: ${currentBaseFontSize.toFixed(1)} pt (縮小テスト中)\n`;
      } else {
        currentBaseFontSize = getNonEmptyFontSize(shape);
        if (currentBaseFontSize === null) {
          const firstRun = textRange.getRuns()[0];
          if (firstRun && firstRun.getTextStyle()) {
              const fs = firstRun.getTextStyle().getFontSize();
              if(typeof fs === 'number') {
                  currentBaseFontSize = Math.floor(fs * 2) / 2;
              } else {
                 currentBaseFontSize = 1;
              }
          } else {
             currentBaseFontSize = 1;
          }
        }
        details += `基準フォントサイズ: ${currentBaseFontSize.toFixed(1)} pt (空でない段の代表値)\n`;
      }
      if (!currentBaseFontSize || currentBaseFontSize <= 0) {
        details += "基準フォントサイズが取得できないか無効なため、判定不能です。";
        return { isOverflow: false, details: details };
      }
      const boxWidth = (widthOverride !== null) ? widthOverride : shape.getWidth();
      const boxHeight = (heightOverride !== null) ? heightOverride : shape.getHeight();
      details += `ボックスサイズ (W): ${boxWidth.toFixed(2)} pt\n`;
      details += `ボックスサイズ (H): ${boxHeight.toFixed(2)} pt\n`;
      details += `推定余白 (上下): ${PADDING_TOP_BOTTOM} pt\n`;
      details += `推定余白 (左右): ${PADDING_LEFT_RIGHT} pt\n\n`;
      const textAreaWidth = boxWidth - (PADDING_LEFT_RIGHT * 2);
      const textAreaHeight = boxHeight - (PADDING_TOP_BOTTOM * 2);
      details += `推定テキストエリア内寸 (W): ${textAreaWidth.toFixed(2)} pt\n`;
      details += `推定テキストエリア内寸 (H): ${textAreaHeight.toFixed(2)} pt\n`;
      if (textAreaWidth <= 0 || textAreaHeight <= 0) {
        details += "判定: 内寸が0以下のため「はみ出し」";
        return { isOverflow: true, details: details };
      }
      const initialSizeForModeCheck = originalBaseFontSize !== null ? originalBaseFontSize : currentBaseFontSize;
      const isShortBox = _isShapeShortBox(shape, initialSizeForModeCheck, heightOverride);
      details += `ボックスタイプ: ${isShortBox ? "単一行モード" : "標準（複数行）モード"} (判定基準FS: ${initialSizeForModeCheck.toFixed(1)}pt)\n\n`;
      let isOverflow;
      if (isShortBox) {
        const fontSizeToUse = currentBaseFontSize;
        let calculatedTextWidthPt = -1; 
        if (nonEmptyFontSizeOverride !== null && preCalculatedWidthPt !== null && preCalculatedWidthPt >= 0 && originalBaseFontSize !== null && originalBaseFontSize > 0) {
          calculatedTextWidthPt = preCalculatedWidthPt * (fontSizeToUse / originalBaseFontSize); 
          details += `使用するテキスト幅: ${calculatedTextWidthPt.toFixed(2)} pt (${fontSizeToUse.toFixed(1)}pt で事前計算値をスケーリング)\n`;
        }
        else if (preCalculatedWidthPt !== null && preCalculatedWidthPt >= 0) {
          calculatedTextWidthPt = preCalculatedWidthPt;
          details += `使用するテキスト幅: ${calculatedTextWidthPt.toFixed(2)} pt (事前計算値)\n`;
        }
        else {
          const estimatedFullCharWidth = fontSizeToUse * AVG_CHAR_WIDTH_FACTOR; 
          const totalEffectiveCount = getEffectiveCharCount(originalText); 
          calculatedTextWidthPt = totalEffectiveCount * estimatedFullCharWidth; 
          details += `使用するテキスト幅: ${calculatedTextWidthPt.toFixed(2)} pt (${fontSizeToUse.toFixed(1)}pt で文字数ベース推定)\n`;
        }
        isOverflow = calculatedTextWidthPt > textAreaWidth; 
        const hasManualLineBreaks = (paragraphCount > 1); 
        isOverflow = isOverflow || hasManualLineBreaks; 
        details += `\n--- 判定 (単一行モード) ---\n`; 
        details += `(使用フォントサイズ: ${fontSizeToUse.toFixed(1)} pt)\n`; 
        details += `(計算/推定テキスト幅: ${calculatedTextWidthPt.toFixed(2)} pt)\n`; 
        details += `(テキストエリア内寸幅: ${textAreaWidth.toFixed(2)} pt)\n`; 
        details += `手動改行あり: ${hasManualLineBreaks}\n`; 
        details += `幅オーバー: ${calculatedTextWidthPt > textAreaWidth}\n`; 
        details += `結果: ${isOverflow ? "はみ出し" : "収まっている"}`; 
      } else {
        details += `\n--- 判定 (標準モード・段落ごと) ---\n`; 
        let calculatedTotalHeight = 0; 
        let paraIndex = 0; 
        for (const paragraph of paragraphs) {
            const paraRange = paragraph.getRange(); 
            const paraText = paraRange.asString(); 
            const isParaEmpty = paraText.trim().length === 0; 
            let currentFontSize;
            if (isParaEmpty) { 
              if (nonEmptyFontSizeOverride !== null) {
                currentFontSize = 1; 
              } else {
                const run = paraRange.getRuns()[0];
                let fs = run && run.getTextStyle() ? run.getTextStyle().getFontSize() : null;
                currentFontSize = (typeof fs === 'number') ? Math.floor(fs * 2) / 2 : 1;
              }
            } else { 
              if (nonEmptyFontSizeOverride !== null) {
                currentFontSize = nonEmptyFontSizeOverride; 
              } else {
                const run = paraRange.getRuns()[0];
                let fs = run && run.getTextStyle() ? run.getTextStyle().getFontSize() : null;
                currentFontSize = (typeof fs === 'number') ? Math.floor(fs * 2) / 2 : currentBaseFontSize;
              }
            }
             currentFontSize = Math.max(1, currentFontSize);
            details += `  段落${paraIndex + 1}: ${isParaEmpty ? "(空)" : ""} FS=${currentFontSize.toFixed(1)}pt\n`; 
            const estimatedLineHeight = ( currentFontSize * LINE_HEIGHT_FACTOR ) + 0.5; 
            const estimatedFullCharWidth = currentFontSize * AVG_CHAR_WIDTH_FACTOR; 
            const maxEffectiveCharsPerLine = (estimatedFullCharWidth > 0) ? Math.floor(textAreaWidth / estimatedFullCharWidth) : 0;
            let linesInThisParagraph = 0; 
            if (isParaEmpty) {
                linesInThisParagraph = 1; 
            } else {
                const effectiveCount = getEffectiveCharCount(paraText); 
                linesInThisParagraph = (maxEffectiveCharsPerLine > 0) ? Math.ceil(effectiveCount / maxEffectiveCharsPerLine) : 1;
                 linesInThisParagraph = Math.max(1, linesInThisParagraph); 
            }
            const paragraphStyle = paraRange.getParagraphStyle(); 
            const spaceAbove = (paragraphStyle && typeof paragraphStyle.getSpaceAbove === 'function' ? paragraphStyle.getSpaceAbove() : 0) || 0;
            const spaceBelow = (paragraphStyle && typeof paragraphStyle.getSpaceBelow === 'function' ? paragraphStyle.getSpaceBelow() : 0) || 0;
            const paragraphSpacing = spaceAbove + spaceBelow; 
            const calculatedLinesHeight = linesInThisParagraph * estimatedLineHeight;
            calculatedTotalHeight += calculatedLinesHeight;
            details += `    -> ${linesInThisParagraph}行 * ${estimatedLineHeight.toFixed(2)}pt = ${calculatedLinesHeight.toFixed(2)}pt (文字数ベース推定)\n`; 
            if (paraIndex > 0) { 
              calculatedTotalHeight += paragraphSpacing;
              details += `    -> 段落間隔 ${paragraphSpacing.toFixed(2)}pt 加算\n`; 
            }
            paraIndex++; 
        }
        details += `\n算出された必要な総高さ: ${calculatedTotalHeight.toFixed(2)} pt\n`; 
        isOverflow = calculatedTotalHeight > textAreaHeight;
        details += `必要な高さ (${calculatedTotalHeight.toFixed(2)} pt) > エリアの高さ (${textAreaHeight.toFixed(2)} pt)\n`; 
        details += `結果: ${isOverflow ? "はみ出し" : "収まっている"}`; 
      }
      return { isOverflow: isOverflow, details: details }; 
    } catch (e) {
      return { isOverflow: false, details: "エラーが発生しました: " + e.message };
    }
  }
  function minimizeEmptyParagraphs(shape) {
    let modified = false; 
    try {
      const textRange = shape.getText(); 
      if (!textRange || textRange.isEmpty()) return false; 
      const paragraphs = textRange.getParagraphs(); 
      for (const paragraph of paragraphs) { 
        const paraRange = paragraph.getRange(); 
        if (paraRange.asString().trim().length === 0) {
          const style = paraRange.getTextStyle(); 
          if (style && style.getFontSize() !== 1) {
            style.setFontSize(1); 
            modified = true; 
          }
        }
      }
    } catch (e) {
    }
    return modified; 
  }
  function applyNonEmptyFontSize(shape, size) {
    try {
      const textRange = shape.getText(); 
      if (!textRange || textRange.isEmpty()) return; 
      const paragraphs = textRange.getParagraphs(); 
      for (const paragraph of paragraphs) { 
        const paraRange = paragraph.getRange(); 
        if (paraRange.asString().trim().length > 0) {
           const style = paraRange.getTextStyle(); 
           if (style) { 
               style.setFontSize(size);
           }
        }
      }
    } catch (e) {
    }
  }
  function findOptimalFontSize(shape, startFontSize, preCalculatedWidthPt = null, originalBaseFontSize = null, widthOverride = null, heightOverride = null) {
    if (typeof startFontSize !== 'number' || startFontSize < 10) {
        return null; 
    }
    const MIN_FONT_SIZE = (widthOverride !== null) ? 1 : 10;
    for (let testSize = startFontSize - 0.5; testSize >= MIN_FONT_SIZE; testSize -= 0.5) {
      const result = isTextOverflowing(shape, testSize, preCalculatedWidthPt, originalBaseFontSize, widthOverride, heightOverride);
      if (!result.isOverflow) { 
        return testSize; 
      }
    }
    return null; 
  }
  let logDetails = ""; 
  let finalAppliedFontSize = null; 
  try {
    if (!shape || typeof shape.getText !== 'function' || shape.getText().isEmpty()) { 
      return { log: "テキストが空のため、処理をスキップしました。", finalSize: null }; 
    }
    logDetails += "--- 事前チェック (面積比較) ---\n"; 
    const textRange = shape.getText(); 
    const originalText = textRange.asString(); 
    const totalCharCount = originalText.length; 
    const initialBaseFontSize = getNonEmptyFontSize(shape); 
    finalAppliedFontSize = initialBaseFontSize; 
    if (initialBaseFontSize === null) { 
      logDetails += "エラー: 基準フォントサイズが取得できず、事前チェックをスキップします。\n\n";
    } else {
      const boxWidth = (widthOverride !== null) ? widthOverride : shape.getWidth();
      const boxHeight = (heightOverride !== null) ? heightOverride : shape.getHeight();
      const textAreaWidth = boxWidth - (PADDING_LEFT_RIGHT * 2);
      const textAreaHeight = boxHeight - (PADDING_TOP_BOTTOM * 2);
      const textAreaArea = textAreaWidth * textAreaHeight;
      if (textAreaArea <= 0) { 
        logDetails += "警告: テキストエリアの内寸面積が0以下です。詳細チェックに進みます。\n\n";
      } else {
        const estimatedCharArea = initialBaseFontSize * initialBaseFontSize;
        const estimatedTotalCharArea = estimatedCharArea * totalCharCount;
        const thresholdArea = textAreaArea * 0.4;
        logDetails += `基準フォントサイズ: ${initialBaseFontSize.toFixed(1)} pt\n`; 
        logDetails += `総文字数: ${totalCharCount} 文字\n`;
        logDetails += `推定1文字の面積 (FS*FS): ${estimatedCharArea.toFixed(2)} pt^2\n`;
        logDetails += `推定テキスト総面積: ${estimatedTotalCharArea.toFixed(2)} pt^2\n`;
        logDetails += `推定テキストエリア内寸面積: ${textAreaArea.toFixed(2)} pt^2\n`;
        logDetails += `判定閾値 (エリアの40%): ${thresholdArea.toFixed(2)} pt^2\n`;
        if (estimatedTotalCharArea <= thresholdArea) {
          logDetails += "結果: テキスト総面積がエリアの40%以下のため、詳細チェックをスキップします。✅\n";
          return { log: logDetails, finalSize: initialBaseFontSize }; 
        } else {
          logDetails += "結果: テキスト総面積がエリアの40%を超えているため、詳細チェックに進みます。\n\n";
        }
      }
    }
    const initialResult = isTextOverflowing(shape, null, preCalculatedWidthPt, initialBaseFontSize, widthOverride, heightOverride);
    logDetails += "--- 詳細診断 (実行前の状態) --- \n" + initialResult.details + "\n\n"; 
    if (!initialResult.isOverflow) {
      logDetails += "--- 実行結果 --- \nはみ出しは検出されませんでした。調整は不要です。 ✅"; 
      return { log: logDetails, finalSize: initialBaseFontSize }; 
    }
    logDetails += "--- 調整実行 --- \nはみ出しを検出しました 🚨\n\n"; 
    const currentActualFontSize = getNonEmptyFontSize(shape); 
     if (currentActualFontSize === null) { 
        logDetails += "エラー: 現在のフォントサイズが取得できず、調整を実行できません。";
        return { log: logDetails, finalSize: null}; 
    }
    const isShortBox = _isShapeShortBox(shape, initialBaseFontSize || currentActualFontSize, heightOverride);
    if (isShortBox) {
      logDetails += "（単一行モードとして処理します）\n"; 
      const optimalSize = findOptimalFontSize(shape, currentActualFontSize, preCalculatedWidthPt, initialBaseFontSize, widthOverride, heightOverride);
      if (optimalSize !== null) { 
        shape.getText().getTextStyle().setFontSize(optimalSize);
        finalAppliedFontSize = optimalSize; 
        logDetails += `調整： 全体のフォントサイズを ${optimalSize.toFixed(1)} pt に縮小しました。\n`; 
        const finalResult = isTextOverflowing(shape, optimalSize, preCalculatedWidthPt, initialBaseFontSize, widthOverride, heightOverride);
        logDetails += "\n--- 調整後の計算詳細 ---\n" + (finalResult.details || "計算失敗"); 
      } else {
        const MIN_FONT_SIZE = (widthOverride !== null) ? 1 : 10;
        shape.getText().getTextStyle().setFontSize(MIN_FONT_SIZE); 
        finalAppliedFontSize = MIN_FONT_SIZE; 
        logDetails += `調整： フォントサイズ ${MIN_FONT_SIZE} pt でもテキストが収まりませんでした。\n`; 
      }
    } else {
      logDetails += "（標準モードとして処理します）\n"; 
      const modifiedEmpty = minimizeEmptyParagraphs(shape);
      if (modifiedEmpty) {
        logDetails += "ステップ1: 空の段落のフォントサイズを 1pt に縮小しました。\n"; 
      } else {
        logDetails += "ステップ1: 対象となる空の段落はありませんでした。\n"; 
      }
      const resultAfterStep1 = isTextOverflowing(shape, null, preCalculatedWidthPt, initialBaseFontSize, widthOverride, heightOverride);
      logDetails += "ステップ1実行後の再計算結果：\n" + (resultAfterStep1.details || "計算失敗") + "\n\n"; 
      if (!resultAfterStep1.isOverflow) {
        logDetails += "--- 実行結果 --- \n空の段落の縮小のみで、はみ出しが解消しました。✅"; 
        return { log: logDetails, finalSize: initialBaseFontSize }; 
      }
      logDetails += "ステップ2: 文字を含む段落のフォントサイズを縮小します。\n"; 
      const currentNonEmptySizeAfterStep1 = getNonEmptyFontSize(shape); 
       if (currentNonEmptySizeAfterStep1 === null) { 
           logDetails += "エラー: 現在のフォントサイズが取得できず、ステップ2を実行できません。";
           return { log: logDetails, finalSize: null}; 
       }
      const optimalSize = findOptimalFontSize(shape, currentNonEmptySizeAfterStep1, preCalculatedWidthPt, initialBaseFontSize, widthOverride, heightOverride);
      if (optimalSize !== null) { 
        applyNonEmptyFontSize(shape, optimalSize);
        finalAppliedFontSize = optimalSize; 
        logDetails += `調整： 文字のある段落を ${optimalSize.toFixed(1)} pt に縮小しました。\n`; 
        const finalResult = isTextOverflowing(shape, null, preCalculatedWidthPt, initialBaseFontSize, widthOverride, heightOverride);
        logDetails += "\n--- 調整後の計算詳細 ---\n" + (finalResult.details || "計算失敗"); 
      } else {
        const MIN_FONT_SIZE = (widthOverride !== null) ? 1 : 10;
        applyNonEmptyFontSize(shape, MIN_FONT_SIZE); 
        finalAppliedFontSize = MIN_FONT_SIZE; 
        logDetails += `調整： 文字のある段落を ${MIN_FONT_SIZE}pt にしてもテキストが収まりませんでした。\n`; 
      }
    }
    return { log: logDetails, finalSize: finalAppliedFontSize };
  } catch (e) {
    if (e.message.includes("getWidth is not a function") || e.message.includes("getHeight is not a function") || e.message.includes("shape.getWidth is not a function")) {
        return { log: logDetails + "\n\n--- 致命的なエラー --- \n" + "TableCellが渡されましたが、幅または高さのOverride引数がありませんでした。 " + e.message, finalSize: null };
    }
    return { log: logDetails + "\n\n--- 致命的なエラー --- \n" + e.message, finalSize: null };
  }
}
function processImagesAndUpload(imagesToUpload, slideDataString, settings) {
  if (!imagesToUpload || imagesToUpload.length === 0) {
    return {
      updatedSlideDataString: slideDataString,
      imageFolderId: null,
      error: null
    };
  }
  try {
    const parentFolderId = settings.driveFolderId || 'root';
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const tempFolderName = `temp_images_${new Date().getTime()}`;
    const tempFolder = parentFolder.createFolder(tempFolderName);
    const tempFolderId = tempFolder.getId();
    const imageIdMap = {};
    imagesToUpload.forEach((image, index) => {
      const base64Data = image.data.split(',')[1];
      const decodedData = Utilities.base64Decode(base64Data);
      const contentType = image.data.match(/^data:(.*?);/)[1] || 'image/jpeg';
      const fileName = `image_${index}_${image.key.substring(0,20).replace(/[^a-zA-Z0-9]/g,'_')}.png`;
      const blob = Utilities.newBlob(decodedData, contentType, fileName);
      const file = tempFolder.createFile(blob);
      const fileId = file.getId();
      imageIdMap[image.key] = fileId; 
    });
    let slideData = JSON.parse(slideDataString);
    const updatedSlideData = slideData.map(slide => {
      let currentImageKey = null;
      if (slide.image) {
          if (typeof slide.image === 'string') {
              currentImageKey = slide.image;
          } else if (typeof slide.image === 'object' && slide.image !== null) {
              try {
                currentImageKey = JSON.stringify(slide.image);
              } catch (e) {
              }
          }
      }
      if (currentImageKey && imageIdMap[currentImageKey]) {
          return { ...slide, image: imageIdMap[currentImageKey] };
      } else {
          if (currentImageKey) {
          }
          return slide; 
      }
    });
    const finalJsonString = JSON.stringify(updatedSlideData);
    return {
      updatedSlideDataString: finalJsonString,
      imageFolderId: tempFolderId,
      error: null
    };
  } catch (e) {
    return {
      updatedSlideDataString: slideDataString,
      imageFolderId: null,
      error: 'サーバーサイドで画像のアップロードまたは処理中にエラーが発生しました: ' + e.message
    };
  }
}
function deleteFolderById(folderId) {
  try {
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      folder.setTrashed(true);
    }
  } catch (e) {
  }
}
function checkUserActivation() {
  const SCRIPT_PROPERTY_KEY = 'ACTIVATED_USER_LIST';
  let userEmail = '';
  let isActivated = false;
  try {
    userEmail = Session.getEffectiveUser().getEmail();
    const scriptProperties = PropertiesService.getScriptProperties();
    const jsonString = scriptProperties.getProperty(SCRIPT_PROPERTY_KEY);
    if (jsonString) {
      const activatedUsers = JSON.parse(jsonString);
      if (Array.isArray(activatedUsers) && activatedUsers.includes(userEmail)) {
        isActivated = true;
      }
    }
  } catch (e) {
    isActivated = false;
  }
  return {
    isActivated: isActivated,
    userEmail: userEmail
  };
}
function activateUser(userEmail) {
  const SCRIPT_PROPERTY_KEY = 'ACTIVATED_USER_LIST';
  if (!userEmail) {
    return { status: 'error', message: 'ユーザーIDがありません。' };
  }
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const jsonString = scriptProperties.getProperty(SCRIPT_PROPERTY_KEY);
    let activatedUsers = [];
    if (jsonString) {
      try {
        activatedUsers = JSON.parse(jsonString);
        if (!Array.isArray(activatedUsers)) {
          activatedUsers = [];
        }
      } catch (e) {
        activatedUsers = [];
      }
    }
    if (!activatedUsers.includes(userEmail)) {
      activatedUsers.push(userEmail);
      scriptProperties.setProperty(SCRIPT_PROPERTY_KEY, JSON.stringify(activatedUsers));
    } else {
    }
    return { status: 'success', message: 'アクティベートが完了しました。' };
  } catch (e) {
    return { status: 'error', message: 'サーバーエラー: ' + e.message };
  }
}
function convertJsonToSvgBatch(jsonStrings) {
  if (!Array.isArray(jsonStrings)) {
    return []; 
  }
  const results = jsonStrings.map(jsonText => {
    try {
      const svgString = jsonToSVG(jsonText);
      return {
        key: jsonText,   
        svg: svgString, 
        error: null
      };
    } catch (e) {
      return {
        key: jsonText,
        svg: null,
        error: e.message
      };
    }
  });
  return results; 
}
function jsonToSVG(jsonText) {
  let config;
  try {
    config = JSON.parse(jsonText);
  } catch (e) {
    throw new Error(`無効なJSONが提供されました: ${e.message}`);
  }
  if (!config || !config.chartType) {
    throw new Error("無効なJSON構造です: 'chartType' キーが見つかりません。");
  }
  const chartType = config.chartType;
  try {
    const template = getTemplate(chartType);
    const svgOutput = injectJsonIntoSvg(template, jsonText);
    return svgOutput;
  } catch (e) {
    throw e;
  }
}
function injectJsonIntoSvg(template, jsonText) {
  const regex = /(<script id="chart-json-data"[^>]*>)([\s\S]*?)(<\/script>)/;
  if (!regex.test(template)) {
    throw new Error("SVGテンプレートにターゲットのscriptタグ (id='chart-json-data') が見つかりません。");
  }
  const replacedSvg = template.replace(regex, `$1\n    ${jsonText}\n    $3`);
  return replacedSvg;
}
function getTemplate(chartType) {
  switch (chartType) {
    case "combo":
      return COMBO_CHART_TEMPLATE;
    case "multi-line":
      return MULTI_LINE_CHART_TEMPLATE;
    case "stacked-bar":
      return STACKED_BAR_CHART_TEMPLATE;
    case "bar":
      return BAR_CHART_TEMPLATE;
    case "100-stacked-bar":
      return PERCENT_STACKED_BAR_CHART_TEMPLATE;
    case "line":
      return LINE_CHART_TEMPLATE;
    case "donut":
      return DONUT_CHART_TEMPLATE;
    default:
      throw new Error(`不明な chart type です: '${chartType}'`);
  }
}
function getTriangleArrowSvgTemplate() {
  try {
    return TRIANGLE_ARROW_SVG_TEMPLATE;
  } catch (e) {
    return '';
  }
}
const TRIANGLE_ARROW_SVG_TEMPLATE = `<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<svg
   width="500px"
   height="500px"
   viewBox="0 0 500 500"
   version="1.1"
   id="svg1"
   xmlns="http://www.w3.org/2000/svg"
   xmlns:svg="http://www.w3.org/2000/svg">
  <defs id="defs1">
    <style id="variable-controller">
      :root {
        --end-x: {{END_X}};
        --end-y: {{END_Y}};
        --curve-intensity: {{CURVE_INTENSITY}};
        --rotation: {{ROTATION}};
        --flip-x: {{FLIP_X}};
        --flip-y: {{FLIP_Y}};
        --ghost-base-frequency: {{GHOST_BASE_FREQUENCY}};
        --ghost-color-type: '{{GHOST_COLOR_TYPE}}';
        --ghost-alpha-slope: {{GHOST_ALPHA_SLOPE}};
        --ghost-contrast-slope: {{GHOST_CONTRAST_SLOPE}};
        --ghost-front-color: '{{GHOST_FRONT_COLOR}}';
        --ghost-front-alpha: {{GHOST_FRONT_ALPHA}};
      }
    </style>
    <filter id="ghostNoiseFilter"
       filterUnits="userSpaceOnUse"
       x="0" y="0" width="500" height="500">
      <feTurbulence id="ghostTurbulence"
          type="fractalNoise"
          baseFrequency="{{GHOST_BASE_FREQUENCY}}"
          numOctaves="3"
          seed="0"
          result="noise"/>
      <feColorMatrix id="ghostColorMatrix" in="noise" type="matrix"
          values="0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0 0 0 1 0"
          result="coloredNoise"/>
      <feComposite id="ghostComposite" in="coloredNoise" in2="SourceAlpha" operator="in" result="maskedNoise"/>
      <feComponentTransfer in="maskedNoise" result="finalTexture">
          <feFuncR id="ghostFuncR" type="linear" slope="{{GHOST_CONTRAST_SLOPE}}" intercept="0"/>
          <feFuncG id="ghostFuncG" type="linear" slope="{{GHOST_CONTRAST_SLOPE}}" intercept="0"/>
          <feFuncB id="ghostFuncB" type="linear" slope="{{GHOST_CONTRAST_SLOPE}}" intercept="0"/>
          <feFuncA id="ghostFuncA" type="linear" slope="{{GHOST_ALPHA_SLOPE}}" intercept="0"/>
      </feComponentTransfer>
      <feMerge>
        <feMergeNode in="finalTexture"/>
      </feMerge>
    </filter>
    <mask id="arrowMask">
      <g id="arrow-mask-layer" style="fill:white; stroke:white;">
        <path
           id="arrow-head-mask"
           d="M 0 -43.30127 L -25 0 H 25 Z"
           style="fill-opacity:1; stroke:none;"
           transform="translate(103.28583, 117.1908) rotate(0)" />
        <path
           id="arrow-stem-curve-mask"
           d=""
           fill="none"
           stroke-width="21.566"
           stroke-linecap="butt"
           stroke-linejoin="bevel"
           stroke-opacity="1" />
      </g>
    </mask>
  </defs>
  <g id="layer1">
    <g id="arrow-ghost-layer" filter="url(#ghostNoiseFilter)">
        <path
           id="arrow-head-ghost"
           d="M 0 -43.30127 L -25 0 H 25 Z"
           style="fill:#000000;fill-opacity:1;stroke:none;"
           transform="translate(103.28583, 117.1908) rotate(0)" />
        <path
           id="arrow-stem-curve-ghost"
           d=""
           fill="none"
           stroke="#000000"
           stroke-width="21.566"
           stroke-linecap="butt"
           stroke-linejoin="bevel"
           stroke-opacity="1" />
    </g>
    <g id="arrow-front-layer">
        <path
           id="arrow-head-front"
           d="M 0 -43.30127 L -25 0 H 25 Z"
           style="fill:{{GHOST_FRONT_COLOR}};fill-opacity:{{GHOST_FRONT_ALPHA}};stroke:none;"
           transform="translate(103.28583, 117.1908) rotate(0)" />
        <path
           id="arrow-stem-curve-front"
           d=""
           fill="none"
           stroke="{{GHOST_FRONT_COLOR}}"
           stroke-width="21.566"
           stroke-linecap="butt"
           stroke-linejoin="bevel"
           stroke-opacity="{{GHOST_FRONT_ALPHA}}" />
    </g>
  </g>
  <script>
      (function() {
        const root = document.documentElement;
        const arrowHeadGhost = document.getElementById('arrow-head-ghost');
        const curvePathGhost = document.getElementById('arrow-stem-curve-ghost');
        const arrowHeadFront = document.getElementById('arrow-head-front');
        const curvePathFront = document.getElementById('arrow-stem-curve-front');
        const feTurbulence = document.getElementById('ghostTurbulence');
        const feColorMatrix = document.getElementById('ghostColorMatrix');
        const feComposite = document.getElementById('ghostComposite');
        const feFuncR = document.getElementById('ghostFuncR');
        const feFuncG = document.getElementById('ghostFuncG');
        const feFuncB = document.getElementById('ghostFuncB');
        const feFuncA = document.getElementById('ghostFuncA');
        const startX = 100; const startY = 100;
        const canvasMin = 0; const canvasMax = 500;
        const margin = 10;
        const endPointMin = canvasMin + margin; const endPointMax = canvasMax - margin;
        function getCssVarNum(name, fallback) { const value = getComputedStyle(root).getPropertyValue(name).trim(); const parsedValue = parseFloat(value); return isNaN(parsedValue) ? fallback : parsedValue; }
        function getCssVarString(name, fallback) { const value = getComputedStyle(root).getPropertyValue(name).trim(); return value.replace(/^['"]|['"]$/g, '') || fallback; }
        function updatePath() {
          const relativeX_base = getCssVarNum('--end-x', 0);
          let relativeY_base = getCssVarNum('--end-y', 60);
          let intensity_base = getCssVarNum('--curve-intensity', 0);
          relativeY_base = relativeY_base * -1;
          intensity_base = Math.max(-100, Math.min(intensity_base, 100));
          const rotation = getCssVarNum('--rotation', 0);
          const flipX = (getCssVarNum('--flip-x', 0) === 1) ? -1 : 1;
          const flipY = (getCssVarNum('--flip-y', 0) === 1) ? -1 : 1;
          const baseFrequency = getCssVarNum('--ghost-base-frequency', 0.01);
          const colorType = getCssVarString('--ghost-color-type', 'gray');
          const alphaSlope = getCssVarNum('--ghost-alpha-slope', 0.50);
          const contrastSlope = getCssVarNum('--ghost-contrast-slope', 1.0);
          const frontColor = getCssVarString('--ghost-front-color', '#000000');
          const frontAlpha = getCssVarNum('--ghost-front-alpha', 0.15);
          const rX_flipped = relativeX_base * flipX;
          const rY_flipped = relativeY_base * flipY;
          const intensity = intensity_base * flipX * flipY;
          const rad = rotation * Math.PI / 180;
          const cos_r = Math.cos(rad); const sin_r = Math.sin(rad);
          const relativeX = rX_flipped * cos_r - rY_flipped * sin_r;
          const relativeY = rX_flipped * sin_r + rY_flipped * cos_r;
          let endX = startX + relativeX; let endY = startY + relativeY;
          endX = Math.max(endPointMin, Math.min(endX, endPointMax)); endY = Math.max(endPointMin, Math.min(endY, endPointMax));
          const midX = (startX + endX) / 2; const midY = (startY + endY) / 2;
          const vx = endX - startX; const vy = endY - startY; const len = Math.sqrt(vx * vx + vy * vy);
          let controlX, controlY;
          if (len === 0 || intensity === 0) { controlX = midX; controlY = midY; } else { const normX = -vy / len; const normY = vx / len; controlX = midX + normX * intensity; controlY = midY + normY * intensity; }
          const d = \`M \${startX},\${startY} Q \${controlX},\${controlY} \${endX},\${endY}\`;
          curvePathGhost.setAttribute('d', d); curvePathFront.setAttribute('d', d);
          const tangentVx = controlX - startX;
          const tangentVy = controlY - startY;
          const angleRad = Math.atan2(startY - controlY, startX - controlX);
          const angleDeg = (angleRad * 180 / Math.PI) + 90;
          const transform = \`translate(\${startX}, \${startY}) rotate(\${angleDeg})\`;
          arrowHeadGhost.setAttribute('transform', transform); arrowHeadFront.setAttribute('transform', transform);
          feTurbulence.setAttribute('baseFrequency', baseFrequency);
          feTurbulence.setAttribute('seed', Math.floor(Math.random() * 1000));
          let matrixValues = "0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0.33 0.33 0.33 0 0 0 0 0 1 0";
          if (colorType !== 'fullcolor') {
              if (colorType === 'red') matrixValues = "1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0";
              else if (colorType === 'green') matrixValues = "0 0 0 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 1 0";
              else if (colorType === 'blue') matrixValues = "0 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 1 0";
              feColorMatrix.setAttribute('values', matrixValues);
              feColorMatrix.style.display = 'block'; feComposite.setAttribute('in', 'coloredNoise');
          } else {
              feColorMatrix.style.display = 'none'; feComposite.setAttribute('in', 'noise');
          }
          const contrastIntercept = (0.5 * (1 - contrastSlope)).toFixed(2);
          feFuncR.setAttribute('slope', contrastSlope); feFuncR.setAttribute('intercept', contrastIntercept);
          feFuncG.setAttribute('slope', contrastSlope); feFuncG.setAttribute('intercept', contrastIntercept);
          feFuncB.setAttribute('slope', contrastSlope); feFuncB.setAttribute('intercept', contrastIntercept);
          feFuncA.setAttribute('slope', alphaSlope);
          arrowHeadFront.style.fill = frontColor; arrowHeadFront.style.fillOpacity = frontAlpha;
          curvePathFront.style.stroke = frontColor; curvePathFront.style.strokeOpacity = frontAlpha;
        }
        updatePath();
      })();
  </script>
</svg>`;
const COMBO_CHART_TEMPLATE = `
<svg width="600" height="510" viewBox="25 25 580 510" xmlns="http://www.w3.org/2000/svg" font-family="'Noto Sans JP', sans-serif" id="combo-chart-svg">
    <script id="chart-json-data" type="application/json">
        {
            "chartType": "combo",
            "data": {
                "title": "サンプル複合グラフ",
                "subtitle": "（系列データの比較）",
                "source": "出典：サンプルデータソース",
                "legendBarLabel": "系列 A",
                "legendLineLabel": "系列 B",
                "yAxisLeftLabel": "（数量）",
                "yAxisRightLabel": "（割合）",
                "colors": {
                    "bar": {
                        "start": "#e68a9c",
                        "end": "#b469b8"
                    },
                    "line": "#6b5ce0"
                },
                "items": [
                    {
                        "label": "項目 1",
                        "barValue": 245,
                        "lineValue": 16
                    },
                    {
                        "label": "項目 2",
                        "barValue": 270,
                        "lineValue": 19
                    },
                    {
                        "label": "項目 3",
                        "barValue": 310,
                        "lineValue": 21
                    },
                    {
                        "label": "項目 4",
                        "barValue": 290,
                        "lineValue": 18
                    },
                    {
                        "label": "項目 5",
                        "barValue": 300,
                        "lineValue": 23
                    }
                ],
                "layout": {
                    "width": 600,
                    "height": 510,
                    "marginTop": 180,
                    "marginBottom": 50,
                    "marginLeft": 70,
                    "marginRight": 70
                },
                "barOptions": {
                    "barToSlotRatio": 0.55,
                    "labelPosition": "auto"
                },
                "lineOptions": {
                    "markerRadius": 5
                },
                "yAxisLeft": {
                    "max": 400,
                    "min": 0,
                    "tickCount": 4
                },
                "yAxisRight": {
                    "max": 25,
                    "min": 15,
                    "tickCount": 4,
                    "unit": "%"
                },
                "animation": 1
            }
        }
    </script>
    <defs>
        <filter id="shadow">
            <feDropShadow dx="1" dy="2" stdDeviation="2" flood-color="#000" flood-opacity="0.2" />
        </filter>
         <filter id="text-shadow-dark" x="-50%" y="-50%" width="200%" height="200%">
            <feDropShadow dx="0" dy="0" stdDeviation="0.9" flood-color="#000000" flood-opacity="0.9" in="SourceAlpha" result="shadow" />
            <feMerge>
                <feMergeNode in="shadow" />
                <feMergeNode in="SourceGraphic" />
            </feMerge>
        </filter>
         <filter id="text-halo-white" x="-50%" y="-50%" width="200%" height="200%">
            <feDropShadow dx="0" dy="0" stdDeviation="2.5" flood-color="white" flood-opacity="0.95" in="SourceAlpha" result="shadow" />
            <feMerge>
                <feMergeNode in="shadow" />
                <feMergeNode in="SourceGraphic" />
            </feMerge>
        </filter>
     </defs>
    <style>
        <![CDATA[
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&amp;family=Roboto:wght@400;700&amp;display=swap');
        .title {
            font-size: 24px;
            font-weight: 500;
            text-anchor: middle;
            fill: #5f6368;
        }
        .subtitle {
            font-size: 16px;
            text-anchor: middle;
            fill: #5f6368;
        }
        .source {
            font-size: 11px;
            fill: #70757a;
            text-anchor: end;
        }
        .legend-item {
            font-size: 13px;
            fill: #202124;
        }
        .axis-label {
            font-size: 13px;
            fill: #5f6368;
            text-anchor: middle;
        }
        .grid-line {
            stroke: #e0e0e0;
            stroke-dasharray: 2 2;
            stroke-width: 0.8;
        }
        .y-axis-value {
            font-family: 'Roboto', sans-serif;
            font-size: 12px;
            fill: #757575;
        }
        .data-label-bar-inside {
            font-family: 'Roboto', sans-serif;
            font-size: 14px;
            font-weight: bold;
            fill: white; 
            text-anchor: middle;
            filter: url(#text-shadow-dark);
        }
        .data-label-bar-outside {
            font-family: 'Roboto', sans-serif;
            font-size: 14px;
            font-weight: bold;
            fill: #5f6368;
            text-anchor: middle;
            filter: url(#text-halo-white);
        }
        ]]>
    </style>
    <g id="title-group"></g>
    <g id="legend-group"></g>
    <g id="y-axes-group"></g>
    <g id="x-axis-group"></g>
    <g id="bar-chart-group"></g>
    <g id="line-chart-group"></g>
    <g id="source-group"></g>
    <script type="text/javascript">
        const svgNS = "http://www.w3.org/2000/svg";
        function createSVGElement(name, attributes, textContent) {
            const el = document.createElementNS(svgNS, name);
            for (const key in attributes) el.setAttribute(key, attributes[key]);
            if (textContent !== undefined && textContent !== null) el.textContent = textContent;
            return el;
        }
        function truncateText(text, maxWidth, className) {
            const svgRoot = document.getElementById('combo-chart-svg');
            const tempText = createSVGElement('text', {
                class: className,
                style: 'visibility: hidden;'
            });
            svgRoot.appendChild(tempText);
            tempText.textContent = text;
            if (tempText.getComputedTextLength() <= maxWidth) {
                svgRoot.removeChild(tempText);
                return text;
            }
            let truncatedText = text;
            while (truncatedText.length > 0) {
                truncatedText = truncatedText.slice(0, -1);
                tempText.textContent = truncatedText + '...';
                if (tempText.getComputedTextLength() <= maxWidth) {
                    svgRoot.removeChild(tempText);
                    return truncatedText + '...';
                }
            }
            svgRoot.removeChild(tempText);
            return '...';
        }
        function hexToRgb(hex) {
          let shorthandRegex = /^#?([a-f0-9])([a-f0-9])([a-f0-9])$/i;
          hex = hex.replace(shorthandRegex, function(m, r, g, b) {
            return r + r + g + g + b + b;
          });
          let result = /^#?([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})$/i.exec(hex);
          return result ? {
            r: parseInt(result[1], 16),
            g: parseInt(result[2], 16),
            b: parseInt(result[3], 16)
          } : null;
        }
        function rgbToHsl(r, g, b) {
          r /= 255; g /= 255; b /= 255;
          let max = Math.max(r, g, b), min = Math.min(r, g, b);
          let h, s, l = (max + min) / 2;
          l = Math.round(l * 100);
          if (max == min) {
            h = s = 0;
          } else {
            let d = max - min;
            s = l > 50 ? d / (2 - max - min) : d / (max + min);
            s = Math.round(s * 100);
            switch (max) {
              case r: h = (g - b) / d + (g < b ? 6 : 0); break;
              case g: h = (b - r) / d + 2; break;
              case b: h = (r - g) / d + 4; break;
            }
            h /= 6;
          }
          h = Math.round(h * 360);
          return [h, s, l];
        }
        function hslToRgb(h, s, l) {
          let r, g, b;
          h /= 360;
          s /= 100;
          l /= 100;
          if (s == 0) {
            r = g = b = l;
          } else {
            function hue2rgb(p, q, t) {
              if (t < 0) t += 1;
              if (t > 1) t -= 1;
              if (t < 1/6) return p + (q - p) * 6 * t;
              if (t < 1/2) return q;
              if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
              return p;
            }
            let q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            let p = 2 * l - q;
            r = hue2rgb(p, q, h + 1/3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1/3);
          }
          return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
        }
        function rgbToHex(r, g, b) {
          function componentToHex(c) {
            let hex = c.toString(16);
            return hex.length == 1 ? "0" + hex : hex;
          }
          return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
        }
        function correctTickCount(range, currentTickCount) {
            if (range <= 0) return 1;
            if (currentTickCount <= 0) currentTickCount = 1;
            if (range % currentTickCount === 0) {
                return currentTickCount;
            }
            let lowerCount = 1;
            for (let i = currentTickCount - 1; i > 1; i--) {
                if (range % i === 0) {
                    lowerCount = i;
                    break;
                }
            }
            let upperCount = range;
            for (let i = currentTickCount + 1; i < range; i++) {
                if (range % i === 0) {
                    upperCount = i;
                    break;
                }
            }
            if ((upperCount - currentTickCount) <= (currentTickCount - lowerCount)) {
                return upperCount;
            } else {
                return lowerCount;
            }
        }
        function buildChart(config) {
            const svgRoot = document.getElementById('combo-chart-svg');
            const defs = svgRoot.querySelector('defs');
            const style = svgRoot.querySelector('style');
            const existingGrads = defs.querySelectorAll('linearGradient');
            existingGrads.forEach(grad => grad.remove());
            let styleContent = style.textContent;
            styleContent = styleContent.replace(/[.]bar {[^}]*}/g, '');
            styleContent = styleContent.replace(/[.]line {[^}]*}/g, '');
            styleContent = styleContent.replace(/[.]marker {[^}]*}/g, '');
            styleContent = styleContent.replace(/[.]data-label-line {[^}]*}/g, '');
            style.textContent = styleContent.trim();
            const gradId = 'barGradient';
            let gradient = createSVGElement('linearGradient', {
                id: gradId, x1: "0", x2: "0", y1: "0", y2: "1"
             });
            gradient.appendChild(createSVGElement('stop', {
                offset: '0%', 'stop-color': config.colors.bar.start
            }));
            gradient.appendChild(createSVGElement('stop', {
                offset: '100%', 'stop-color': config.colors.bar.end
            }));
            defs.appendChild(gradient);
            const nl = String.fromCharCode(10);
            const stylesArray = [
                "",
                "            .bar { fill: url(#" + gradId + ");",
                "            filter: url(#shadow); }",
                "            .line { fill: none; stroke: " + config.colors.line + ";",
                "            stroke-width: 3; stroke-linejoin: round; filter: url(#shadow); }",
                "            .marker { fill: " + config.colors.marker + ";",
                "            stroke: white; stroke-width: 2; filter: url(#shadow); }",
                "            .data-label-line { font-family: 'Roboto', sans-serif;",
                "            font-size: 13px; font-weight: bold; fill: " + config.colors.lineLabel + "; text-anchor: middle; filter: url(#text-halo-white); }",
                 "        "
            ];
            let dynamicStyles = stylesArray.join(nl);
            style.textContent += dynamicStyles;
            const groups = {
                title: document.getElementById('title-group'),
                legend: document.getElementById('legend-group'),
                yAxes: document.getElementById('y-axes-group'),
                xAxis: document.getElementById('x-axis-group'),
                 barChart: document.getElementById('bar-chart-group'),
                lineChart: document.getElementById('line-chart-group'),
                source: document.getElementById('source-group')
            };
            for (const key in groups) {
                while (groups[key].firstChild) groups[key].removeChild(groups[key].firstChild);
            }
            const chartAreaWidth = config.layout.width - config.layout.marginLeft - config.layout.marginRight;
            const chartAreaHeight = config.layout.height - config.layout.marginTop - config.layout.marginBottom;
            const bottomY = config.layout.marginTop + chartAreaHeight;
            const maxBarValue = Math.max(...config.items.map(d => d.barValue), config.yAxisLeft.max);
            const minBarValue = Math.min(...config.items.map(d => d.barValue), config.yAxisLeft.min);
            const leftTickCount = config.yAxisLeft.tickCount > 0 ? config.yAxisLeft.tickCount : 1;
            const preliminaryLeftRange = Math.max(1, config.yAxisLeft.max - config.yAxisLeft.min);
            let leftTickInterval = preliminaryLeftRange / leftTickCount;
            let effectiveYAxisLeftMin, effectiveYAxisLeftMax;
            if (minBarValue < config.yAxisLeft.min) {
                const limit = Math.max(Math.abs(minBarValue), maxBarValue);
                if (leftTickInterval <= 0 || isNaN(leftTickInterval)) leftTickInterval = Math.max(1, limit / 5);
                effectiveYAxisLeftMax = Math.ceil(limit / leftTickInterval) * leftTickInterval;
                effectiveYAxisLeftMin = -effectiveYAxisLeftMax;
                if (config.yAxisLeft.min > effectiveYAxisLeftMin) effectiveYAxisLeftMin = config.yAxisLeft.min;
                if (config.yAxisLeft.max < effectiveYAxisLeftMax) effectiveYAxisLeftMax = config.yAxisLeft.max;
            } else {
                effectiveYAxisLeftMin = config.yAxisLeft.min;
                if (leftTickInterval <= 0 || isNaN(leftTickInterval)) leftTickInterval = Math.max(1, (maxBarValue - effectiveYAxisLeftMin) / 5) || 1;
                effectiveYAxisLeftMax = Math.max(config.yAxisLeft.max, Math.ceil(maxBarValue / leftTickInterval) * leftTickInterval);
            }
            if (effectiveYAxisLeftMax <= effectiveYAxisLeftMin) {
                effectiveYAxisLeftMax = effectiveYAxisLeftMin + leftTickInterval * leftTickCount;
                if (effectiveYAxisLeftMax <= effectiveYAxisLeftMin) effectiveYAxisLeftMax = effectiveYAxisLeftMin + 1;
            }
            if (effectiveYAxisLeftMax === effectiveYAxisLeftMin) effectiveYAxisLeftMax += leftTickInterval;
            const rightRangeOriginal = config.yAxisRight.max - config.yAxisRight.min;
            let rightTickCount = config.yAxisRight.tickCount > 0 ? config.yAxisRight.tickCount : 1;
            if (rightRangeOriginal > 0) {
                rightTickCount = correctTickCount(rightRangeOriginal, rightTickCount);
            }
            const rightTickInterval = (rightRangeOriginal > 0 && rightTickCount > 0) ?
                (rightRangeOriginal) / rightTickCount : 1;
            const maxLineValue = Math.max(...config.items.map(d => d.lineValue), config.yAxisRight.max);
            const minLineValue = Math.min(...config.items.map(d => d.lineValue), config.yAxisRight.min);
            let effectiveYAxisRightMin = config.yAxisRight.min;
            if (minLineValue < config.yAxisRight.min) {
                effectiveYAxisRightMin = Math.floor(minLineValue / rightTickInterval) * rightTickInterval;
            }
            let effectiveYAxisRightMax = config.yAxisRight.max;
            if (maxLineValue > config.yAxisRight.max) {
                effectiveYAxisRightMax = Math.ceil(maxLineValue / rightTickInterval) * rightTickInterval;
            }
            if (effectiveYAxisRightMax <= effectiveYAxisRightMin) effectiveYAxisRightMax = effectiveYAxisRightMin + rightTickInterval * (rightTickCount || 1);
            const leftYRange = effectiveYAxisLeftMax - effectiveYAxisLeftMin;
            const rightYRange = effectiveYAxisRightMax - effectiveYAxisRightMin;
            const xScale = (index) => config.layout.marginLeft + (chartAreaWidth / config.items.length) * (index + 0.5);
            const yScaleLeft = (value) => bottomY - (leftYRange > 0 ? chartAreaHeight * ((value - effectiveYAxisLeftMin) / leftYRange) : chartAreaHeight / 2);
            const yScaleRight = (value) => bottomY - (rightYRange > 0 ? chartAreaHeight * ((value - effectiveYAxisRightMin) / rightYRange) : chartAreaHeight / 2);
        const animProgress = (config.animation == null) 
            ? 1.0 
            : Math.max(0, Math.min(1, Number(config.animation)));
        let animOpacity;
        const fadeStart = 0.7;
        if (animProgress <= fadeStart) {
            animOpacity = 0;
        } else {
            animOpacity = (animProgress - fadeStart) / (1.0 - fadeStart);
        }
            groups.title.appendChild(createSVGElement('text', { x: 312, y: 65, class: 'title' }, config.title));
            groups.title.appendChild(createSVGElement('text', { x: 312, y: 90, class: 'subtitle' }, config.subtitle));
            groups.source.appendChild(createSVGElement('text', { x: 605, y: 525, class: 'source' }, config.source));
            const legendY = config.layout.marginTop - 60;
            const legendAGroup = createSVGElement('g', { transform: 'translate(' + config.layout.marginLeft + ', ' + legendY + ')' });
            legendAGroup.appendChild(createSVGElement('rect', { x: 0, y: -6, width: 12, height: 12, rx: 2, class: 'bar' }));
            legendAGroup.appendChild(createSVGElement('text', { x: 18, y: 0, class: 'legend-item', 'dominant-baseline': 'middle' }, config.legendBarLabel));
            groups.legend.appendChild(legendAGroup);
            const legendBGroup = createSVGElement('g', { transform: 'translate(' + (config.layout.width - config.layout.marginRight) + ', ' + legendY + ')' });
            legendBGroup.appendChild(createSVGElement('line', { x1: 0, y1: 0, x2: -12, y2: 0, stroke: config.colors.line, 'stroke-width': 2 }));
            legendBGroup.appendChild(createSVGElement('circle', { cx: -6, cy: 0, r: 3, fill: config.colors.marker, stroke: 'white', 'stroke-width': 1 }));
            legendBGroup.appendChild(createSVGElement('text', { x: -20, y: 0, class: 'legend-item', 'dominant-baseline': 'middle', 'text-anchor': 'end' }, config.legendLineLabel));
            groups.legend.appendChild(legendBGroup);
            for (let i = 0; i <= leftTickCount; i++) {
                const tickValue = effectiveYAxisLeftMin + (leftYRange / leftTickCount) * i;
                const y = yScaleLeft(tickValue);
                if (i > 0 && i <= leftTickCount) {
                    groups.yAxes.appendChild(createSVGElement('line', { x1: config.layout.marginLeft, y1: y, x2: config.layout.width - config.layout.marginRight, y2: y, class: 'grid-line' }));
                }
                groups.yAxes.appendChild(createSVGElement('text', { x: config.layout.marginLeft - 10, y, class: 'y-axis-value', 'text-anchor': 'end', 'dominant-baseline': 'middle' }, Math.round(tickValue)));
            }
            groups.yAxes.appendChild(createSVGElement('text', { x: 30, y: config.layout.marginTop - 20, class: 'y-axis-value' }, config.yAxisLeftLabel));
            for (let i = 0; i <= rightTickCount; i++) {
                const tickValue = effectiveYAxisRightMin + (rightYRange / rightTickCount) * i;
                const y = yScaleRight(tickValue);
                groups.yAxes.appendChild(createSVGElement('text', { x: config.layout.width - config.layout.marginRight + 10, y: y, class: 'y-axis-value', 'text-anchor': 'start', 'dominant-baseline': 'middle' }, Math.round(tickValue) + (config.yAxisRight.unit || '')));
            }
            groups.yAxes.appendChild(createSVGElement('text', { x: 585, y: config.layout.marginTop - 20, class: 'y-axis-value', 'text-anchor': 'end' }, config.yAxisRightLabel));
            const slotWidth = chartAreaWidth / config.items.length;
            config.items.forEach((item, index) => {
                const labelContent = truncateText(item.label, slotWidth * 0.9, 'axis-label');
                groups.xAxis.appendChild(createSVGElement('text', { x: xScale(index), y: bottomY + 20, class: 'axis-label' }, labelContent));
            });
            groups.xAxis.appendChild(createSVGElement('line', { x1: config.layout.marginLeft, y1: bottomY, x2: config.layout.width - config.layout.marginRight, y2: bottomY, stroke: '#bdbdbd' }));
            if (effectiveYAxisLeftMin < 0 && effectiveYAxisLeftMax > 0) {
                const yZero = yScaleLeft(0);
                groups.yAxes.appendChild(createSVGElement('line', { x1: config.layout.marginLeft, y1: yZero, x2: config.layout.width - config.layout.marginRight, y2: yZero, stroke: '#bdbdbd', 'stroke-width': 1 }));
            }
            const barWidth = slotWidth * config.barOptions.barToSlotRatio;
            const barLabelYPositions = [];
            const yZero = yScaleLeft(Math.max(0, effectiveYAxisLeftMin));
            config.items.forEach((item, index) => {
                const x = xScale(index) - barWidth / 2;
                let finalBarY, finalBarHeight;
                let animatedBarY, animatedBarHeight;
                if (item.barValue >= 0) {
                    finalBarY = yScaleLeft(item.barValue);
                    finalBarHeight = Math.max(0, yZero - finalBarY);
                    animatedBarHeight = finalBarHeight * animProgress;
                    animatedBarY = yZero - animatedBarHeight;
                } else {
                    finalBarY = yZero;
                    finalBarHeight = Math.max(0, yScaleLeft(item.barValue) - yZero);
                    animatedBarHeight = finalBarHeight * animProgress;
                    animatedBarY = yZero;
                }
                if (animatedBarHeight >= 1) {
                    groups.barChart.appendChild(createSVGElement('rect', { 
                        x: x, 
                        y: animatedBarY,
                        width: barWidth, 
                        height: animatedBarHeight,
                        class: 'bar', 
                        rx: 4, 
                        ry: 4 
                    }));
                }
                const isOutside = animatedBarHeight < 30;
                let labelY, labelClass;
                const labelMarginInside = 20;
                const labelMarginOutside = 10;
                if (isOutside) {
                    labelClass = 'data-label-bar-outside';
                    labelY = (item.barValue >= 0) ?
                        animatedBarY - labelMarginOutside : animatedBarY + animatedBarHeight + labelMarginOutside + 8;
                } else {
                    labelClass = 'data-label-bar-inside';
                    labelY = (item.barValue >= 0) ?
                        animatedBarY + labelMarginInside : animatedBarY + animatedBarHeight - labelMarginInside;
                }
                if (!isOutside && item.barValue >= 0 && labelY > yZero - 5) labelY = yZero - 5;
                if (!isOutside && item.barValue < 0 && labelY < yZero + 15) labelY = yZero + 15;
                barLabelYPositions.push({ index: index, y: labelY });
                groups.barChart.appendChild(createSVGElement('text', { 
                    x: xScale(index), 
                    y: labelY,
                    class: labelClass,
                    style: 'opacity: ' + animOpacity
                }, item.barValue));
            });
            const yZeroLine = yScaleRight(Math.max(effectiveYAxisRightMin, 0));
            let allLineValuesSame = false;
            if (config.items.length > 0) {
                const firstValue = config.items[0].lineValue;
                allLineValuesSame = config.items.every(item => item.lineValue === firstValue);
            }
            const pointsData = config.items.map((item, index) => {
                const x = xScale(index);
                const finalY = yScaleRight(item.lineValue);
                let animatedY = yZeroLine + (finalY - yZeroLine) * animProgress;
                if (allLineValuesSame) {
                    const jiggle = (index % 2 === 0) ? 0.01 : -0.01;
                    animatedY += jiggle;
                }
                return {
                    x: x,
                    finalY: finalY,
                    animatedY: animatedY,
                    value: item.lineValue
                };
            });
            const points = pointsData.map(d => d.x + ',' + d.animatedY).join(' ');
groups.lineChart.appendChild(createSVGElement('polyline', { class: 'line', points }));
            pointsData.forEach((data, index) => {
                const x = data.x;
                const y = data.animatedY;
                groups.lineChart.appendChild(createSVGElement('circle', { 
                    cx: x, 
                    cy: y,
                    r: config.lineOptions.markerRadius, 
                    class: 'marker' 
                }));
                let lineLabelY = y - 20;
const correspondingBarLabel = barLabelYPositions.find(pos => pos.index === index);
                if (correspondingBarLabel && Math.abs(lineLabelY - correspondingBarLabel.y) < 20) {
                    let belowY = y + 25;
if (belowY < bottomY - 5) { 
                         lineLabelY = belowY;
                    }
                }
                groups.lineChart.appendChild(createSVGElement('text', { 
                    x: x, 
                    y: lineLabelY,
                    class: 'data-label-line',
                    style: 'opacity: ' + animOpacity
                }, data.value + (config.yAxisRight.unit || '')));
            });
}
        try {
            const configJson = document.getElementById('chart-json-data').textContent;
            if (!configJson || configJson.trim() === '') {
                throw new Error("JSON data is empty.");
            }
            let fullConfig;
            try {
                fullConfig = JSON.parse(configJson);
            } catch (parseError) {
                throw new Error('JSON parsing error: ' + parseError.message);
            }
            if (!fullConfig || typeof fullConfig !== 'object' || !fullConfig.data) {
                throw new Error("Invalid JSON structure: Missing 'data' key.");
            }
            const config = fullConfig.data;
            config.layout = { 
                "width": 600, 
                "height": 450, 
                "marginTop": 180, 
                "marginBottom": 0, 
                "marginLeft": 70, 
                "marginRight": 70 
            };
            config.barOptions = config.barOptions || {};
            config.barOptions.barToSlotRatio = (config.barOptions.barToSlotRatio == null) ? 0.55 : config.barOptions.barToSlotRatio;
            config.barOptions.labelPosition = config.barOptions.labelPosition || "auto";
            config.lineOptions = config.lineOptions || {};
            config.lineOptions.markerRadius = (config.lineOptions.markerRadius == null) ? 5 : config.lineOptions.markerRadius;
            if (config.colors && typeof config.colors === 'string') {
                const colorKeyword = config.colors.toLowerCase();
                let generatedColorObject;
                if (colorKeyword === '#gemini') {
                    generatedColorObject = {
                        bar: { start: "#E68A9C", end: "#9F63D0" },
                        line: "#616AD8",
                        marker: "#616AD8",
                        lineLabel: "#4D46AE"
                    };
                } else if (colorKeyword === '#night') {
                    var NIGHT_COLORS = [
                        { start: "#8a2be2", end: "#4169e1" },
                        { start: '#8f83d6', end: '#5a5188' },
                        { start: '#7891dd', end: '#4456a0' },
                        { start: '#FF8F00', end: '#FF8F00' },
                        { start: '#54c5d5', end: '#267b87' }
                    ];
                    generatedColorObject = {
                        bar: { start: NIGHT_COLORS[0].start, end: NIGHT_COLORS[0].end },
                        line: NIGHT_COLORS[3].start,
                        marker: NIGHT_COLORS[3].start,
                        lineLabel: NIGHT_COLORS[3].end
                    };
              } else {
                    const inputColorHex = config.colors;
                    const rgb = hexToRgb(inputColorHex);
                    if (!rgb) {
                        throw new Error("Invalid HEX color string in 'colors': " + inputColorHex);
                    }
                    let [h_orig, s_orig, l_orig] = rgbToHsl(rgb.r, rgb.g, rgb.b);
                    const MIN_L = 25;
                    const MAX_L = 75;
                    const ACHROMATIC_THRESHOLD_S = 5;
                    const MIN_S_FOR_COLOR = 30;
                    let h_base = h_orig;
                    let s_base;
                    let l_base = Math.max(MIN_L, Math.min(MAX_L, l_orig));
                    if (s_orig < ACHROMATIC_THRESHOLD_S) {
                        s_base = 0;
                    } else {
                        s_base = Math.max(s_orig, MIN_S_FOR_COLOR);
                    }
                    const hex_base_color = rgbToHex(...hslToRgb(h_base, s_base, l_base));
                    const hex_bar_end = hex_base_color;
                    let l_bar_start = Math.min(100, l_base + 15);
                    let s_bar_start = (l_base > MAX_L - 10) ? s_base * 0.8 : s_base;
                    const hex_bar_start = rgbToHex(...hslToRgb(h_base, s_bar_start, l_bar_start));
                    const hex_line_main = hex_base_color;
                    let l_line_label = Math.max(0, l_base - 25);
                    let s_line_label = (l_base < MIN_L + 10) ? Math.min(100, s_base * 1.2) : s_base;
                    const hex_line_label = rgbToHex(...hslToRgb(h_base, s_line_label, l_line_label));
                    generatedColorObject = {
                        bar: { start: hex_bar_start, end: hex_bar_end },
                        line: hex_line_main,
                        marker: hex_line_main,
                        lineLabel: hex_line_label
                    };
                }
                config.colors = generatedColorObject;
            } else if (!config.colors || typeof config.colors !== 'object' || !config.colors.bar || !config.colors.line) {
                 throw new Error("'colors' must be a valid color object {bar, line, ...} or a single color string.");
            }
             if (config.colors && typeof config.colors === 'object') {
                 if (!config.colors.marker) config.colors.marker = config.colors.line;
                 if (!config.colors.lineLabel) {
                     const markerRgb = hexToRgb(config.colors.marker);
                     if (markerRgb) {
                         const markerHsl = rgbToHsl(markerRgb.r, markerRgb.g, markerRgb.b);
                         const label_l = Math.max(10, markerHsl[2] - 25);
                         const labelRgb = hslToRgb(markerHsl[0], markerHsl[1], label_l);
                         config.colors.lineLabel = rgbToHex(labelRgb[0], labelRgb[1], labelRgb[2]);
                     } else {
                        config.colors.lineLabel = config.colors.line;
                     }
                 }
            }
            if (!config || typeof config !== 'object' || !config.items || !Array.isArray(config.items)) {
                throw new Error("Invalid JSON structure: Missing required fields like 'items' inside 'data'.");
            }
            buildChart(config);
        } catch (e) {
            console.error("Chart Error:", e);
            const svgRoot = document.getElementById('combo-chart-svg');
            const existingError = svgRoot.querySelector('.error-message');
            if (existingError) existingError.remove();
            const textEl = createSVGElement('text', {
                x: svgRoot.viewBox.baseVal.width / 2 + svgRoot.viewBox.baseVal.x,
                y: svgRoot.viewBox.baseVal.height / 2 + svgRoot.viewBox.baseVal.y,
                'text-anchor': 'middle',
                fill: 'red',
                'font-family': "'Noto Sans JP', sans-serif",
                class: 'error-message'
            });
            textEl.textContent = e.message.includes('JSON') ? 'JSONデータにエラーがあります。' : 'グラフ描画エラーが発生しました。';
            if (e.message.includes('JSON parsing error')) {
                textEl.textContent += ' (' + e.message.split(': ')[1] + ')';
            }
            svgRoot.appendChild(textEl);
        }
    </script>
</svg>
`;
const MULTI_LINE_CHART_TEMPLATE = `
<svg width="650" height="510" viewBox="25 15 620 485" xmlns="http://www.w3.org/2000/svg" font-family="'Noto Sans JP', sans-serif" id="multi-line-chart-svg">
  <script id="chart-json-data" type="application/json">
    {
      "chartType": "multi-line",
      "data": {
        "title": "データ系列の比較",
        "subtitle": "（年間サンプルデータ）",
        "source": "出典：サンプルデータ",
        "yAxisUnitLabel": "（単位）",
        "xAxisLabels": ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
        "series": [
          { "id": "A", "label": "系列A", "values": [10, 12, 15, 20, 24, 27, 29, 28, 25, 20, 16, 12] },
          { "id": "B", "label": "系列B", "values": [7, 8, 11, 15, 18, 20, 21, 20, 16, 13, 10, 8] },
          { "id": "C", "label": "系列C", "values": [4, 5, 6, 8, 10, 11, 12, 10, 8, 7, 5, 4] }
        ],
"colors": 
[
          { "id": "A", "start": "#e68a9c", "end": "#d96d8f" },
          { "id": "B", "start": "#b469b8", "end": "#a656ad" },
          { "id": "C", "start": "#7c6ce8", "end": "#6b5ce0" }
        ],
        "layout": {
          "width": 650,
          "height": 510,
"marginTop": 140,
          "marginBottom": 80,
          "marginLeft": 75,
          "marginRight": 35,
          "horizontalPadding": 20
        },
        "yAxis": {
          "min": 0,
          "max": 30,
          "tickCount": 6
   },
        "lineOptions": {
          "markerRadius": 5,
          "dataLabelOffsetY": 8
        },
        "animation": 1
      }
    }
  </script>
  <defs>
    <filter id="shadow"><feDropShadow dx="1" dy="2" stdDeviation="2" flood-color="#000" flood-opacity="0.2"/></filter>
    <filter id="text-halo-white" x="-50%" y="-50%" width="200%" height="200%">
      <feDropShadow dx="0" dy="0" stdDeviation="3" flood-color="white" flood-opacity="0.9" in="SourceAlpha" result="shadow"/>
      <feMerge>
       <feMergeNode in="shadow"/>
        <feMergeNode in="SourceGraphic"/>
      </feMerge>
    </filter>
  </defs>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500&amp;family=Roboto:wght@400&amp;display=swap');
    .title { font-size: 24px; font-weight: 500; text-anchor: middle;
fill: #5f6368; }
    .subtitle { font-size: 16px; text-anchor: middle; fill: #5f6368;
}
    .source { font-size: 11px; fill: #70757a; text-anchor: end; }
    .legend-item { font-size: 13px;
fill: #202124; dominant-baseline: middle; }
    .grid-line { stroke: #e0e0e0; stroke-dasharray: 2 2; stroke-width: 0.8;
}
    .zero-line { stroke: #bdbdbd; stroke-width: 1; }
    .axis-label { font-size: 13px; fill: #5f6368;
text-anchor: middle; }
    .y-axis-unit-label { font-family: 'Noto Sans JP', sans-serif; font-size: 12px; fill: #757575; text-anchor: start;
}
    .y-axis-value { font-family: 'Roboto', sans-serif; font-size: 12px;
fill: #757575; text-anchor: end; }
    .data-label { font-family: 'Roboto', sans-serif; font-size: 12px; font-weight: normal; text-anchor: middle;
filter: url(#text-halo-white); }
  </style>
  <g id="title-group"></g>
  <g id="legend-group"></g>
  <g id="y-axis-group"></g>
  <g id="x-axis-group"></g>
  <g id="chart-area-group"></g>
  <g id="source-group"></g>
  <script type="text/javascript">
    const svgNS = "http://www.w3.org/2000/svg";
function createSVGElement(name, attributes, textContent) {
      const el = document.createElementNS(svgNS, name);
for (const key in attributes) el.setAttribute(key, attributes[key]);
      if (textContent !== undefined && textContent !== null) el.textContent = textContent;
      return el;
}
    function hexToRgb(hex) {
      var shorthandRegex = /^#?([a-f0-9])([a-f0-9])([a-f0-9])$/i;
hex = hex.replace(shorthandRegex, function(m, r, g, b) {
        return r + r + g + g + b + b;
      });
var result = /^#?([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})$/i.exec(hex);
      return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
      } : null;
}
    function rgbToHsl(r, g, b) {
      r /= 255;
g /= 255; b /= 255;
      var max = Math.max(r, g, b), min = Math.min(r, g, b);
var h, s, l = (max + min) / 2;
      l = Math.round(l * 100);
      if (max == min) {
        h = s = 0;
      } else {
        var d = max - min;
s = l > 50 ? d / (2 - max - min) : d / (max + min);
s = Math.round(s * 100);
        switch (max) {
          case r: h = (g - b) / d + (g < b ? 6 : 0);
break;
          case g: h = (b - r) / d + 2; break;
case b: h = (r - g) / d + 4; break;
}
        h /= 6;
}
      h = Math.round(h * 360);
      return [h, s, l];
}
    function hslToRgb(h, s, l) {
      var r, g, b;
h /= 360;
      s /= 100;
      l /= 100;
      if (s == 0) {
        r = g = b = l;
      } else {
        function hue2rgb(p, q, t) {
          if (t < 0) t += 1;
if (t > 1) t -= 1;
          if (t < 1/6) return p + (q - p) * 6 * t;
if (t < 1/2) return q;
          if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
return p;
        }
        var q = l < 0.5 ?
l * (1 + s) : l + s - l * s;
var p = 2 * l - q;
        r = hue2rgb(p, q, h + 1/3);
        g = hue2rgb(p, q, h);
b = hue2rgb(p, q, h - 1/3);
      }
      return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
}
    function rgbToHex(r, g, b) {
      function componentToHex(c) {
        var hex = c.toString(16);
return hex.length == 1 ? "0" + hex : hex;
}
      return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
}
    function linearInterpolateColor(hex1, hex2, t) {
      var rgb1 = hexToRgb(hex1);
var rgb2 = hexToRgb(hex2);
      var r = Math.round(rgb1.r + (rgb2.r - rgb1.r) * t);
var g = Math.round(rgb1.g + (rgb2.g - rgb1.g) * t);
      var b = Math.round(rgb1.b + (rgb2.b - rgb1.b) * t);
return rgbToHex(r, g, b);
    }
    function generateGradientScale(gradientStops, numSteps) {
      var palette = [];
if (numSteps <= 0) return [];
      if (numSteps === 1) return [gradientStops[0]];
      var numSegments = gradientStops.length - 1;
for (var i = 0; i < numSteps; i++) {
        var t_global = (numSteps === 1) ?
0 : (i / (numSteps - 1));
        var segmentProgress = t_global * numSegments;
var stopIndex1 = Math.floor(segmentProgress);
        var stopIndex2 = Math.min(stopIndex1 + 1, numSegments);
        var t_local = segmentProgress - stopIndex1;
var hex1 = gradientStops[stopIndex1];
        var hex2 = gradientStops[stopIndex2];
        palette.push(linearInterpolateColor(hex1, hex2, t_local));
      }
      return palette;
}
    function generateMonochromaticScale(primaryHex, colorIds) {
        var numSteps = colorIds.length;
        if (numSteps === 0) return [];
        var rgb = hexToRgb(primaryHex);
        if (!rgb) return [];
        var hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
        var h = hsl[0], s = hsl[1], l_mid_original = hsl[2];
        var gradientDarken = 7;
        var L_MAX = 80; 
        var L_MIN = 45;
        var l_steps_start;
        if (numSteps === 1) {
            l_steps_start = [L_MIN];
        } else {
            var l_step_amount = (L_MAX - L_MIN) / (numSteps - 1);
            var l_steps_calc = [];
            for (var i = 0; i < numSteps; i++) {
                var l_val = L_MAX - (i * l_step_amount);
                l_steps_calc.push(l_val);
            }
            var l_steps_light = [];
            var l_steps_dark = [];
            var mid_index = Math.ceil(numSteps / 2);
            for (var i = 0; i < mid_index; i++) {
                l_steps_light.push(l_steps_calc[i]);
            }
            for (var i = mid_index; i < numSteps; i++) {
                l_steps_dark.push(l_steps_calc[i]);
            }
            l_steps_start = [];
            for (var i = 0; i < numSteps; i++) {
                if (i % 2 === 0) {
                    l_steps_start.push(l_steps_light[i / 2]);
                } else {
                    l_steps_start.push(l_steps_dark[(i - 1) / 2]);
                }
            }
        }
        var colors = [];
        for (var i = 0; i < numSteps; i++) {
            var l_start_step = l_steps_start[i];
            var l_start = Math.max(0, Math.min(100, l_start_step));
            var l_end = Math.max(0, Math.min(100, l_start - gradientDarken));
            var s_adjusted = s;
            if (l_start > 85) s_adjusted = s * (1 - (l_start - 85) / 15);
            if (l_start < 20) s_adjusted = s * (l_start / 20);
            s_adjusted = Math.max(0, Math.min(100, s_adjusted));
            var rgb_start = hslToRgb(h, s_adjusted, l_start);
            var hex_start = rgbToHex(rgb_start[0], rgb_start[1], rgb_start[2]);
            var rgb_end = hslToRgb(h, s_adjusted, l_end);
            var hex_end = rgbToHex(rgb_end[0], rgb_end[1], rgb_end[2]);
            colors.push({
                id: colorIds[i],
                start: hex_start,
                end: hex_end
            });
        }
        return colors;
    }
    function buildChart(config) {
      const svgRoot = document.getElementById('multi-line-chart-svg');
const defs = svgRoot.querySelector('defs');
      const style = svgRoot.querySelector('style');
      const existingGrads = defs.querySelectorAll('linearGradient');
existingGrads.forEach(function(grad) { grad.remove(); });
      let styleContent = style.textContent;
const dynamicRulePattern = /\.((line|marker|area|data-label)-[a-z])\s*{[^}]*}/g;
      styleContent = styleContent.replace(dynamicRulePattern, '');
      style.textContent = styleContent.trim();
      const nl = String.fromCharCode(10);
      let dynamicStyleParts = [];
      config.colors.forEach(function(color) {
        const id = color.id;
        const markerGradId = 'markerGradient' + id;
        const areaGradId = 'areaGradient' + id;
        const cssPrefix 
= id.toLowerCase();
        let markerGrad = createSVGElement('linearGradient', { id: markerGradId, x1: "0", x2: "0", y1: "0", y2: "1" });
        markerGrad.appendChild(createSVGElement('stop', { offset: '0%', 'stop-color': color.start }));
        markerGrad.appendChild(createSVGElement('stop', { offset: '100%', 'stop-color': color.end }));
        defs.appendChild(markerGrad);
        let areaGrad = createSVGElement('linearGradient', { id: areaGradId, x1: "0", x2: "0", y1: "0", y2: "1" });
      areaGrad.appendChild(createSVGElement('stop', { offset: '0%', 'stop-color': color.start, 'stop-opacity': '0.2' }));
        areaGrad.appendChild(createSVGElement('stop', { offset: '100%', 'stop-color': color.end, 'stop-opacity': '0' }));
defs.appendChild(areaGrad);
        dynamicStyleParts.push(".line-" + cssPrefix + " { fill: none; stroke: " + color.end + "; stroke-width: 3; stroke-linejoin: round; stroke-linecap: round; filter: url(#shadow); }");
dynamicStyleParts.push(".marker-" + cssPrefix + " { fill: url(#" + markerGradId + "); stroke: white; stroke-width: 2; filter: url(#shadow); }");
dynamicStyleParts.push(".area-" + cssPrefix + " { fill: url(#" + areaGradId + "); }");
dynamicStyleParts.push(".data-label-" + cssPrefix + " { fill: " + color.end + "; }");
      });
      const dynamicStyles = dynamicStyleParts.join(nl);
      style.textContent += nl + dynamicStyles + nl;
      const groups = {
        title: document.getElementById('title-group'),
        legend: document.getElementById('legend-group'),
        yAxis: document.getElementById('y-axis-group'),
        xAxis: document.getElementById('x-axis-group'),
        chartArea: document.getElementById('chart-area-group'),
        source: document.getElementById('source-group')
      };
      for (const key in groups) {
        while (groups[key].firstChild) groups[key].removeChild(groups[key].firstChild);
}
      const allValues = config.series.reduce(function(acc, s) { return acc.concat(s.values); }, []);
const dataMin = Math.min.apply(null, allValues.concat([config.yAxis.min]));
      const dataMax = Math.max.apply(null, allValues.concat([config.yAxis.max]));
      let effectiveYAxisMin, effectiveYAxisMax;
const tickCount = config.yAxis.tickCount > 0 ? config.yAxis.tickCount : 1;
      const preliminaryRange = Math.max(1, config.yAxis.max - config.yAxis.min);
let tickInterval = preliminaryRange / tickCount;
      if (dataMin < 0) {
        const absLimit = Math.max(Math.abs(dataMin), dataMax);
if (tickInterval <= 0 || isNaN(tickInterval)) tickInterval = Math.max(1, absLimit / 5);
        effectiveYAxisMax = Math.ceil(absLimit / tickInterval) * tickInterval;
effectiveYAxisMin = -effectiveYAxisMax;
        if(config.yAxis.min > effectiveYAxisMin) effectiveYAxisMin = config.yAxis.min;
        if(config.yAxis.max < effectiveYAxisMax) effectiveYAxisMax = config.yAxis.max;
} else {
        effectiveYAxisMin = config.yAxis.min;
if (tickInterval <= 0 || isNaN(tickInterval)) tickInterval = Math.max(1, (dataMax - effectiveYAxisMin) / 5) || 1;
effectiveYAxisMax = Math.max(config.yAxis.max, Math.ceil(dataMax / tickInterval) * tickInterval);
      }
      if (effectiveYAxisMax <= effectiveYAxisMin) {
        effectiveYAxisMax = effectiveYAxisMin + tickInterval * tickCount;
if (effectiveYAxisMax <= effectiveYAxisMin) effectiveYAxisMax = effectiveYAxisMin + 1;
      }
      var yAxisRange = effectiveYAxisMax - effectiveYAxisMin;
      const tempYLabel = createSVGElement('text', { class: 'y-axis-value', style: 'visibility: hidden;' }, Math.round(effectiveYAxisMax));
svgRoot.appendChild(tempYLabel);
      const yAxisLabelWidth = tempYLabel.getComputedTextLength();
      svgRoot.removeChild(tempYLabel);
      const dynamicMarginLeft = Math.max(config.layout.marginLeft, yAxisLabelWidth + 15);
const chartAreaWidth = config.layout.width - dynamicMarginLeft - config.layout.marginRight;
      const chartAreaHeight = config.layout.height - config.layout.marginTop - config.layout.marginBottom;
const bottomY = config.layout.marginTop + chartAreaHeight;
      const horizontalPadding = config.layout.horizontalPadding ||
0;
      const effectivePlotWidth = chartAreaWidth - (horizontalPadding * 2);
      const numPoints = config.xAxisLabels.length;
      const xScale = function(index) { return (dynamicMarginLeft + horizontalPadding) + (numPoints > 1 ? (effectivePlotWidth * (index / (numPoints - 1))) : effectivePlotWidth / 2);
};
      const yScale = function(value) { return bottomY - ((value - effectiveYAxisMin) / yAxisRange) * chartAreaHeight; };
      const zeroLineY = yScale(0);
      var animProgress = (config.animation == null) 
          ? 1.0 
          : Math.max(0, Math.min(1, Number(config.animation)));
      var animOpacity;
      var fadeStart = 0.7;
      if (animProgress <= fadeStart) {
          animOpacity = 0;
      } else {
          animOpacity = (animProgress - fadeStart) / (1.0 - fadeStart);
      }
      groups.title.appendChild(createSVGElement('text', { x: 325, y: 40, class: 'title' }, config.title));
groups.title.appendChild(createSVGElement('text', { x: 325, y: 65, class: 'subtitle' }, config.subtitle));
groups.source.appendChild(createSVGElement('text', { x: config.layout.width - config.layout.marginRight, y: bottomY + 60, class: 'source' }, config.source));
    const legendY = 95;
    const numLegends = config.series.length;
    const rectSize = 12;
    const rectTextGap = 6;
    const legendFontSize = 13;
    const fontSizeCharWidthRatio = 1.0; 
    let maxCharCount = 0;
    config.series.forEach(function(series) {
        const currentLabelLength = (series.label != null) ? series.label.length : 0;
        if (currentLabelLength > maxCharCount) {
            maxCharCount = currentLabelLength;
        }
    });
    const estimatedMaxTextWidth = maxCharCount * legendFontSize * fontSizeCharWidthRatio;
    const legendItemEndPadding = 25;
    const idealItemSlotWidth = rectSize + rectTextGap + estimatedMaxTextWidth + legendItemEndPadding;
    const totalIdealWidth = idealItemSlotWidth * numLegends;
    const availableLegendAreaLeft = dynamicMarginLeft;
    const availableLegendAreaRight = config.layout.width - config.layout.marginRight;
    const availableLegendAreaWidth = availableLegendAreaRight - availableLegendAreaLeft;
    let finalItemSlotWidth = 0;
    let legendBlockStartX = 0;
    if (totalIdealWidth <= availableLegendAreaWidth) {
        finalItemSlotWidth = idealItemSlotWidth;
        const svgCenterX = svgRoot.viewBox.baseVal.x + (svgRoot.viewBox.baseVal.width / 2);
        legendBlockStartX = svgCenterX - (totalIdealWidth / 2);
    } else {
        const constrainedItemSlotWidth = availableLegendAreaWidth / numLegends;
        finalItemSlotWidth = constrainedItemSlotWidth;
        legendBlockStartX = availableLegendAreaLeft;
    }
    config.series.forEach(function(series, i) {
        const columnStartX = legendBlockStartX + (i * finalItemSlotWidth);
        const itemGroup = createSVGElement('g');
        var colorEntry = config.colors.find(function(c) { return c.id === series.id; });
        const color = (colorEntry ? colorEntry.end : undefined) || '#000';
        itemGroup.appendChild(createSVGElement('rect', { x: columnStartX, y: legendY, width: rectSize, height: rectSize, rx: 2, fill: color, filter: 'url(#shadow)' }));
        let labelContent = series.label;
        const textStartX = columnStartX + rectSize + rectTextGap;
        const availableTextWidth = finalItemSlotWidth - (rectSize + rectTextGap) - (legendItemEndPadding / 2); 
        if (estimatedMaxTextWidth > availableTextWidth) {
            const tempText = createSVGElement('text', { class: 'legend-item', style: 'visibility: hidden;' }, labelContent);
            svgRoot.appendChild(tempText);
            if (tempText.getComputedTextLength() > availableTextWidth) {
                labelContent = labelContent.substring(0, Math.max(1, Math.floor(labelContent.length * availableTextWidth / tempText.getComputedTextLength()))) + '..';
            }
            svgRoot.removeChild(tempText);
        }
        itemGroup.appendChild(createSVGElement('text', { x: textStartX, y: legendY + rectSize / 2, class: 'legend-item' }, labelContent));
        groups.legend.appendChild(itemGroup);
    });
      groups.yAxis.appendChild(createSVGElement('text', {
        x: 30,
        y: config.layout.marginTop - 15,
        class: 'y-axis-unit-label'
      }, config.yAxisUnitLabel));
      const finalTickStep = yAxisRange / tickCount;
for (let i = 0; i <= tickCount; i++) {
        const tickValue = effectiveYAxisMin + (finalTickStep * i);
const y = yScale(tickValue);
        groups.yAxis.appendChild(createSVGElement('text', { x: dynamicMarginLeft - 10, y, class: 'y-axis-value', 'dominant-baseline': 'middle' }, Math.round(tickValue)));
if (Math.abs(tickValue) > 1e-6 || effectiveYAxisMin >= 0) {
          groups.yAxis.appendChild(createSVGElement('line', { x1: dynamicMarginLeft, y1: y, x2: config.layout.width - config.layout.marginRight, y2: y, class: 'grid-line' }));
}
      }
      if (effectiveYAxisMin < 0 && effectiveYAxisMax > 0) {
        groups.xAxis.appendChild(createSVGElement('line', { x1: dynamicMarginLeft, y1: zeroLineY, x2: config.layout.width - config.layout.marginRight, y2: zeroLineY, class: 'zero-line' }));
}
      groups.xAxis.appendChild(createSVGElement('line', { x1: dynamicMarginLeft, y1: bottomY, x2: config.layout.width - config.layout.marginRight, y2: bottomY, class: 'zero-line' }));
    const labelSlotWidth = effectivePlotWidth / (numPoints > 1 ? numPoints - 1 : 1);
    const rotationThreshold = 40; 
    const needsRotation = (labelSlotWidth < rotationThreshold && numPoints > 1);
    const labelYOffset = needsRotation ? bottomY + 25 : bottomY + 30;
    config.xAxisLabels.forEach(function(label, i) {
      let labelContent = label;
      const x = xScale(i); 
      const tempText = createSVGElement('text', { class: 'axis-label', style: 'visibility: hidden;' });
      svgRoot.appendChild(tempText);
      tempText.textContent = label;
      if (numPoints > 1 && tempText.getComputedTextLength() > labelSlotWidth * 0.9 && !needsRotation) {
        labelContent = label.substring(0, Math.max(1, Math.floor(label.length * labelSlotWidth * 0.8 / tempText.getComputedTextLength()))) + '..';
      }
      svgRoot.removeChild(tempText);
      var attributes = {
        x: x,
        y: labelYOffset,
        class: 'axis-label'
      };
      if (needsRotation) {
        attributes['text-anchor'] = 'end';
        attributes['dominant-baseline'] = 'middle';
        attributes.transform = 'rotate(-45, ' + x + ', ' + labelYOffset + ')';
      } else {
      }
      groups.xAxis.appendChild(createSVGElement('text', attributes, labelContent));
    });
    const lastDataPoints = [];
    const totalLabelCharThreshold = 75;
    config.series.forEach(function(series, seriesIndex) {
        if (!series.values || series.values.length === 0) return;
        const cssPrefix = series.id.toLowerCase();
        const areaBaselineY = yScale(Math.max(0, effectiveYAxisMin));
        var pointsData = series.values.map(function(v, i) {
            var x = xScale(i);
            var finalY = yScale(v);
            var animatedY = areaBaselineY + (finalY - areaBaselineY) * animProgress;
            return {
                x: x,
                finalY: finalY,
                animatedY: animatedY,
                value: v
            };
        });
        var points = pointsData.map(function(d) { return d.x + ',' + d.animatedY; }).join(' ');
        var areaPoints = xScale(0) + ',' + areaBaselineY 
            + ' ' + points + ' ' + xScale(series.values.length - 1) + ',' + areaBaselineY;
        groups.chartArea.appendChild(createSVGElement('polygon', { class: 'area-' + cssPrefix, points: areaPoints }));
        groups.chartArea.appendChild(createSVGElement('polyline', { class: 'line-' + cssPrefix, points: points }));
        var totalLabelChars = 0;
        series.values.forEach(function(value) {
            totalLabelChars += (value != null) ? value.toString().length : 0;
        });
        const skipEveryOtherForThisSeries = (totalLabelChars > totalLabelCharThreshold);
        pointsData.forEach(function(pointData, i) {
            var pointForLabel = { 
                x: pointData.x, 
                y: pointData.finalY,
                value: pointData.value, 
                seriesIndex: seriesIndex,
                skip: (skipEveryOtherForThisSeries && i % 2 === 1)
            };
            lastDataPoints.push(pointForLabel);
            groups.chartArea.appendChild(createSVGElement('circle', { 
                cx: pointData.x, 
                cy: pointData.animatedY,
                r: config.lineOptions.markerRadius, 
                class: 'marker-' + cssPrefix 
            }));
        });
    });
    lastDataPoints.sort(function(a, b) { return a.x - b.x || a.y - b.y; });
    const drawnLabels = [];
    lastDataPoints.forEach(function(point) {
        if (point.skip) {
            return;
        }
        const series = config.series[point.seriesIndex];
        const cssPrefix = series.id.toLowerCase();
        const verticalThreshold = 18;
        const horizontalThreshold = labelSlotWidth * 0.4;
        const pointThreshold = config.lineOptions.markerRadius + 12;
        const yAbove = point.y - config.lineOptions.dataLabelOffsetY - config.lineOptions.markerRadius;
        const yBelow = point.y + config.lineOptions.dataLabelOffsetY + config.lineOptions.markerRadius + 6;
        let finalLabelY = yAbove;
        let shouldDraw = true;
        let collidesUpLabel = false;
        let collidesUpPoint = false;
        for (const drawn of drawnLabels) {
            if (Math.abs(point.x - drawn.x) < horizontalThreshold && Math.abs(yAbove - drawn.y) < verticalThreshold) {
                collidesUpLabel = true;
                break;
            }
        }
        if (!collidesUpLabel) {
             for (const otherPoint of lastDataPoints) { 
                if (point.seriesIndex === otherPoint.seriesIndex) continue;
                if (Math.abs(point.x - otherPoint.x) < horizontalThreshold && Math.abs(yAbove - otherPoint.y) < pointThreshold) {
                    collidesUpPoint = true;
                    break;
                }
            }
        }
        let collidesDownLabel = false;
        if (collidesUpLabel || collidesUpPoint) {
            for (const drawn of drawnLabels) {
                if (Math.abs(point.x - drawn.x) < horizontalThreshold && Math.abs(yBelow - drawn.y) < verticalThreshold) {
                    collidesDownLabel = true;
                    break;
                }
            }
        }
        if (collidesUpLabel) {
            if (collidesDownLabel) {
                shouldDraw = false;
            } else {
                finalLabelY = yBelow;
            }
        } else if (collidesUpPoint) {
            if (collidesDownLabel) {
                finalLabelY = yAbove;
            } else {
                finalLabelY = yBelow;
            }
        } else {
            finalLabelY = yAbove;
        }
        if (shouldDraw) {
            groups.chartArea.appendChild(createSVGElement('text', { 
                x: point.x, 
                y: finalLabelY,
                class: 'data-label data-label-' + cssPrefix,
                style: 'opacity: ' + animOpacity 
            }, point.value));
            drawnLabels.push({x: point.x, y: finalLabelY});
        }
    });
}
    try {
      const configJson = document.getElementById('chart-json-data').textContent;
if (!configJson || configJson.trim() === '') {
        throw new Error("JSON data is empty.");
}
      let fullConfig;
      try {
        fullConfig = JSON.parse(configJson);
} catch (parseError) {
        throw new Error('JSON parsing error: ' + parseError.message);
      }
      if (!fullConfig || typeof fullConfig !== 'object' || !fullConfig.data) {
        throw new Error("Invalid JSON structure: Missing 'data' key.");
}
      const config = fullConfig.data;
      if (config.colors && typeof config.colors === 'string') {
        const seriesIds = config.series.map(function(s) { return s.id; });
const numColors = seriesIds.length;
        let generatedColors;
        const colorKeyword = config.colors.toLowerCase();
const gradientDarken = 15;
        if (colorKeyword === '#gemini') {
          const GEMINI_STOPS = ['#E68A9C', '#9F63D0', '#616AD8'];
const hexPalette = generateGradientScale(GEMINI_STOPS, numColors);
          generatedColors = [];
          for (let i = 0; i < numColors; i++) {
            const hex_start = hexPalette[i];
const rgb_start = hexToRgb(hex_start);
            const hsl = rgbToHsl(rgb_start.r, rgb_start.g, rgb_start.b);
            const h = hsl[0], s = hsl[1], l = hsl[2];
const l_end = Math.max(0, Math.min(100, l - gradientDarken));
            const rgb_end = hslToRgb(h, s, l_end);
            const hex_end = rgbToHex(rgb_end[0], rgb_end[1], rgb_end[2]);
generatedColors.push({
              id: seriesIds[i],
              start: hex_start,
              end: hex_end
            });
}
        } else if (colorKeyword === '#night') {
          var NIGHT_COLORS = [
            { start: '#8d58d3', end: '#704ca4' },
            { start: '#8f83d6', end: 
'#7469a8' },
            { start: '#7891dd', end: '#5c6fbb' },
            { start: '#50a89d', end: '#258a7f' },
            { start: '#54c5d5', end: '#34a0b0' }
          ];
generatedColors = [];
          var M_numColors = numColors;
          var N_paletteSize = NIGHT_COLORS.length;
          if (M_numColors <= N_paletteSize) {
            for (var i = 0; i < M_numColors; i++) {
              var progress = (M_numColors 
=== 1) ? 0 : (i / (M_numColors - 1));
              var colorIndex = Math.round(progress * (N_paletteSize - 1));
var colorPair = NIGHT_COLORS[colorIndex];
              generatedColors.push({
                id: seriesIds[i],
                start: colorPair.start,
                end: colorPair.end
              });
}
          } else {
            for (var i = 0; i < M_numColors; i++) {
              var colorPair = NIGHT_COLORS[i] ||
NIGHT_COLORS[N_paletteSize - 1];
              generatedColors.push({
                id: seriesIds[i],
                start: colorPair.start,
                end: colorPair.end
              });
}
          }
        } else {
          const primaryColor = config.colors;
          generatedColors = generateMonochromaticScale(primaryColor, seriesIds);
          if (generatedColors.length === 0) {
             throw new Error("Invalid HEX color string in 'colors': " + primaryColor);
}
        }
        if (generatedColors && generatedColors.length > 0) {
          config.colors = generatedColors;
} else {
           throw new Error("Failed to generate colors from keyword: " + config.colors);
}
      } else if (!config.colors || !Array.isArray(config.colors)) {
         throw new Error("'colors' must be an array of color objects or a single primary color string.");
}
      if (!config.lineOptions) {
        config.lineOptions = {};
      }
      if (config.lineOptions.markerRadius == null) {
        config.lineOptions.markerRadius = 5;
}
      if (config.lineOptions.dataLabelOffsetY == null) {
        config.lineOptions.dataLabelOffsetY = 8;
}
      if (!config || typeof config !== 'object' || !config.series || !Array.isArray(config.series)) {
        throw new Error("Invalid JSON structure: Missing required fields like 'series' inside 'data'.");
}
      buildChart(config);
    } catch (e) {
      console.error("Chart Error:", e);
const svgRoot = document.getElementById('multi-line-chart-svg');
      const existingError = svgRoot.querySelector('.error-message');
      if(existingError) existingError.remove();
const textEl = createSVGElement('text', {
        x: svgRoot.viewBox.baseVal.width / 2 + (svgRoot.viewBox.baseVal.x || 0),
        y: svgRoot.viewBox.baseVal.height / 2 + (svgRoot.viewBox.baseVal.y || 0),
        'text-anchor': 'middle',
        fill: 'red',
        'font-family': "'Noto Sans JP', sans-serif",
        class: 'error-message'
      });
textEl.textContent = e.message.includes('JSON') ? 'JSONデータにエラーがあります。' : 'グラフ描画エラーが発生しました。';
      if (e.message.includes('JSON parsing error')) {
        textEl.textContent += ' (' + e.message.split(': ')[1] + ')';
      }
      svgRoot.appendChild(textEl);
}
  </script>
</svg>
`;
const STACKED_BAR_CHART_TEMPLATE = `
<svg width="600" height="550" viewBox="25 0 575 570" xmlns="http://www.w3.org/2000/svg" font-family="'Noto Sans JP', sans-serif" id="stacked-bar-chart-svg">
  <script id="chart-json-data" type="application/json">
    {
      "chartType": "stacked-bar",
      "data": {
        "title": "サンプルグラフ：カテゴリ別数量",
        "subtitle": "（月別サンプルデータ）",
        "source": "出典：サンプルデータソース",
        "yAxisUnitLabel": "（単位）",
        "colors": [
          { "id": "A", "start": "#e68a9c", "end": "#d96d8f" },
      { "id": "B", "start": "#b469b8", "end": "#a656ad" },
          { "id": "C", "start": "#9f63d0", "end": "#8c4fc8" },
          { "id": "D", "start": "#7c6ce8", "end": "#6b5ce0" },
          { "id": "E", "start": "#616ad8", "end": "#5059d1" }
        ],
        "legendLabels": ["項目 A", "項目 B", "項目 C", "項目 D", "項目 E"],
        "barData": [
         { "label": "サンプル 1", "values": [60, 50, 40, 25, 20] },
          { "label": "サンプル 2", "values": [90, 60, 50, 40, 25] },
          { "label": "サンプル 3", "values": [60, 80, 70, 50, 30] },
          { "label": "サンプル 4", "values": [40, 50, 60, 70, 80] },
          { "label": "サンプル 5", "values": [70, 30, 55, 45, 65] }
       ],
        "layout": {
          "width": 600,
          "height": 550,
          "marginTop": 170,
          "marginBottom": 50,
          "marginLeft": 75,
          "marginRight": 50
        },
        "barOptions": {
         "width": 50,
          "cornerRadius": 4,
          "totalLabelOffset": 10
        },
        "yAxis": {
          "max": 300,
          "tickCount": 3
        },
        "animation": 1
      }
    }
  </script>
  <defs>
    <filter id="shadow" x="-10%" 
y="-10%" width="120%" height="120%"><feDropShadow dx="1" dy="2" stdDeviation="1.5" flood-color="#000" flood-opacity="0.3"/></filter>
    <filter id="text-shadow"><feDropShadow dx="1" dy="1" stdDeviation="1" flood-color="#000" flood-opacity="0.7"/></filter>
  </defs>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&amp;family=Roboto:wght@400;700&amp;display=swap');
    .bar-segment { filter: url(#shadow);
}
    .axis-label { font-size: 13px; fill: #5f6368; text-anchor: middle;
}
    .title { font-size: 24px;
font-weight: 500; text-anchor: middle; fill: #5f6368; }
    .subtitle { font-size: 16px; text-anchor: middle; fill: #5f6368;
} 
    .source { font-size: 11px; fill: #70757a; text-anchor: end;
}
    .grid-line { stroke: #e0e0e0; stroke-dasharray: 2 2; stroke-width: 0.8;
}
    .legend-item { font-size: 13px; fill: #202124; dominant-baseline: middle;
}
    .segment-label { font-family: 'Roboto', sans-serif; font-size: 13px; font-weight: normal;
fill: white; filter: url(#text-shadow); text-anchor: middle; pointer-events: none; }
    .total-label { font-family: 'Roboto', sans-serif; font-size: 14px;
font-weight: bold; fill: #5f6368; text-anchor: middle; }
    .y-axis-value { font-family: 'Roboto', sans-serif; font-size: 12px; fill: #757575;
text-anchor: end; }
  </style>
  <g id="title-group"></g>
  <g id="legend-group"></g>
  <g id="y-axis-group"></g>
  <g id="chart-area-group"></g>
  <g id="source-group"></g>
  <script type="text/javascript">
    var svgNS = "http://www.w3.org/2000/svg";
function createSVGElement(name, attributes, textContent) {
      var el = document.createElementNS(svgNS, name);
for (var key in attributes) el.setAttribute(key, attributes[key]);
      if (textContent !== undefined && textContent !== null) {
        el.textContent = textContent;
}
      return el;
    }
    function getEstimatedTextWidth(text, fontSize) {
        var width = 0;
        var fullWidthCharSize = fontSize;
        var halfWidthCharSize = fontSize * 0.6;
        if (!text) return 0;
        for (var i = 0; i < text.length; i++) {
            var charCode = text.charCodeAt(i);
            if ((charCode >= 0x0020 && charCode <= 0x007E) || (charCode >= 0xFF61 && charCode <= 0xFF9F)) {
                width += halfWidthCharSize;
            } else {
                width += fullWidthCharSize;
            }
        }
        return width;
    }
    function hexToRgb(hex) {
      var shorthandRegex = /^#?([a-f0-9])([a-f0-9])([a-f0-9])$/i;
hex = hex.replace(shorthandRegex, function(m, r, g, b) {
        return r + r + g + g + b + b;
      });
var result = /^#?([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})$/i.exec(hex);
      return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
      } : null;
}
    function rgbToHsl(r, g, b) {
      r /= 255;
g /= 255; b /= 255;
      var max = Math.max(r, g, b), min = Math.min(r, g, b);
var h, s, l = (max + min) / 2;
      l = Math.round(l * 100);
      if (max == min) {
        h = s = 0;
      } else {
        var d = max - min;
s = l > 50 ? d / (2 - max - min) : d / (max + min);
s = Math.round(s * 100);
        switch (max) {
          case r: h = (g - b) / d + (g < b ? 6 : 0);
break;
          case g: h = (b - r) / d + 2; break;
case b: h = (r - g) / d + 4; break;
}
        h /= 6;
}
      h = Math.round(h * 360);
      return [h, s, l];
}
    function hslToRgb(h, s, l) {
      var r, g, b;
h /= 360;
      s /= 100;
      l /= 100;
      if (s == 0) {
        r = g = b = l;
      } else {
        function hue2rgb(p, q, t) {
          if (t < 0) t += 1;
if (t > 1) t -= 1;
          if (t < 1/6) return p + (q - p) * 6 * t;
if (t < 1/2) return q;
          if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
return p;
        }
        var q = l < 0.5 ?
l * (1 + s) : l + s - l * s;
var p = 2 * l - q;
        r = hue2rgb(p, q, h + 1/3);
        g = hue2rgb(p, q, h);
b = hue2rgb(p, q, h - 1/3);
      }
      return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
}
    function rgbToHex(r, g, b) {
      function componentToHex(c) {
        var hex = c.toString(16);
return hex.length == 1 ? "0" + hex : hex;
}
      return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
}
    function linearInterpolateColor(hex1, hex2, t) {
      var rgb1 = hexToRgb(hex1);
var rgb2 = hexToRgb(hex2);
      var r = Math.round(rgb1.r + (rgb2.r - rgb1.r) * t);
var g = Math.round(rgb1.g + (rgb2.g - rgb1.g) * t);
      var b = Math.round(rgb1.b + (rgb2.b - rgb1.b) * t);
return rgbToHex(r, g, b);
    }
    function generateGradientScale(gradientStops, numSteps) {
      var palette = [];
if (numSteps <= 0) return [];
      if (numSteps === 1) return [gradientStops[0]];
      var numSegments = gradientStops.length - 1;
for (var i = 0; i < numSteps; i++) {
        var t_global = (numSteps === 1) ?
0 : (i / (numSteps - 1));
        var segmentProgress = t_global * numSegments;
var stopIndex1 = Math.floor(segmentProgress);
        var stopIndex2 = Math.min(stopIndex1 + 1, numSegments);
        var t_local = segmentProgress - stopIndex1;
var hex1 = gradientStops[stopIndex1];
        var hex2 = gradientStops[stopIndex2];
        palette.push(linearInterpolateColor(hex1, hex2, t_local));
      }
      return palette;
}
    function generateMonochromaticScale(primaryHex, colorIds) {
      var numSteps = colorIds.length;
if (numSteps === 0) return [];
      var rgb = hexToRgb(primaryHex);
      if (!rgb) return [];
      var hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
var h = hsl[0], s = hsl[1], l_mid_original = hsl[2];
      var gradientDarken = 7;
      var L_MAX = 80; 
      var L_MIN = 45;
var l_step_amount = numSteps > 1 ? (L_MAX - L_MIN) / (numSteps - 1) : 0;
      var l_steps_calc = [];
for (var i = 0; i < numSteps; i++) {
        var l_val = L_MAX - (i * l_step_amount);
l_steps_calc.push(l_val);
      }
      var l_steps_light = [];
var l_steps_dark = [];
      var mid_index = Math.ceil(numSteps / 2);
      for (var i = 0; i < mid_index; i++) {
        l_steps_light.push(l_steps_calc[i]);
}
      for (var i = mid_index; i < numSteps; i++) {
        l_steps_dark.push(l_steps_calc[i]);
}
      var l_steps_start = [];
for (var i = 0; i < numSteps; i++) {
        if (i % 2 === 0) {
          l_steps_start.push(l_steps_light[i / 2]);
} else {
          l_steps_start.push(l_steps_dark[(i - 1) / 2]);
}
      }
      var colors = [];
for (var i = 0; i < numSteps; i++) {
        var l_start_step = l_steps_start[i];
var l_start = Math.max(0, Math.min(100, l_start_step));
        var l_end = Math.max(0, Math.min(100, l_start - gradientDarken));
        var s_adjusted = s;
if (l_start > 85) s_adjusted = s * (1 - (l_start - 85) / 15);
if (l_start < 20) s_adjusted = s * (l_start / 20);
        s_adjusted = Math.max(0, Math.min(100, s_adjusted));
var rgb_start = hslToRgb(h, s_adjusted, l_start);
        var hex_start = rgbToHex(rgb_start[0], rgb_start[1], rgb_start[2]);
        var rgb_end = hslToRgb(h, s_adjusted, l_end);
var hex_end = rgbToHex(rgb_end[0], rgb_end[1], rgb_end[2]);
        colors.push({
          id: colorIds[i],
          start: hex_start,
          end: hex_end
        });
}
      return colors;
    }
    function buildChart(config) {
      var chartConfig = config;
var colorConfig = config.colors;
      var svgRoot = document.getElementById('stacked-bar-chart-svg');
      var defs = svgRoot.querySelector('defs');
      var existingGrads = defs.querySelectorAll('linearGradient');
      existingGrads.forEach(function(grad) { grad.remove(); });
      var styleContent = document.querySelector('style').textContent;
      var dynamicRulePattern = /\.bar-[a-z]\s*{[^}]*}/g;
      styleContent = styleContent.replace(dynamicRulePattern, '');
      document.querySelector('style').textContent = styleContent.trim();
      var dynamicStyles = [];
      colorConfig.forEach(function(color) {
        var gradId = 'grad' + color.id;
        var cssClass = 'bar-' + color.id.toLowerCase();
        var gradient = createSVGElement('linearGradient', { id: gradId, x1:"0", x2:"0", y1:"0", y2:"1" });
        gradient.appendChild(createSVGElement('stop', { offset: '0%', 'stop-color': color.start }));
        gradient.appendChild(createSVGElement('stop', { offset: '100%', 'stop-color': color.end }));
        defs.appendChild(gradient);
        dynamicStyles.push('.' + cssClass + ' { fill: url(#' + gradId + '); }');
      });
      document.querySelector('style').textContent += dynamicStyles.join(String.fromCharCode(10));
      var titleGroup = document.getElementById('title-group');
var legendGroup = document.getElementById('legend-group');
      var yAxisGroup = document.getElementById('y-axis-group');
      var chartAreaGroup = document.getElementById('chart-area-group');
      var sourceGroup = document.getElementById('source-group');
      var groups = [titleGroup, legendGroup, yAxisGroup, chartAreaGroup, sourceGroup];
for(var i = 0; i < groups.length; i++) {
        var group = groups[i];
while (group.firstChild) group.removeChild(group.firstChild);
      }
      var chartAreaWidth = chartConfig.layout.width - chartConfig.layout.marginLeft - chartConfig.layout.marginRight;
var chartAreaHeight = chartConfig.layout.height - chartConfig.layout.marginTop - chartConfig.layout.marginBottom;
      var bottomY = chartConfig.layout.marginTop + chartAreaHeight;
      var totals = chartConfig.barData.map(function(d) { 
        return d.values.reduce(function(sum, v) { return sum + v; }, 0); 
      });
var maxTotal = Math.max.apply(null, totals.concat([config.yAxis.max]));
      var tickCount = chartConfig.yAxis.tickCount > 0 ?
chartConfig.yAxis.tickCount : 1;
      var preliminaryTickInterval = (config.yAxis.max > 0 && tickCount > 0) ? config.yAxis.max / tickCount : 1;
      var effectiveYAxisMax = Math.max(config.yAxis.max, Math.ceil(maxTotal / preliminaryTickInterval) * preliminaryTickInterval);
      if (effectiveYAxisMax <= 0) effectiveYAxisMax = preliminaryTickInterval * tickCount;
      if (effectiveYAxisMax === 0) effectiveYAxisMax = 1;
      var yAxisRange = effectiveYAxisMax;
      var xScale = function(index) { return chartConfig.layout.marginLeft + (chartAreaWidth / chartConfig.barData.length) * (index + 0.5);
};
      var yScale = function(value) { return bottomY - (value / effectiveYAxisMax) * chartAreaHeight; };
      var animProgress = (config.animation == null) 
          ? 1.0 
          : Math.max(0, Math.min(1, Number(config.animation)));
      var animOpacity;
      var fadeStart = 0.7;
      if (animProgress <= fadeStart) {
          animOpacity = 0;
      } else {
          animOpacity = (animProgress - fadeStart) / (1.0 - fadeStart);
      }
      var svgWidth = svgRoot.viewBox.baseVal.width;
titleGroup.appendChild(createSVGElement('text', { x: svgWidth/2, y: 50, class: 'title' }, chartConfig.title));
      titleGroup.appendChild(createSVGElement('text', { x: svgWidth/2, y: 75, class: 'subtitle' }, chartConfig.subtitle));
sourceGroup.appendChild(createSVGElement('text', { x: 600, y: 555, class: 'source' }, chartConfig.source));
//
      var legendY = 110;
      var numLegends = chartConfig.legendLabels.length;
      var legendStartX = chartConfig.layout.marginLeft;
      var rectSize = 12;
      var rectTextGap = 6;
      var legendFontSize = 13;
      var legendPadding = 15;
      var requiredWidths = chartConfig.legendLabels.map(function(label) {
          var textWidth = getEstimatedTextWidth(label, legendFontSize);
          return rectSize + rectTextGap + textWidth + legendPadding;
      });
      var totalRequiredWidth = requiredWidths.reduce(function(sum, w) { return sum + w; }, 0) - legendPadding;
      if (numLegends === 0) totalRequiredWidth = 0;
      var useEqualWidths = false;
      var legendColumnWidth = 0;
      if (totalRequiredWidth > chartAreaWidth || numLegends === 0) {
          useEqualWidths = true;
          legendColumnWidth = (numLegends > 0) ? (chartAreaWidth / numLegends) : 0;
      } else {
          useEqualWidths = false;
          legendStartX = chartConfig.layout.marginLeft + (chartAreaWidth - totalRequiredWidth) / 2;
      }
      var currentX = legendStartX;
      chartConfig.legendLabels.forEach(function(label, i) {
        var cssClass = 'bar-' + colorConfig[i].id.toLowerCase();
        var columnStartX;
        var currentColumnWidth;
        if (useEqualWidths) {
            columnStartX = legendStartX + (i * legendColumnWidth);
            currentColumnWidth = legendColumnWidth;
        } else {
            columnStartX = currentX;
            currentColumnWidth = requiredWidths[i];
            currentX += currentColumnWidth;
        }
        var itemGroup = createSVGElement('g');
        itemGroup.appendChild(createSVGElement('rect', {
          x: columnStartX, y: legendY, width: rectSize, height: rectSize, rx: 2, class: 'bar-segment ' + cssClass
        }));
        var textX = columnStartX + rectSize + rectTextGap;
        var clipId = 'legend-clip-' + i;
        var clipMargin = (legendPadding / 2);
        if (useEqualWidths) clipMargin = 5;
        var clipWidth = (currentColumnWidth - rectSize - rectTextGap - clipMargin); 
        clipWidth = Math.max(0, clipWidth); 
        var clipPath = createSVGElement('clipPath', { id: clipId });
        clipPath.appendChild(createSVGElement('rect', {
            x: textX,
            y: legendY - (rectSize / 2), 
            width: clipWidth,
            height: rectSize * 2 
        }));
        defs.appendChild(clipPath);
        itemGroup.appendChild(createSVGElement('text', { 
            x: textX, 
            y: legendY + rectSize / 2, 
            class: 'legend-item',
            'clip-path': 'url(#' + clipId + ')' 
        }, label));
         legendGroup.appendChild(itemGroup);
      });
//
yAxisGroup.appendChild(createSVGElement('text', { x: 65, y: chartConfig.layout.marginTop - 30, class: 'y-axis-value' }, chartConfig.yAxisUnitLabel));
      var finalTickStep = yAxisRange / tickCount;
      for (var i = 0; i <= tickCount; i++) {
        var value = Math.round(finalTickStep * i);
        var y = yScale(value);
yAxisGroup.appendChild(createSVGElement('line', { x1: chartConfig.layout.marginLeft, y1: y, x2: chartConfig.layout.width - chartConfig.layout.marginRight, y2: y, class: 'grid-line' }));
yAxisGroup.appendChild(createSVGElement('text', { x: 65, y: y, class: 'y-axis-value', 'dominant-baseline': 'middle' }, value));
}
      chartConfig.barData.forEach(function(bar, index) {
        var barGroup = createSVGElement('g', { class: 'bar-group' });
        var centerX = xScale(index);
        var total = bar.values.reduce(function(sum, v) { return sum + v; }, 0);
        var finalTotal = total;
        var animatedTotal = finalTotal * animProgress;
        if (finalTotal > 0) {
          barGroup.appendChild(createSVGElement('text', { 
            x: centerX, 
            y: yScale(animatedTotal) - chartConfig.barOptions.totalLabelOffset,
            class: 'total-label',
            style: 'opacity: ' + animOpacity
          }, finalTotal));
        }
        barGroup.appendChild(createSVGElement('text', { x: centerX, y: bottomY + 20, class: 'axis-label' }, bar.label));
        var yOffset = 0;
        bar.values.slice().reverse().forEach(function(value, i) {
          if (value <= 0) return;
          var finalValue = value;
          var finalYOffset = yOffset;
          var animatedYOffset = finalYOffset * animProgress;
          var animatedValue = finalValue * animProgress;
          var segmentIndex = bar.values.length - 1 - i;
          var cssClass = 'bar-' + colorConfig[segmentIndex].id.toLowerCase();
          var yTop = yScale(animatedYOffset + animatedValue);
          var yBottom = yScale(animatedYOffset);
          var height = yBottom - yTop;
          if (height <= 0.5) return;
          var r = chartConfig.barOptions.cornerRadius;
var x = centerX - chartConfig.barOptions.width / 2;
          var w = chartConfig.barOptions.width;
          var segment;
          var isBottomSegment = (finalYOffset === 0);
          var isTopSegment = Math.abs((finalYOffset + finalValue) - finalTotal) < 0.01;
          if (isBottomSegment && isTopSegment) {
            segment = createSVGElement('rect', { x: x, y: yTop, width: w, height: height, rx: r, ry: r, class: 'bar-segment ' + cssClass });
          } else if (isBottomSegment) {
            var d = 'M ' + x + ' ' + yTop + ' H ' + (x+w) + ' V ' + (yBottom-r) + ' A ' + r + ' ' + r + ' 0 0 1 ' + (x+w-r) + ' ' + yBottom + ' H ' + (x+r) + 
' A ' + r + ' ' + r + ' 0 0 1 ' + x + ' ' + (yBottom-r) + ' Z';
            segment = createSVGElement('path', { d: d, class: 'bar-segment ' + cssClass });
          } else if (isTopSegment) {
            var d = 'M ' + x + ' ' + yBottom + ' H ' + (x+w) + ' V ' + (yTop+r) + ' A ' + r + ' ' + r + ' 0 0 0 ' + (x+w-r) + ' ' + yTop + ' H ' + (x+r) + 
' A ' + r + ' ' + r + ' 0 0 0 ' + x + ' ' + (yTop+r) + ' Z';
            segment = createSVGElement('path', { d: d, class: 'bar-segment ' + cssClass });
          } else {
            segment = createSVGElement('rect', { x: x, y: yTop, width: w, height: height, class: 'bar-segment ' + cssClass });
          }
          barGroup.appendChild(segment);
          if(height > 15){
            barGroup.appendChild(createSVGElement('text', { 
                x: centerX, 
                y: yTop + height / 2,
                'dominant-baseline': 'middle', 
                class: 'segment-label',
                style: 'opacity: ' + animOpacity
            }, finalValue));
}
          yOffset += value;
        });
        chartAreaGroup.appendChild(barGroup);
      });
}
    try {
      var configJson = document.getElementById('chart-json-data').textContent;
if (!configJson || configJson.trim() === '') {
        throw new Error("JSON data is empty.");
}
      var fullConfig;
      try {
        fullConfig = JSON.parse(configJson);
} catch (parseError) {
        throw new Error('JSON parsing error: ' + parseError.message);
      }
      if (!fullConfig || typeof fullConfig !== 'object' || !fullConfig.data) {
        throw new Error("Invalid JSON structure: Missing 'data' key.");
}
      var config = fullConfig.data;
      if (config.colors && typeof config.colors === 'string') {
        var numColors = config.legendLabels ?
config.legendLabels.length : 5;
        var generatedIds = [];
        for (var i = 0; i < numColors; i++) {
          generatedIds.push(String.fromCharCode(65 + i));
        }
        var generatedColors;
        var colorKeyword = config.colors.toLowerCase();
        var gradientDarken = 15;
        if (colorKeyword === '#gemini') {
          var GEMINI_STOPS = ['#E68A9C', '#9F63D0', '#616AD8'];
var hexPalette = generateGradientScale(GEMINI_STOPS, numColors);
          generatedColors = [];
          for (var i = 0; i < numColors; i++) {
            var hex_start = hexPalette[i];
var rgb_start = hexToRgb(hex_start);
            var hsl = rgbToHsl(rgb_start.r, rgb_start.g, rgb_start.b);
            var h = hsl[0], s = hsl[1], l = hsl[2];
var l_end = Math.max(0, Math.min(100, l - gradientDarken));
            var rgb_end = hslToRgb(h, s, l_end);
            var hex_end = rgbToHex(rgb_end[0], rgb_end[1], rgb_end[2]);
generatedColors.push({
              id: generatedIds[i],
              start: hex_start,
              end: hex_end
            });
}
        } else if (colorKeyword === '#night') {
          var NIGHT_COLORS = [
            { start: '#8d58d3', end: '#704ca4' },
            { start: '#8f83d6', end: '#7469a8' 
},
            { start: '#7891dd', end: '#5c6fbb' },
            { start: '#50a89d', end: '#258a7f' },
            { start: '#54c5d5', end: '#34a0b0' }
          ];
generatedColors = [];
          for (var i = 0; i < numColors; i++) {
            var colorPair = NIGHT_COLORS[i] ||
NIGHT_COLORS[NIGHT_COLORS.length - 1];
            generatedColors.push({
              id: generatedIds[i],
              start: colorPair.start,
              end: colorPair.end
            });
}
        } else {
          var primaryColor = config.colors;
          generatedColors = generateMonochromaticScale(primaryColor, generatedIds);
          if (generatedColors.length === 0) {
             throw new Error("Invalid HEX color string in 'colors': " + primaryColor);
}
        }
        if (generatedColors && generatedColors.length > 0) {
          config.colors = generatedColors;
} else {
           throw new Error("Failed to generate colors from keyword: " + config.colors);
}
      } else if (!config.colors || !Array.isArray(config.colors)) {
         throw new Error("'colors' must be an array of color objects or a single primary color string.");
}
      if (!config.layout) {
        config.layout = {
          width: 600,
          height: 550,
          marginTop: 170,
          marginBottom: 50,
          marginLeft: 75,
          marginRight: 50
        };
      }
      if (!config.barOptions) {
        config.barOptions = {
          width: 50,
          cornerRadius: 4,
          totalLabelOffset: 10
        };
      }
      if (!config || typeof config !== 'object' || !config.barData || !Array.isArray(config.barData)) {
        throw new Error("Invalid JSON structure: Missing required fields like 'barData' inside 'data'.");
}
      buildChart(config);
    } catch(e) {
      console.error("Chart Error:", e);
var svgRoot = document.getElementById('stacked-bar-chart-svg');
      var existingError = svgRoot.querySelector('.error-message');
      if(existingError) existingError.remove();
var textEl = createSVGElement('text', {
        x: svgRoot.viewBox.baseVal.width / 2 + svgRoot.viewBox.baseVal.x,
        y: svgRoot.viewBox.baseVal.height / 2 + svgRoot.viewBox.baseVal.y,
        'text-anchor': 'middle', fill: 'red',
        'font-family': "'Noto Sans JP', sans-serif",
        class: 'error-message'
      });
textEl.textContent = e.message.includes('JSON') ? 'JSONデータにエラーがあります。' : 'グラフ描画エラーが発生しました。';
      if (e.message.includes('JSON parsing error')) {
        textEl.textContent += ' (' + e.message.split(': ')[1] + ')';
      }
      svgRoot.appendChild(textEl);
}
  </script>
</svg>
`;
const BAR_CHART_TEMPLATE = `
<svg width="600" height="450" viewBox="25 25 570 435" xmlns="http://www.w3.org/2000/svg" font-family="'Noto Sans JP', sans-serif" id="bar-chart-svg">
    <script id="chart-json-data" type="application/json">
{
        "chartType": "bar",
        "data": {
            "title": "サンプル棒グラフ（動的レイアウト）",
            "subtitle": "項目数に応じて幅が均等に調整されます",
            "source": "出典：サンプルデータソース",
            "items": [
                { "label": "項目A", "value": 85 },
                { "label": "項目B", "value": 72 },
                { "label": "項目C", "value": 93 },
                { "label": "項目D", "value": 65 },
                { "label": "項目E", "value": 48 },
                { "label": "項目F", "value": 78 },
                { "label": "項目G", "value": 55 },
                { "label": "項目H", "value": 30 }
             ],
            "colors": {
                "start": "#e68a9c",
                "end": "#9f63d0"
            },
            "layout": {
                "width": 600,
                "height": 450,
                "marginTop": 100,
                "marginBottom": 65,
                "marginLeft": 70,
                "marginRight": 40
            },
            "barOptions": {
                "barToSlotRatio": 0.6
            },
            "yAxis": {
                "max": 100,
                "min": 0,
                "tickCount": 4,
                "unit": "%"
            },
            "animation": 1
        }
    }
    </script>
    <defs>
        <filter id="shadow">
            <feDropShadow dx="1" dy="2" stdDeviation="2" flood-color="#000" flood-opacity="0.2"/>
        </filter>
         <filter id="text-shadow">
            <feDropShadow dx="0" dy="0" stdDeviation="2.5" flood-color="#000000" flood-opacity="0.9"/>
        </filter>
    </defs>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500&amp;family=Roboto:wght@400&amp;display=swap');
        .bar { filter: url(#shadow); }
        .axis-label { font-size: 13px; fill: #5f6368; text-anchor: middle; }
        .title { font-size: 21px; font-weight: 500; text-anchor: middle; fill: #5f6368; }
        .subtitle { font-size: 16px; text-anchor: middle; fill: #5f6368; }
        .source { font-size: 11px; fill: #70757a; text-anchor: end; }
        .grid-line { stroke: #e0e0e0; stroke-dasharray: 2 2; stroke-width: 0.8; }
        .y-axis-unit { font-size: 12px; fill: #5f6368; text-anchor: end; } 
        .data-label { font-family: 'Roboto', sans-serif; font-size: 14px; font-weight: normal; text-anchor: middle; }
        .y-axis-value { font-family: 'Roboto', sans-serif; font-size: 12px; fill: #757575; text-anchor: end; }
    </style>
    <g id="title-group"></g>
    <g id="y-axis-group"></g>
    <g id="chart-area-group"></g>
    <g id="x-axis-group"></g>
    <g id="source-group"></g>
    <script type="text/javascript">
    const svgNS = "http://www.w3.org/2000/svg";
    function createSVGElement(name, attributes, textContent) {
        const el = document.createElementNS(svgNS, name);
        for (const key in attributes) el.setAttribute(key, attributes[key]);
        if (textContent !== undefined && textContent !== null) {
            el.textContent = textContent;
        }
        return el;
    }
    function hexToRgb(hex) {
      let shorthandRegex = /^#?([a-f0-9])([a-f0-9])([a-f0-9])$/i;
      hex = hex.replace(shorthandRegex, function(m, r, g, b) {
        return r + r + g + g + b + b;
      });
      let result = /^#?([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})$/i.exec(hex);
      return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
      } : null;
    }
    function rgbToHsl(r, g, b) {
      r /= 255; g /= 255; b /= 255;
      let max = Math.max(r, g, b), min = Math.min(r, g, b);
      let h, s, l = (max + min) / 2;
      l = Math.round(l * 100);
      if (max == min) {
        h = s = 0;
      } else {
        let d = max - min;
        s = l > 50 ? d / (2 - max - min) : d / (max + min);
        s = Math.round(s * 100);
        switch (max) {
          case r: h = (g - b) / d + (g < b ? 6 : 0); break;
          case g: h = (b - r) / d + 2; break;
          case b: h = (r - g) / d + 4; break;
        }
        h /= 6;
      }
      h = Math.round(h * 360);
      return [h, s, l];
    }
    function hslToRgb(h, s, l) {
      let r, g, b;
      h /= 360;
      s /= 100;
      l /= 100;
      if (s == 0) {
        r = g = b = l;
      } else {
        function hue2rgb(p, q, t) {
          if (t < 0) t += 1;
          if (t > 1) t -= 1;
          if (t < 1/6) return p + (q - p) * 6 * t;
          if (t < 1/2) return q;
          if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
          return p;
        }
        let q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        let p = 2 * l - q;
        r = hue2rgb(p, q, h + 1/3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1/3);
      }
      return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
    }
    function rgbToHex(r, g, b) {
      function componentToHex(c) {
        let hex = c.toString(16);
        return hex.length == 1 ? "0" + hex : hex;
      }
      return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
    }
    function clampLuminance(hex, minL, maxL) {
      var rgb = hexToRgb(hex);
      if (!rgb) return hex;
      var hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
      var h = hsl[0], s = hsl[1], l = hsl[2];
      var clamped_l = Math.max(minL, Math.min(maxL, l));
      if (clamped_l === l) {
        return hex;
      }
      var clamped_rgb = hslToRgb(h, s, clamped_l);
      return rgbToHex(clamped_rgb[0], clamped_rgb[1], clamped_rgb[2]);
    }
function truncateText(text, maxWidth) {
    const AVERAGE_CHAR_WIDTH = 10; 
    const ELLIPSIS_WIDTH = 3 * AVERAGE_CHAR_WIDTH; 
    if (text.length * AVERAGE_CHAR_WIDTH <= maxWidth) {
        return text;
    }
    const maxVisibleChars = Math.floor((maxWidth - ELLIPSIS_WIDTH) / AVERAGE_CHAR_WIDTH);
    if (maxVisibleChars <= 0) {
        return '...';
    }
    return text.slice(0, maxVisibleChars) + '...';
}
    function buildChart(config) {
        const svgRoot = document.getElementById('bar-chart-svg');
        const defs = svgRoot.querySelector('defs');
        let animRatio = (config.animation == null) ? 1.0 : parseFloat(config.animation);
        config.animationRatio = Math.max(0.0, Math.min(1.0, isNaN(animRatio) ? 1.0 : animRatio));
        const existingGrad = document.getElementById('barGradient');
        if (existingGrad) existingGrad.remove();
        const styleElement = svgRoot.querySelector('style');
        if (styleElement) {
            styleElement.textContent = styleElement.textContent.replace(/[.]bar {[^}]*}/g, '');
        }
        const gradId = 'barGradient';
        let gradient = createSVGElement('linearGradient', { id: gradId, x1:"0", x2:"0", y1:"0", y2:"1" });
        gradient.appendChild(createSVGElement('stop', { offset: '0%', 'stop-color': config.color.start }));
        gradient.appendChild(createSVGElement('stop', { offset: '100%', 'stop-color': config.color.end }));
        defs.appendChild(gradient);
        const barStyleRule = '.bar { fill: url(#' + gradId + '); filter: url(#shadow); }';
        if(styleElement && !styleElement.textContent.includes(barStyleRule)) {
             styleElement.textContent += barStyleRule;
        }
        const groups = {
            titleGroup: document.getElementById('title-group'),
            yAxisGroup: document.getElementById('y-axis-group'),
            chartAreaGroup: document.getElementById('chart-area-group'),
            xAxisGroup: document.getElementById('x-axis-group'),
            sourceGroup: document.getElementById('source-group')
        };
        for (let key in groups) {
            while (groups[key].firstChild) groups[key].removeChild(groups[key].firstChild);
        }
        const chartAreaHeight = config.layout.height - config.layout.marginTop - config.layout.marginBottom;
        const dataMax = Math.max(...config.items.map(item => item.value), config.yAxis.max);
        const dataMin = Math.min(...config.items.map(item => item.value), config.yAxis.min);
        let effectiveYAxisMax, effectiveYAxisMin;
        const tickCount = config.yAxis.tickCount > 0 ? config.yAxis.tickCount : 1;
        const preliminaryRange = Math.max(1, config.yAxis.max - config.yAxis.min);
        let tickInterval = preliminaryRange / tickCount;
        if (dataMin < config.yAxis.min) {
            const limit = Math.max(Math.abs(dataMin), dataMax);
            if (tickInterval <= 0 || isNaN(tickInterval)) tickInterval = Math.max(1, limit / 5);
             effectiveYAxisMax = Math.ceil(limit / tickInterval) * tickInterval;
             effectiveYAxisMin = -effectiveYAxisMax;
             if (config.yAxis.min > effectiveYAxisMin) effectiveYAxisMin = config.yAxis.min;
             if (config.yAxis.max < effectiveYAxisMax) effectiveYAxisMax = config.yAxis.max;
        } else {
             effectiveYAxisMin = config.yAxis.min;
             if (tickInterval <= 0 || isNaN(tickInterval)) tickInterval = Math.max(1, (dataMax - effectiveYAxisMin) / 5) || 1;
             effectiveYAxisMax = Math.max(config.yAxis.max, Math.ceil(dataMax / tickInterval) * tickInterval);
        }
        if (effectiveYAxisMax <= effectiveYAxisMin) {
            effectiveYAxisMax = effectiveYAxisMin + tickInterval * tickCount;
            if (effectiveYAxisMax <= effectiveYAxisMin) effectiveYAxisMax = effectiveYAxisMin + 1;
        }
        const yRange = effectiveYAxisMax - effectiveYAxisMin;
        const yScale = (value) => (yRange === 0) ? config.layout.marginTop + chartAreaHeight / 2 : config.layout.marginTop + chartAreaHeight * (1 - (value - effectiveYAxisMin) / yRange);
        groups.titleGroup.appendChild(createSVGElement('text', { x: 300, y: 50, class: 'title' }, config.title));
        groups.titleGroup.appendChild(createSVGElement('text', { x: 300, y: 75, class: 'subtitle' }, config.subtitle));
        groups.yAxisGroup.appendChild(createSVGElement('text', { x: config.layout.marginLeft, y: config.layout.marginTop - 25, class: 'y-axis-unit' }, '(' + config.yAxis.unit + ')'));
        const finalTickStep = yRange / tickCount;
        for (let i = 0; i <= tickCount; i++) {
            const tickValue = effectiveYAxisMin + (finalTickStep * i);
            const y = yScale(tickValue);
            groups.yAxisGroup.appendChild(createSVGElement('text', { x: config.layout.marginLeft - 10, y: y, class: 'y-axis-value', 'dominant-baseline': 'middle' }, Math.round(tickValue)));
            if (i > 0 && i <= tickCount && Math.abs(tickValue) > 1e-6) {
                groups.yAxisGroup.appendChild(createSVGElement('line', { x1: config.layout.marginLeft, y1: y, x2: config.layout.width - config.layout.marginRight, y2: y, class: 'grid-line' }));
            }
        }
        const bottomY = config.layout.marginTop + chartAreaHeight;
        groups.xAxisGroup.appendChild(createSVGElement('line', { x1: config.layout.marginLeft, y1: bottomY, x2: config.layout.width - config.layout.marginRight, y2: bottomY, stroke: '#bdbdbd' }));
        if (effectiveYAxisMin < 0 && effectiveYAxisMax > 0) {
            const yZero = yScale(0);
            groups.yAxisGroup.appendChild(createSVGElement('line', { x1: config.layout.marginLeft, y1: yZero, x2: config.layout.width - config.layout.marginRight, y2: yZero, stroke: '#bdbdbd', 'stroke-width': 1 }));
        }
        const plotAreaWidth = config.layout.width - config.layout.marginLeft - config.layout.marginRight;
        const slotWidth = plotAreaWidth / config.items.length;
        const barWidth = slotWidth * config.barOptions.barToSlotRatio;
        const yZero = yScale(Math.max(0, effectiveYAxisMin));
        const rotationThreshold = 50; 
        const numItems = config.items.length;
        const needsRotation = (slotWidth < rotationThreshold && numItems > 1);
        const labelYOffset = needsRotation ? bottomY + 25 : bottomY + 30;
        const animationRatio = config.animationRatio;
        let animationOpacity;
        const fadeStart = 0.7;
        if (animationRatio <= fadeStart) {
            animationOpacity = 0;
        } else {
            animationOpacity = (animationRatio - fadeStart) / (1.0 - fadeStart);
        }
        config.items.forEach((item, index) => {
            const isEllipsisItem = (item.label === '...');
            const animatedValue = item.value * animationRatio;
            const centerX = config.layout.marginLeft + (index * slotWidth) + (slotWidth / 2);
            const barX = centerX - barWidth / 2;
            let barY, barHeight;
            if (item.value >= 0) {
                barY = yScale(animatedValue); 
                barHeight = Math.max(0, yZero - barY); 
            } else {
                barY = yZero;
                barHeight = Math.max(0, yScale(animatedValue) - yZero); 
            }
            if (barHeight >= 1 && !isEllipsisItem) {
                 groups.chartAreaGroup.appendChild(createSVGElement('rect', { x: barX, y: barY, width: barWidth, height: barHeight, rx: 4, ry: 4, class: 'bar' }));
            }
            const isOutside = barHeight < 30;
            if (!isEllipsisItem) {
                const dataLabel = createSVGElement('text', { 
                    x: centerX, 
                    class: 'data-label',
                    style: 'opacity: ' + animationOpacity
                }, item.value);
                if (isOutside) {
                    dataLabel.setAttribute('fill', '#424242');
                    dataLabel.setAttribute('filter', 'none');
                    dataLabel.setAttribute('y', item.value >= 0 ? barY - 8 : barY + barHeight + 18);
                } else {
                    dataLabel.setAttribute('fill', 'white');
                    dataLabel.setAttribute('filter', 'url(#text-shadow)');
                    dataLabel.setAttribute('y', item.value >= 0 ? barY + 20 : barY + barHeight - 10);
                }
                groups.chartAreaGroup.appendChild(dataLabel);
            }
            let labelContent;
            if (needsRotation && !isEllipsisItem) {
                labelContent = item.label;
            } else {
                if (isEllipsisItem) {
                    labelContent = item.label;
                } else {
                    labelContent = truncateText(item.label, slotWidth * 0.9);
                }
            }
            var attributes = {
                x: centerX,
                y: labelYOffset,
                class: 'axis-label'
            };
            if (needsRotation && !isEllipsisItem) {
                attributes['text-anchor'] = 'end';
                attributes['dominant-baseline'] = 'middle';
                attributes.transform = 'rotate(-45, ' + centerX + ', ' + labelYOffset + ')';
            } else {
            }
            groups.xAxisGroup.appendChild(createSVGElement('text', attributes, labelContent));
        });
        groups.sourceGroup.appendChild(createSVGElement('text', { x: 595, y: config.layout.height - 10, class: 'source' }, config.source));
    }
    try {
        const configJson = document.getElementById('chart-json-data').textContent;
        if (!configJson || configJson.trim() === '') {
            throw new Error("JSON data is empty.");
        }
        let fullConfig;
        try {
            fullConfig = JSON.parse(configJson);
        } catch (parseError) {
             throw new Error('JSON parsing error: ' + parseError.message);
        }
        if (!fullConfig || typeof fullConfig !== 'object' || !fullConfig.data) {
             throw new Error("Invalid JSON structure: Missing 'data' key.");
        }
        const config = fullConfig.data;
        if (config.colors) {
            config.color = config.colors;
        }
        if (config.color && typeof config.color === 'string') {
            const colorKeyword = config.color.toLowerCase();
            let generatedColorObject;
            if (colorKeyword === '#gemini') {
                generatedColorObject = {
                    start: "#E68A9C",
                    end:   "#9F63D0"
                };
            } else if (colorKeyword === '#night') {
                const NIGHT_COLORS_SUBSET = [
                     { start: '#8d58d3', end: '#704ca4' },
                     { start: '#8f83d6', end: '#7469a8' },
                     { start: '#7891dd', end: '#5c6fbb' }
                ];
                generatedColorObject = {
                    start: NIGHT_COLORS_SUBSET[0].start,
                    end:   NIGHT_COLORS_SUBSET[2].end
                };
            } else {
                const primaryColor = config.color;
                const rgb = hexToRgb(primaryColor);
                if (!rgb) {
                    throw new Error("Invalid HEX color string in 'color': " + primaryColor);
                }
                const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
                const h = hsl[0], s = hsl[1], l = hsl[2];
                const l_start = Math.min(100, l + 20);
                const s_start = (l > 80) ? s * 0.7 : s;
                const rgb_start = hslToRgb(h, s_start, l_start);
                const hex_start = rgbToHex(rgb_start[0], rgb_start[1], rgb_start[2]);
                generatedColorObject = {
                    start: hex_start,
                    end:   primaryColor
                };
            }
            config.color = generatedColorObject;
        } else if (!config.color || typeof config.color !== 'object') {
             throw new Error("'color' must be a color object or a single color string.");
        }
        const L_MIN_CLAMP = 30;
        const L_MAX_CLAMP = 70;
        if (config.color && typeof config.color === 'object') {
            if (config.color.start) {
                config.color.start = clampLuminance(config.color.start, L_MIN_CLAMP, L_MAX_CLAMP);
            }
            if (config.color.end) {
                config.color.end = clampLuminance(config.color.end, L_MIN_CLAMP, L_MAX_CLAMP);
            }
        }
        if (!config.barOptions) {
          config.barOptions = {};
        }
        if (config.barOptions.barToSlotRatio == null) {
          config.barOptions.barToSlotRatio = 0.7;
        }
         if (!config || typeof config !== 'object' || !config.items || !Array.isArray(config.items)) {
             throw new Error("Invalid JSON structure: Missing required fields like 'items' inside 'data'.");
         }
        buildChart(config);
    } catch(e) {
        console.error("Chart Error:", e);
        const svgRoot = document.getElementById('bar-chart-svg');
        const existingError = svgRoot.querySelector('.error-message');
        if(existingError) existingError.remove();
        const textEl = createSVGElement('text', {
            x: svgRoot.viewBox.baseVal.width / 2 + svgRoot.viewBox.baseVal.x,
            y: svgRoot.viewBox.baseVal.height / 2 + svgRoot.viewBox.baseVal.y,
            'text-anchor': 'middle', fill: 'red',
            'font-family': "'Noto Sans JP', sans-serif",
            class: 'error-message'
            });
        textEl.textContent = e.message.includes('JSON') ? 'JSONデータにエラーがあります。' : 'グラフ描画エラーが発生しました。';
        if (e.message.includes('JSON parsing error')) {
           textEl.textContent += ' (' + e.message.split(': ')[1] + ')';
        }
        svgRoot.appendChild(textEl);
    }
    </script>
</svg>
`;
const PERCENT_STACKED_BAR_CHART_TEMPLATE = `
<svg width="600" height="510" viewBox="15 15 580 505" xmlns="http://www.w3.org/2000/svg" font-family="'Noto Sans JP', sans-serif" id="100-stacked-bar-chart-svg">
  <script id="chart-json-data" type="application/json">
    {
      "chartType": "100-stacked-bar",
      "data": {
        "title": "サンプルグラフ：カテゴリ別割合",
        "subtitle": "（月別サンプルデータ）",
        "source": "出典：サンプルデータソース",
        "colors": [
          { "id": "A", "start": "#e68a9c", "end": "#d96d8f" },
          { "id": "B", 
"start": "#b469b8", "end": "#a656ad" },
          { "id": "C", "start": "#9f63d0", "end": "#8c4fc8" },
          { "id": "D", "start": "#7c6ce8", "end": "#6b5ce0" },
          { "id": "E", "start": "#616ad8", "end": "#5059d1" }
        ],
        "legendLabels": ["項目 A", "項目 B", "項目 C", "項目 D", "項目 E"],
        "barData": [
{ "label": "サンプル 1", "values": [40, 60, 30, 50, 20] },
          { "label": "サンプル 2", "values": [70, 50, 40, 20, 20] },
          { "label": "サンプル 3", "values": [30, 40, 60, 40, 30] },
          { "label": "サンプル 4", "values": [25, 35, 55, 65, 45] },
          { "label": "サンプル 5", "values": [50, 40, 30, 20, 60] }
        ],
      "layout": {
          "width": 600,
          "height": 510,
          "marginTop": 150,
          "marginBottom": 80,
          "marginLeft": 70,
          "marginRight": 25
        },
        "barOptions": {
"width": 50,
          "cornerRadius": 4
        },
        "yAxis": {
          "tickCount": 4
        },
        "animation": 1
      }
    }
  </script>
  <defs>
    <filter id="shadow"><feDropShadow dx="1" dy="2" stdDeviation="2" flood-color="#000" flood-opacity="0.4"/></filter>
    <filter id="text-shadow"><feDropShadow dx="1" dy="1" stdDeviation="1" flood-color="#000" flood-opacity="0.7"/></filter>
  </defs>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500&amp;family=Roboto:wght@400&amp;display=swap');
    .bar-segment { filter: url(#shadow); }
    .axis-label { font-size: 13px; fill: #5f6368; text-anchor: middle; }
    .title { font-size: 24px; font-weight: 500; text-anchor: middle; fill: #5f6368; }
    .subtitle { font-size: 16px; text-anchor: middle; fill: #5f6368; } 
    .source { font-size: 11px; fill: #70757a; text-anchor: end; }
    .grid-line { stroke: #e0e0e0; stroke-dasharray: 2 2; stroke-width: 0.8; }
    .legend-item { font-size: 13px; fill: #202124; dominant-baseline: middle; }
    .segment-label { font-family: 'Roboto', sans-serif; font-size: 14px; font-weight: normal; fill: white; filter: url(#text-shadow); text-anchor: middle; pointer-events: none; }
    .y-axis-value { font-family: 'Roboto', sans-serif; font-size: 12px; fill: #757575; text-anchor: end; }
  </style>
  <g id="title-group"></g>
  <g id="legend-group"></g>
  <g id="y-axis-group"></g>
  <g id="chart-area-group"></g>
  <g id="source-group"></g>
  <script type="text/javascript">
    const svgNS = "http://www.w3.org/2000/svg";
    function createSVGElement(name, attributes, textContent) {
      const el = document.createElementNS(svgNS, name);
      for (const key in attributes) el.setAttribute(key, attributes[key]);
      if (textContent !== undefined && textContent !== null) {
        el.textContent = textContent;
      }
      return el;
    }
    function getEstimatedTextWidth(text, fontSize) {
        let width = 0;
        const fullWidthCharSize = fontSize;
        const halfWidthCharSize = fontSize * 0.6;
        if (!text) return 0;
        for (let i = 0; i < text.length; i++) {
            const charCode = text.charCodeAt(i);
            if ((charCode >= 0x0020 && charCode <= 0x007E) || (charCode >= 0xFF61 && charCode <= 0xFF9F)) {
                width += halfWidthCharSize;
            } else {
                width += fullWidthCharSize;
            }
        }
        return width;
    }
    function hexToRgb(hex) {
      let shorthandRegex = /^#?([a-f0-9])([a-f0-9])([a-f0-9])$/i;
      hex = hex.replace(shorthandRegex, function(m, r, g, b) {
        return r + r + g + g + b + b;
      });
      let result = /^#?([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})$/i.exec(hex);
      return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
      } : null;
    }
    function rgbToHsl(r, g, b) {
      r /= 255; g /= 255; b /= 255;
      let max = Math.max(r, g, b), min = Math.min(r, g, b);
      let h, s, l = (max + min) / 2;
      l = Math.round(l * 100);
      if (max == min) {
        h = s = 0;
      } else {
        let d = max - min;
        s = l > 50 ? d / (2 - max - min) : d / (max + min);
        s = Math.round(s * 100);
        switch (max) {
          case r: h = (g - b) / d + (g < b ? 6 : 0); break;
          case g: h = (b - r) / d + 2; break;
          case b: h = (r - g) / d + 4; break;
        }
        h /= 6;
      }
      h = Math.round(h * 360);
      return [h, s, l];
    }
    function hslToRgb(h, s, l) {
      let r, g, b;
      h /= 360;
      s /= 100;
      l /= 100;
      if (s == 0) {
        r = g = b = l;
      } else {
        function hue2rgb(p, q, t) {
          if (t < 0) t += 1;
          if (t > 1) t -= 1;
          if (t < 1/6) return p + (q - p) * 6 * t;
          if (t < 1/2) return q;
          if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
          return p;
        }
        let q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        let p = 2 * l - q;
        r = hue2rgb(p, q, h + 1/3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1/3);
      }
      return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
    }
    function rgbToHex(r, g, b) {
      function componentToHex(c) {
        let hex = c.toString(16);
        return hex.length == 1 ? "0" + hex : hex;
      }
      return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
    }
    function linearInterpolateColor(hex1, hex2, t) {
      const rgb1 = hexToRgb(hex1);
      const rgb2 = hexToRgb(hex2);
      const r = Math.round(rgb1.r + (rgb2.r - rgb1.r) * t);
      const g = Math.round(rgb1.g + (rgb2.g - rgb1.g) * t);
      const b = Math.round(rgb1.b + (rgb2.b - rgb1.b) * t);
      return rgbToHex(r, g, b);
    }
    function generateGradientScale(gradientStops, numSteps) {
      const palette = [];
      if (numSteps <= 0) return [];
      if (numSteps === 1) return [gradientStops[0]];
      const numSegments = gradientStops.length - 1;
      for (let i = 0; i < numSteps; i++) {
        const t_global = (numSteps === 1) ? 0 : (i / (numSteps - 1));
        const segmentProgress = t_global * numSegments;
        const stopIndex1 = Math.floor(segmentProgress);
        const stopIndex2 = Math.min(stopIndex1 + 1, numSegments); 
        const t_local = segmentProgress - stopIndex1;
        const hex1 = gradientStops[stopIndex1];
        const hex2 = gradientStops[stopIndex2];
        palette.push(linearInterpolateColor(hex1, hex2, t_local));
      }
      return palette;
    }
    function generateMonochromaticScale(primaryHex, colorIds) {
      const numSteps = colorIds.length;
      if (numSteps === 0) return [];
      const rgb = hexToRgb(primaryHex);
      if (!rgb) return [];
      const [h, s, l_mid_original] = rgbToHsl(rgb.r, rgb.g, rgb.b);
      const gradientDarken = 7; 
      const L_MAX = 80; 
      const L_MIN = 45; 
      const l_step_amount = numSteps > 1 ? (L_MAX - L_MIN) / (numSteps - 1) : 0;
      const l_steps_calc = [];
      for (let i = 0; i < numSteps; i++) {
        const l_val = L_MAX - (i * l_step_amount);
        l_steps_calc.push(l_val);
      }
      const l_steps_light = [];
      const l_steps_dark = [];
      const mid_index = Math.ceil(numSteps / 2);
      for (let i = 0; i < mid_index; i++) {
        l_steps_light.push(l_steps_calc[i]);
      }
      for (let i = mid_index; i < numSteps; i++) {
        l_steps_dark.push(l_steps_calc[i]);
      }
      const l_steps_start = [];
      for (let i = 0; i < numSteps; i++) {
        if (i % 2 === 0) {
          l_steps_start.push(l_steps_light[i / 2]);
        } else {
          l_steps_start.push(l_steps_dark[(i - 1) / 2]);
        }
      }
      const colors = [];
      for (let i = 0; i < numSteps; i++) {
        const l_start_step = l_steps_start[i];
        let l_start = Math.max(0, Math.min(100, l_start_step));
        let l_end = Math.max(0, Math.min(100, l_start - gradientDarken));
        let s_adjusted = s;
        if (l_start > 85) s_adjusted = s * (1 - (l_start - 85) / 15);
        if (l_start < 20) s_adjusted = s * (l_start / 20);
        s_adjusted = Math.max(0, Math.min(100, s_adjusted));
        const rgb_start = hslToRgb(h, s_adjusted, l_start);
        const hex_start = rgbToHex(rgb_start[0], rgb_start[1], rgb_start[2]);
        const rgb_end = hslToRgb(h, s_adjusted, l_end);
        const hex_end = rgbToHex(rgb_end[0], rgb_end[1], rgb_end[2]);
        colors.push({
          id: colorIds[i],
          start: hex_start,
          end: hex_end
        });
      }
      return colors;
    }
    function buildChart(config) {
      const nl = String.fromCharCode(10);
  const animProgress = (config.animation == null) 
      ? 1.0 
      : Math.max(0, Math.min(1, Number(config.animation)));
  let animOpacity;
  const fadeStart = 0.7;
  if (animProgress <= fadeStart) {
      animOpacity = 0;
  } else {
      animOpacity = (animProgress - fadeStart) / (1.0 - fadeStart);
  }
         const defaultLayout = {
          width: 600,
          height: 510,
          marginTop: 150,
          marginBottom: 80,
          marginLeft: 70,
          marginRight: 25
      };
      if (config.layout && typeof config.layout === 'object') {
          config.layout = {
              width: (config.layout.width !== undefined) ? config.layout.width : defaultLayout.width,
              height: (config.layout.height !== undefined) ? config.layout.height : defaultLayout.height,
              marginTop: (config.layout.marginTop !== undefined) ? config.layout.marginTop : defaultLayout.marginTop,
              marginBottom: (config.layout.marginBottom !== undefined) ? config.layout.marginBottom : defaultLayout.marginBottom,
              marginLeft: (config.layout.marginLeft !== undefined) ? config.layout.marginLeft : defaultLayout.marginLeft,
              marginRight: (config.layout.marginRight !== undefined) ? config.layout.marginRight : defaultLayout.marginRight
          };
      } else {
          config.layout = defaultLayout;
      }
         const defaultBarOptions = {
               width: 50,
               cornerRadius: 4
         };
         if (config.barOptions && typeof config.barOptions === 'object') {
               config.barOptions = {
                     width: (config.barOptions.width !== undefined && typeof config.barOptions.width === 'number') ? config.barOptions.width : defaultBarOptions.width,
                     cornerRadius: (config.barOptions.cornerRadius !== undefined && typeof config.barOptions.cornerRadius === 'number') ? config.barOptions.cornerRadius : defaultBarOptions.cornerRadius
               };
         } else {
               config.barOptions = defaultBarOptions;
         }
      const chartConfig = config;
      const colorConfig = config.colors;
      const svgRoot = document.getElementById('100-stacked-bar-chart-svg');
      const defs = svgRoot.querySelector('defs');
      const style = svgRoot.querySelector('style');
      const existingGrads = defs.querySelectorAll('linearGradient');
      existingGrads.forEach(grad => grad.remove());
      const existingClips = defs.querySelectorAll('clipPath');
      existingClips.forEach(clip => clip.remove());
      let dynamicStyleParts = [];
      colorConfig.forEach(color => {
        const gradId = 'grad' + color.id;
        const cssClass = 'bar-' + color.id.toLowerCase();
        const gradient = createSVGElement('linearGradient', { id: gradId, x1:"0", x2:"0", y1:"0", y2:"1" });
        gradient.appendChild(createSVGElement('stop', { offset: '0%', 'stop-color': color.start }));
        gradient.appendChild(createSVGElement('stop', { offset: '100%', 'stop-color': color.end }));
        defs.appendChild(gradient);
        dynamicStyleParts.push('.' + cssClass + ' { fill: url(#' + gradId + '); }');
      });
      const dynamicStyles = dynamicStyleParts.join(nl);
      const styleElement = svgRoot.querySelector('style');
      const firstRuleSelector = '.bar-' + (config.colors[0] ? config.colors[0].id.toLowerCase() : '');
      if (styleElement && firstRuleSelector && !styleElement.textContent.includes(firstRuleSelector)) {
        styleElement.textContent += nl + dynamicStyles + nl;
      }
      const titleGroup = document.getElementById('title-group');
      const legendGroup = document.getElementById('legend-group');
      const yAxisGroup = document.getElementById('y-axis-group');
      const chartAreaGroup = document.getElementById('chart-area-group');
      const sourceGroup = document.getElementById('source-group');
      for(let group of [titleGroup, legendGroup, yAxisGroup, chartAreaGroup, sourceGroup]) {
        while (group.firstChild) {
          group.removeChild(group.firstChild);
        }
      }
      const chartAreaWidth = chartConfig.layout.width - chartConfig.layout.marginLeft - chartConfig.layout.marginRight;
      const chartAreaHeight = chartConfig.layout.height - chartConfig.layout.marginTop - chartConfig.layout.marginBottom;
      const bottomY = chartConfig.layout.marginTop + chartAreaHeight;
      const xScale = (index) => chartConfig.layout.marginLeft + (chartAreaWidth / chartConfig.barData.length) * (index + 0.5);
      const yScale = (percentage) => bottomY - (percentage / 100) * chartAreaHeight;
      titleGroup.appendChild(createSVGElement('text', { x: 300, y: 50, class: 'title' }, chartConfig.title));
      titleGroup.appendChild(createSVGElement('text', { x: 300, y: 75, class: 'subtitle' }, chartConfig.subtitle));
      sourceGroup.appendChild(createSVGElement('text', { x: 595, y: 485, class: 'source' }, chartConfig.source));
//
      const legendY = 105;
      const numLegends = chartConfig.legendLabels.length;
      let legendStartX = chartConfig.layout.marginLeft;
      const rectSize = 12;
      const rectTextGap = 6;
      const legendFontSize = 13;
      const legendPadding = 15;
      let maxTextWidth = 0;
      chartConfig.legendLabels.forEach(function(label) {
          const textWidth = getEstimatedTextWidth(label, legendFontSize);
          if (textWidth > maxTextWidth) {
              maxTextWidth = textWidth;
          }
      });
      const standardItemWidth = rectSize + rectTextGap + maxTextWidth + legendPadding;
      const totalRequiredWidth = standardItemWidth * numLegends;
      let useEqualWidths = false;
      let layoutItemWidth = 0;
      if (totalRequiredWidth > chartAreaWidth || numLegends === 0) {
          useEqualWidths = true;
          layoutItemWidth = (numLegends > 0) ? (chartAreaWidth / numLegends) : 0;
      } else {
          useEqualWidths = false;
          layoutItemWidth = standardItemWidth;
          legendStartX = chartConfig.layout.marginLeft + (chartAreaWidth - totalRequiredWidth) / 2;
      }
      chartConfig.legendLabels.forEach(function(label, i) {
        const cssClass = 'bar-' + colorConfig[i].id.toLowerCase();
        const columnStartX = legendStartX + (i * layoutItemWidth);
        const currentColumnWidth = layoutItemWidth;
        const itemGroup = createSVGElement('g');
        itemGroup.appendChild(createSVGElement('rect', {
          x: columnStartX, y: legendY, width: rectSize, height: rectSize, rx: 2, class: 'bar-segment ' + cssClass
        }));
        const textX = columnStartX + rectSize + rectTextGap;
        const clipId = 'legend-clip-' + i;
        let clipMargin = (legendPadding / 2);
        if (useEqualWidths) clipMargin = 5;
        let clipWidth = (currentColumnWidth - rectSize - rectTextGap - clipMargin); 
        clipWidth = Math.max(0, clipWidth);
        const clipPath = createSVGElement('clipPath', { id: clipId });
        clipPath.appendChild(createSVGElement('rect', {
            x: textX,
            y: legendY - (rectSize / 2),
            width: clipWidth,
            height: rectSize * 2
        }));
        defs.appendChild(clipPath);
        const textElement = createSVGElement('text', { 
            x: textX, 
            y: legendY + rectSize / 2, 
            class: 'legend-item',
            'clip-path': 'url(#' + clipId + ')' 
        }, label);
        itemGroup.appendChild(textElement);
         legendGroup.appendChild(itemGroup);
      });
//
      const tickCount = chartConfig.yAxis.tickCount > 0 ? chartConfig.yAxis.tickCount : 1;
      for (let i = 0; i <= tickCount; i++) {
        const tickPercentage = (100 / tickCount) * i;
        const y = yScale(tickPercentage);
        yAxisGroup.appendChild(createSVGElement('text', { x: 60, y: y, class: 'y-axis-value', 'dominant-baseline': 'middle' }, Math.round(tickPercentage) + '%'));
        if (i > 0 && i < tickCount) {
          yAxisGroup.appendChild(createSVGElement('line', { x1: chartConfig.layout.marginLeft, y1: y, x2: chartConfig.layout.width - chartConfig.layout.marginRight, y2: y, class: 'grid-line' }));
        }
      }
      yAxisGroup.appendChild(createSVGElement('line', { x1: chartConfig.layout.marginLeft, y1: bottomY, x2: chartConfig.layout.width - chartConfig.layout.marginRight, y2: bottomY, stroke: '#bdbdbd' }));
      yAxisGroup.appendChild(createSVGElement('line', { x1: chartConfig.layout.marginLeft, y1: config.layout.marginTop, x2: chartConfig.layout.width - chartConfig.layout.marginRight, y2: config.layout.marginTop, stroke: '#bdbdbd' }));
      const slotWidth = chartAreaWidth / chartConfig.barData.length;
      const minGap = 5;
      const maxBarWidth = chartConfig.barOptions.width;
      const calculatedWidth = slotWidth - minGap;
      const finalBarWidth = Math.max(1, Math.min(maxBarWidth, calculatedWidth));
      chartConfig.barData.forEach((bar, index) => {
        const barGroup = createSVGElement('g', { class: 'bar-group' });
        const centerX = xScale(index);
        const total = bar.values.reduce((sum, v) => sum + v, 0);
        barGroup.appendChild(createSVGElement('text', { x: centerX, y: bottomY + 20, class: 'axis-label' }, bar.label));
        let yOffsetPercentage = 0;
    const numItems = bar.values.length;
    const equalPercentage = 100 / numItems;
    const finalPercentages = bar.values.map(value => (total > 0 ? (value / total) * 100 : 0));
    const equalPercentages = bar.values.map(() => equalPercentage);
        let animatedYOffsetPercentage = 0;
        bar.values.slice().reverse().forEach((originalValue, i) => {
            const segmentIndex = bar.values.length - 1 - i;
            const eqPerc = equalPercentages[segmentIndex];
            const finPerc = finalPercentages[segmentIndex];
            const animatedPercentage = eqPerc + (finPerc - eqPerc) * animProgress;
            if (animatedPercentage <= 0.01) return;
            const yTop = yScale(animatedYOffsetPercentage + animatedPercentage);
            const yBottom = yScale(animatedYOffsetPercentage);
            const height = yBottom - yTop;
            if (height <= 0.5) return;
            const cssClass = 'bar-' + colorConfig[segmentIndex].id.toLowerCase();
            const r = chartConfig.barOptions.cornerRadius;
            const w = finalBarWidth;
            const x = centerX - w / 2;
            let segment;
            let currentFinalOffset = 0;
            for(let j = bar.values.length - 1; j > segmentIndex; j--) {
                currentFinalOffset += finalPercentages[j];
            }
            const isBottomSegment = (Math.abs(currentFinalOffset) < 0.01);
            const isTopSegment = (Math.abs((currentFinalOffset + finPerc) - 100) < 0.01);
            if (isBottomSegment && isTopSegment) {
                segment = createSVGElement('rect', { x: x, y: yTop, width: w, height: height, rx: r, ry: r, class: 'bar-segment ' + cssClass });
            } else if (isBottomSegment) {
                const d = [
                    'M ' + x + ' ' + yTop, 'H ' + (x+w), 'V ' + (yBottom-r),
                    'A ' + r + ' ' + r + ' 0 0 1 ' + (x+w-r) + ' ' + yBottom, 'H ' + (x+r),
                    'A ' + r + ' ' + r + ' 0 0 1 ' + x + ' ' + (yBottom-r), 'Z'
                ].join(' ');
                segment = createSVGElement('path', { d: d, class: 'bar-segment ' + cssClass });
            } else if (isTopSegment) {
                const d = [
                    'M ' + x + ' ' + yBottom, 'H ' + (x+w), 'V ' + (yTop+r),
                    'A ' + r + ' ' + r + ' 0 0 0 ' + (x+w-r) + ' ' + yTop, 'H ' + (x+r),
                    'A ' + r + ' ' + r + ' 0 0 0 ' + x + ' ' + (yTop+r), 'Z'
                ].join(' ');
                segment = createSVGElement('path', { d: d, class: 'bar-segment ' + cssClass });
            } else {
                segment = createSVGElement('rect', { x: x, y: yTop, width: w, height: height, class: 'bar-segment ' + cssClass });
            }
            barGroup.appendChild(segment);
            const finalPercentageValue = finPerc;
            if (finalPercentageValue > 5) {
                barGroup.appendChild(createSVGElement('text', { 
                    x: centerX, 
                    y: yTop + height / 2,
                    'dominant-baseline': 'middle', 
                    class: 'segment-label',
                    style: 'opacity: ' + animOpacity
                }, Math.round(finalPercentageValue) + '%'));
            }
            animatedYOffsetPercentage += animatedPercentage;
        });
        chartAreaGroup.appendChild(barGroup);
      });
    }
    try {
      const configJson = document.getElementById('chart-json-data').textContent;
      if (!configJson || configJson.trim() === '') {
        throw new Error("JSON data is empty.");
      }
      let fullConfig;
      try {
        fullConfig = JSON.parse(configJson);
      } catch (parseError) {
        throw new Error('JSON parsing error: ' + parseError.message);
      }
      if (!fullConfig || typeof fullConfig !== 'object' || !fullConfig.data) {
        throw new Error("Invalid JSON structure: Missing 'data' key.");
      }
      const config = fullConfig.data;
      if (config.colors && typeof config.colors === 'string') {
        const numColors = config.legendLabels ? config.legendLabels.length : 5;
        const generatedIds = [];
        for (let i = 0; i < numColors; i++) {
          generatedIds.push(String.fromCharCode(65 + i));
        }
        let generatedColors;
        const colorKeyword = config.colors.toLowerCase();
        const gradientDarken = 15;
        if (colorKeyword === '#gemini') {
          const GEMINI_STOPS = ['#E68A9C', '#9F63D0', '#616AD8'];
          const hexPalette = generateGradientScale(GEMINI_STOPS, numColors);
          generatedColors = [];
          for (let i = 0; i < numColors; i++) {
            const hex_start = hexPalette[i];
            const rgb_start = hexToRgb(hex_start);
            const [h, s, l] = rgbToHsl(rgb_start.r, rgb_start.g, rgb_start.b);
            const l_end = Math.max(0, Math.min(100, l - gradientDarken));
            const rgb_end = hslToRgb(h, s, l_end);
            const hex_end = rgbToHex(rgb_end[0], rgb_end[1], rgb_end[2]);
            generatedColors.push({
              id: generatedIds[i],
              start: hex_start,
              end: hex_end
            });
          }
        } else if (colorKeyword === '#night') {
const NIGHT_COLORS = [
          { start: '#8d58d3', end: '#704ca4' },
          { start: '#8f83d6', end: '#7469a8' },
          { start: '#7891dd', end: '#5c6fbb' },
          { start: '#50a89d', end: '#258a7f' },
          { start: '#54c5d5', end: '#34a0b0' }
        ];
          generatedColors = [];
          for (let i = 0; i < numColors; i++) {
            const colorPair = NIGHT_COLORS[i] || NIGHT_COLORS[NIGHT_COLORS.length - 1];
            generatedColors.push({
              id: generatedIds[i],
              start: colorPair.start,
              end: colorPair.end
            });
          }
        } else {
          const primaryColor = config.colors;
          generatedColors = generateMonochromaticScale(primaryColor, generatedIds);
          if (generatedColors.length === 0) {
             throw new Error("Invalid HEX color string in 'colors': " + primaryColor);
          }
        }
        if (generatedColors && generatedColors.length > 0) {
          config.colors = generatedColors;
        } else {
           throw new Error("Failed to generate colors from keyword: " + config.colors);
        }
      } else if (!config.colors || !Array.isArray(config.colors)) {
         throw new Error("'colors' must be an array of color objects or a single primary color string.");
      }
      if (!config || typeof config !== 'object' || !config.barData || !Array.isArray(config.barData)) {
        throw new Error("Invalid JSON structure: Missing required fields like 'barData' inside 'data'.");
      }
      buildChart(config);
    } catch(e) {
      console.error("Chart Error:", e);
      const svgRoot = document.getElementById('100-stacked-bar-chart-svg');
      const existingError = svgRoot.querySelector('.error-message');
      if(existingError) existingError.remove();
      const textEl = createSVGElement('text', {
        x: svgRoot.viewBox.baseVal.width / 2 + svgRoot.viewBox.baseVal.x,
        y: svgRoot.viewBox.baseVal.height / 2 + svgRoot.viewBox.baseVal.y,
        'text-anchor': 'middle', fill: 'red',
        'font-family': "'Noto Sans JP', sans-serif",
        class: 'error-message'
      });
      textEl.textContent = e.message.includes('JSON') ? 'JSONデータにエラーがあります。' : 'グラフ描画エラーが発生しました。';
      if (e.message.includes('JSON parsing error')) {
        textEl.textContent += ' (' + e.message.split(': ')[1] + ')';
      }
      svgRoot.appendChild(textEl);
    }
  </script>
</svg>
`;
const LINE_CHART_TEMPLATE = `
<svg width="600" height="465" viewBox="25 25 575 430" xmlns="http://www.w3.org/2000/svg" font-family="'Noto Sans JP', sans-serif" id="line-chart-svg">
  <script id="chart-json-data" type="application/json">
    {
      "chartType": "line",
      "data": {
        "title": "年間データ推移",
        "subtitle": "（サンプルデータ）",
        "source": "出典：サンプルデータ",
        "yAxisUnitLabel": "（単位）",
        "items": [
          { "label": "1月", "value": 8 },
          { "label": "2月", "value": 9 },
          { "label": "3月", "value": 12 },
          { "label": "4月", "value": 16 },
          { "label": "5月", "value": 20 },
          { "label": "6月", "value": 24 },
          { "label": "7月", "value": 28 },
          { "label": "8月", "value": 27 },
          { "label": "9月", "value": 35 }
        ],
        "colors": {
          "start": "#e68a9c",
          "end": "#b469b8",
          "line": "#b469b8",
          "label": "#8c4fc8"
        },
        "layout": {
          "width": 600,
          "height": 465,
          "marginTop": 110,
          "marginBottom": 75,
          "marginLeft": 75,
          "marginRight": 25
        },
        "yAxis": {
          "max": 40,
          "min": 0,
          "tickCount": 4
        },
        "lineOptions": {
          "markerRadius": 5,
          "dataLabelOffsetY": 15,
          "horizontalPadding": 30
        },
        "animation": 1
      }
    }
</script>
  <defs>
    <filter id="shadow">
      <feDropShadow dx="1" dy="2" stdDeviation="2" flood-color="#000" flood-opacity="0.2"/>
    </filter>
    <filter id="text-halo">
      <feDropShadow dx="0" dy="0" stdDeviation="3" flood-color="white" flood-opacity="0.9"/>
    </filter>
  </defs>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&amp;family=Roboto:wght@400;700&amp;display=swap');
    .title { font-size: 24px; font-weight: 500; text-anchor: middle; fill: #5f6368; }
    .subtitle { font-size: 16px; text-anchor: middle; fill: #5f6368; }
    .grid-line { stroke: #e0e0e0; stroke-dasharray: 2 2; stroke-width: 0.8; }
    .axis-label { font-size: 13px; fill: #5f6368; text-anchor: middle; }
    .source { font-size: 11px; fill: #70757a; text-anchor: end; }
    .y-axis-value { font-family: 'Roboto', sans-serif; font-size: 12px; fill: #757575; text-anchor: end; }
  </style>
  <g id="title-group"></g>
  <g id="y-axis-group"></g>
  <g id="chart-area-group"></g>
  <g id="x-axis-group"></g>
  <g id="source-group"></g>
  <script type="text/javascript">
    const svgNS = "http://www.w3.org/2000/svg";
    function createSVGElement(name, attributes, textContent) {
      const el = document.createElementNS(svgNS, name);
      for (const key in attributes) el.setAttribute(key, attributes[key]);
      if (textContent !== undefined && textContent !== null) el.textContent = textContent;
      return el;
    }
    function hexToRgb(hex) {
      let shorthandRegex = /^#?([a-f0-9])([a-f0-9])([a-f0-9])$/i;
      hex = hex.replace(shorthandRegex, function(m, r, g, b) {
        return r + r + g + g + b + b;
      });
      let result = /^#?([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})$/i.exec(hex);
      return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
      } : null;
    }
    function rgbToHsl(r, g, b) {
      r /= 255; g /= 255; b /= 255;
      let max = Math.max(r, g, b), min = Math.min(r, g, b);
      let h, s, l = (max + min) / 2;
      l = Math.round(l * 100);
      if (max == min) { h = s = 0; } else {
        let d = max - min;
        s = l > 50 ? d / (2 - max - min) : d / (max + min);
        s = Math.round(s * 100);
        switch (max) {
          case r: h = (g - b) / d + (g < b ? 6 : 0); break;
          case g: h = (b - r) / d + 2; break;
          case b: h = (r - g) / d + 4; break;
        }
        h /= 6;
      }
      h = Math.round(h * 360);
      return [h, s, l];
    }
    function hslToRgb(h, s, l) {
      let r, g, b;
      h /= 360; s /= 100; l /= 100;
      if (s == 0) { r = g = b = l; } else {
        function hue2rgb(p, q, t) {
          if (t < 0) t += 1;
          if (t > 1) t -= 1;
          if (t < 1/6) return p + (q - p) * 6 * t;
          if (t < 1/2) return q;
          if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
          return p;
        }
        let q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        let p = 2 * l - q;
        r = hue2rgb(p, q, h + 1/3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1/3);
      }
      return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
    }
    function rgbToHex(r, g, b) {
      function componentToHex(c) {
        let hex = c.toString(16);
        return hex.length == 1 ? "0" + hex : hex;
      }
      return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
    }
    function truncateText(text, maxWidth, className) {
      const svgRoot = document.getElementById('line-chart-svg');
      const tempText = createSVGElement('text', { class: className, style: 'visibility: hidden;' });
      svgRoot.appendChild(tempText);
      tempText.textContent = text;
      if (tempText.getComputedTextLength() <= maxWidth) {
        svgRoot.removeChild(tempText);
        return text;
      }
      let truncatedText = text;
      while (truncatedText.length > 0) {
        truncatedText = truncatedText.slice(0, -1);
        tempText.textContent = truncatedText + '...';
        if (tempText.getComputedTextLength() <= maxWidth) {
          svgRoot.removeChild(tempText);
          return truncatedText + '...';
        }
      }
      svgRoot.removeChild(tempText);
      return '...';
    }
    function buildChart(config) {
      const svgRoot = document.getElementById('line-chart-svg');
      const defs = svgRoot.querySelector('defs');
      const style = svgRoot.querySelector('style');
      const existingGrads = defs.querySelectorAll('linearGradient');
      existingGrads.forEach(grad => grad.remove());
      let styleContent = style.textContent;
      styleContent = styleContent.replace(/[.]line {[^}]*}/g, '');
      styleContent = styleContent.replace(/[.]marker {[^}]*}/g, '');
      styleContent = styleContent.replace(/[.]area {[^}]*}/g, '');
      styleContent = styleContent.replace(/[.]data-label {[^}]*}/g, '');
      style.textContent = styleContent.trim();
      const markerGradId = 'markerGradient';
      let markerGrad = createSVGElement('linearGradient', { id: markerGradId, x1:"0", x2:"0", y1:"0", y2:"1" });
      markerGrad.appendChild(createSVGElement('stop', { offset: '0%', 'stop-color': config.color.start }));
      markerGrad.appendChild(createSVGElement('stop', { offset: '100%', 'stop-color': config.color.end }));
      defs.appendChild(markerGrad);
      const areaGradId = 'areaGradient';
      let areaGrad = createSVGElement('linearGradient', { id: areaGradId, x1:"0", x2:"0", y1:"0", y2:"1" });
      areaGrad.appendChild(createSVGElement('stop', { offset: '0%', 'stop-color': config.color.start, 'stop-opacity': '0.2' }));
      areaGrad.appendChild(createSVGElement('stop', { offset: '100%', 'stop-color': config.color.end, 'stop-opacity': '0' }));
      defs.appendChild(areaGrad);
      const nl = String.fromCharCode(10);
      let dynamicStyles = [
          "",
          "            .line { fill: none; stroke: " + config.color.line + "; stroke-width: 3; stroke-linejoin: round; stroke-linecap: round; filter: url(#shadow);",
          "            }",
          "            .marker { fill: url(#" + markerGradId + "); stroke: white; stroke-width: 2; filter: url(#shadow); }",
          "            .area { fill: url(#" + areaGradId + ");",
          "            }",
          "            .data-label { font-family: 'Roboto', sans-serif; font-size: 14px; font-weight: bold; fill: " + config.color.label + "; text-anchor: middle; filter: url(#text-halo); }",
          "        "
      ].join(nl);
      style.textContent += dynamicStyles;
      const groups = {
        title: document.getElementById('title-group'),
        yAxis: document.getElementById('y-axis-group'),
        chartArea: document.getElementById('chart-area-group'),
        xAxis: document.getElementById('x-axis-group'),
        source: document.getElementById('source-group')
      };
      for (const key in groups) {
        while (groups[key].firstChild) groups[key].removeChild(groups[key].firstChild);
      }
      const chartAreaWidth = config.layout.width - config.layout.marginLeft - config.layout.marginRight;
      const chartAreaHeight = config.layout.height - config.layout.marginTop - config.layout.marginBottom;
      const effectivePlotWidth = chartAreaWidth - (config.lineOptions.horizontalPadding * 2);
      const dataMax = Math.max.apply(null, config.items.map(function(item) { return item.value; }).concat([config.yAxis.max]));
      const dataMin = Math.min.apply(null, config.items.map(function(item) { return item.value; }).concat([config.yAxis.min]));
      const hasDataDecimals = config.items.some(function(item) {
        return typeof item.value === 'number' && item.value % 1 !== 0;
      });
      const hasAxisDecimals = (typeof config.yAxis.min === 'number' && config.yAxis.min % 1 !== 0) ||
        (typeof config.yAxis.max === 'number' && config.yAxis.max % 1 !== 0);
      const showDecimals = hasDataDecimals || hasAxisDecimals;
      let effectiveYAxisMax, effectiveYAxisMin;
      const tickCount = config.yAxis.tickCount > 0 ? config.yAxis.tickCount : 1;
      const preliminaryRange = Math.max(0.1, config.yAxis.max - config.yAxis.min);
      let tickInterval = preliminaryRange / tickCount;
      if (dataMin < config.yAxis.min) {
        const limit = Math.max(Math.abs(dataMin), dataMax);
        if (tickInterval <= 0 || isNaN(tickInterval)) tickInterval = Math.max(1, limit / 5);
        effectiveYAxisMax = Math.ceil(limit / tickInterval) * tickInterval;
        effectiveYAxisMin = -effectiveYAxisMax;
        if (config.yAxis.min > effectiveYAxisMin) effectiveYAxisMin = config.yAxis.min;
        if (config.yAxis.max < effectiveYAxisMax) effectiveYAxisMax = config.yAxis.max;
      } else {
        effectiveYAxisMin = config.yAxis.min;
        if (tickInterval <= 0 || isNaN(tickInterval)) tickInterval = Math.max(1, (dataMax - effectiveYAxisMin) / 5) || 1;
        effectiveYAxisMax = Math.max(config.yAxis.max, Math.ceil(dataMax / tickInterval) * tickInterval);
      }
      if (effectiveYAxisMax <= effectiveYAxisMin) {
        effectiveYAxisMax = effectiveYAxisMin + tickInterval * tickCount;
        if (effectiveYAxisMax <= effectiveYAxisMin) effectiveYAxisMax = effectiveYAxisMin + 1;
      }
      const yRange = effectiveYAxisMax - effectiveYAxisMin;
      const xScale = function(index) { return (config.layout.marginLeft + config.lineOptions.horizontalPadding) + (config.items.length > 1 ? (index * (effectivePlotWidth / (config.items.length - 1))) : effectivePlotWidth / 2 ); };
      const yScale = function(value) { return (yRange === 0) ? config.layout.marginTop + chartAreaHeight / 2 : config.layout.marginTop + chartAreaHeight * (1 - (value - effectiveYAxisMin) / yRange); };
      const bottomY = config.layout.marginTop + chartAreaHeight;
      const animationProgress = (config.animation == null) 
          ? 1.0 
          : Math.max(0, Math.min(1, Number(config.animation)));
      let animationOpacity;
      const fadeStart = 0.7;
      if (animationProgress <= fadeStart) {
          animationOpacity = 0;
      } else {
          animationOpacity = (animationProgress - fadeStart) / (1.0 - fadeStart);
      }
      groups.title.appendChild(createSVGElement('text', { x: 312, y: 50, class: 'title' }, config.title));
      groups.title.appendChild(createSVGElement('text', { x: 312, y: 75, class: 'subtitle' }, config.subtitle));
      groups.source.appendChild(createSVGElement('text', { x: 600, y: 450, class: 'source' }, config.source));
      groups.yAxis.appendChild(createSVGElement('text', { x: config.layout.marginLeft, y: config.layout.marginTop - 15, class: 'y-axis-value', 'text-anchor': 'end' }, config.yAxisUnitLabel));
      const finalTickStep = yRange / tickCount;
      for (let i = 0; i <= tickCount; i++) {
        const tickValue = effectiveYAxisMin + (finalTickStep * i);
        const y = yScale(tickValue);
        let tickLabel;
        if (showDecimals) {
          tickLabel = parseFloat(tickValue.toFixed(5));
        } else {
          tickLabel = Math.round(tickValue);
        }
        groups.yAxis.appendChild(createSVGElement('text', { x: config.layout.marginLeft - 10, y, class: 'y-axis-value', 'dominant-baseline': 'middle' }, tickLabel));
        if (i > 0 && (Math.abs(tickValue) > 1e-6 || effectiveYAxisMin >= 0)) {
          groups.yAxis.appendChild(createSVGElement('line', { x1: config.layout.marginLeft, y1: y, x2: config.layout.width - config.layout.marginRight, y2: y, class: 'grid-line' }));
        }
      }
      if (effectiveYAxisMin < 0 && effectiveYAxisMax > 0) {
        groups.yAxis.appendChild(createSVGElement('line', { x1: config.layout.marginLeft, y1: yScale(0), x2: config.layout.width - config.layout.marginRight, y2: yScale(0), stroke: '#bdbdbd', 'stroke-width': 1 }));
      }
      groups.xAxis.appendChild(createSVGElement('line', { x1: config.layout.marginLeft, y1: bottomY, x2: config.layout.width - config.layout.marginRight, y2: bottomY, stroke: '#bdbdbd', 'stroke-width': 1 }));
      const xSlotWidth = config.items.length > 1 ? effectivePlotWidth / (config.items.length - 1) : effectivePlotWidth;
      config.items.forEach(function(item, index) {
        const x = xScale(index);
        const labelContent = truncateText(item.label, xSlotWidth * 0.9, 'axis-label');
        groups.xAxis.appendChild(createSVGElement('text', { x: x, y: bottomY + 20, class: 'axis-label' }, labelContent));
      });
      if (config.items.length > 0) {
        const areaBottomY = yScale(Math.max(0, effectiveYAxisMin));
        const pointsData = config.items.map(function(item, index) {
            const x = xScale(index);
            const finalY = yScale(item.value);
            const animatedY = areaBottomY + (finalY - areaBottomY) * animationProgress;
            return {
                x: x,
                finalY: finalY,
                animatedY: animatedY,
                value: item.value
            };
        });
        const points = pointsData.map(function(d) { return d.x + ',' + d.animatedY; });
        const areaPoints = xScale(0) + ',' + areaBottomY + ' ' + points.join(' ') + ' ' + xScale(config.items.length - 1) + ',' + areaBottomY;
        groups.chartArea.appendChild(createSVGElement('polygon', { class: 'area', points: areaPoints }));
        groups.chartArea.appendChild(createSVGElement('polyline', { class: 'line', points: points.join(' ') }));
        const dataPointsGroup = createSVGElement('g', { class: 'data-points' });
        pointsData.forEach(function(data, index) {
          const x = data.x;
          const y = data.animatedY;
          const finalY = data.finalY; 
          dataPointsGroup.appendChild(createSVGElement('circle', { 
              cx: x, 
              cy: y,
              r: config.lineOptions.markerRadius, 
              class: 'marker' 
          }));
          let labelY;
  if (y < config.layout.marginTop + 20) { 
            labelY = y + config.lineOptions.dataLabelOffsetY + 5; 
          } else {
            labelY = y - config.lineOptions.dataLabelOffsetY;
  }
          let dataLabelValue = data.value;
          if (!showDecimals) {
            dataLabelValue = Math.round(dataLabelValue);
          }
          dataPointsGroup.appendChild(createSVGElement('text', { 
              x: x, 
              y: labelY,
              class: 'data-label',
              style: 'opacity: ' + animationOpacity
          }, dataLabelValue));
        });
        groups.chartArea.appendChild(dataPointsGroup);
      }
    }
    try {
      const configJson = document.getElementById('chart-json-data').textContent;
      if (!configJson || configJson.trim() === '') {
        throw new Error("JSON data is empty.");
      }
      let fullConfig;
      try {
        fullConfig = JSON.parse(configJson);
      } catch (parseError) {
        throw new Error('JSON parsing error: ' + parseError.message);
      }
      if (!fullConfig || typeof fullConfig !== 'object' || !fullConfig.data) {
        throw new Error("Invalid JSON structure: Missing 'data' key.");
      }
      const config = fullConfig.data;
      const defaultLineOptions = {
          markerRadius: 5,
          dataLabelOffsetY: 15,
          horizontalPadding: 30
      };
      if (config.lineOptions && typeof config.lineOptions === 'object') {
          config.lineOptions = {
              markerRadius: (config.lineOptions.markerRadius !== undefined && typeof config.lineOptions.markerRadius === 'number') ?
                config.lineOptions.markerRadius : defaultLineOptions.markerRadius,
              dataLabelOffsetY: (config.lineOptions.dataLabelOffsetY !== undefined && typeof config.lineOptions.dataLabelOffsetY === 'number') ?
                config.lineOptions.dataLabelOffsetY : defaultLineOptions.dataLabelOffsetY,
              horizontalPadding: (config.lineOptions.horizontalPadding !== undefined && typeof config.lineOptions.horizontalPadding === 'number') ?
                config.lineOptions.horizontalPadding : defaultLineOptions.horizontalPadding
          };
      } else {
          config.lineOptions = defaultLineOptions;
      }
      if (config.colors) {
          config.color = config.colors;
      }
      if (config.color && typeof config.color === 'string') {
        const colorKeyword = config.color.toLowerCase();
        let generatedColorObject;
        if (colorKeyword === '#gemini') {
          const GEMINI_STOPS = ['#E68A9C', '#9F63D0', '#616AD8'];
          generatedColorObject = {
            start: GEMINI_STOPS[0], line:  GEMINI_STOPS[1],
            end:   GEMINI_STOPS[1], label: GEMINI_STOPS[2]
          };
        } else if (colorKeyword === '#night') {
          const NIGHT_COLORS = [
            { start: '#8d58d3', end: '#704ca4' },
            { start: '#8f83d6', end: '#7469a8' },
            { start: '#7891dd', end: '#5c6fbb' }
          ];
          generatedColorObject = {
            start: NIGHT_COLORS[0].start, end:   NIGHT_COLORS[1].start,
            line:  NIGHT_COLORS[1].end,   label: NIGHT_COLORS[2].end
          };
        } else {
          const primaryColor = config.color;
          const rgb = hexToRgb(primaryColor);
          if (!rgb) { throw new Error("Invalid HEX color string in 'color': " + primaryColor); }
          const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
          const h = hsl[0], s = hsl[1], l = hsl[2];
          const l_start = Math.min(100, l + 20);
          const s_start = (l > 80) ? s * 0.7 : s;
          const rgb_start = hslToRgb(h, s_start, l_start);
          const hex_start = rgbToHex(rgb_start[0], rgb_start[1], rgb_start[2]);
          const l_label = Math.max(0, l - 15);
          const s_label = (l < 20) ? s * 0.7 : s;
          const rgb_label = hslToRgb(h, s_label, l_label);
          const hex_label = rgbToHex(rgb_label[0], rgb_label[1], rgb_label[2]);
          generatedColorObject = {
            start: hex_start, end:   primaryColor,
            line:  primaryColor, label: hex_label
          };
        }
        config.color = generatedColorObject;
      } else if (!config.color || typeof config.color !== 'object') {
         throw new Error("'color' must be a color object or a single color string.");
      }
      if (!config || typeof config !== 'object' || !config.items || !Array.isArray(config.items)) {
        throw new Error("Invalid JSON structure: Missing required fields like 'items' inside 'data'.");
      }
      buildChart(config);
    } catch (e) {
      console.error("Chart Error:", e);
      const svgRoot = document.getElementById('line-chart-svg');
      const existingError = svgRoot.querySelector('.error-message');
      if(existingError) existingError.remove();
      const textEl = createSVGElement('text', {
        x: svgRoot.viewBox.baseVal.width / 2 + svgRoot.viewBox.baseVal.x,
        y: svgRoot.viewBox.baseVal.height / 2 + svgRoot.viewBox.baseVal.y,
        'text-anchor': 'middle', fill: 'red',
        'font-family': "'Noto Sans JP', sans-serif",
        class: 'error-message'
      });
      textEl.textContent = e.message.includes('JSON') ? 'JSONデータにエラーがあります。' : 'グラフ描画エラーが発生しました。';
      if (e.message.includes('JSON parsing error')) {
        textEl.textContent += ' (' + e.message.split(': ')[1] + ')';
      }
      svgRoot.appendChild(textEl);
    }
  </script>
</svg>
`;
const DONUT_CHART_TEMPLATE = `
<svg width="600" height="460" viewBox="0 25 600 425" xmlns="http://www.w3.org/2000/svg" font-family="'Noto Sans JP', sans-serif" id="donut-chart-svg">
  <script id="chart-json-data" type="application/json">
    {
      "chartType": "donut",
      "data": {
        "title": "グラフタイトル",
        "subtitle": "グラフサブタイトル",
        "source": "出典: データソース",
        "centerLabel": "合計ラベル",
        "colors": [
          { "id": "A", "start": "#e68a9c", "end": "#d96d8f" },
          { "id": "B", "start": "#b469b8", "end": "#a656ad" },
          { "id": "C", "start": "#9f63d0", "end": "#8c4fc8" },
          { "id": "D", "start": "#7c6ce8", "end": "#6b5ce0" },
          { "id": "E", "start": "#616ad8", "end": "#5059d1" }
        ],
        "items": [
          { "label": "項目 A", "value": 40, "id": "A" },
          { "label": "項目 B", "value": 25, "id": "B" },
          { "label": "項目 C", "value": 15, "id": "C" },
          { "label": "項目 D", "value": 10, "id": "D" },
          { "label": "項目 E", "value": 10, "id": "E" }
        ],
        "animation" : 1
      }
    }
  </script>
  <defs>
    <filter id="shadow"><feDropShadow dx="1" dy="2" stdDeviation="2" flood-color="#000" flood-opacity="0.4"/></filter>
  </defs>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&amp;family=Roboto:wght@400;700&amp;display=swap');
    .title { font-size: 24px; font-weight: 500; text-anchor: middle; fill: #5f6368; }
    .subtitle { font-size: 16px; text-anchor: middle; fill: #5f6368; }
    .source { font-size: 11px; fill: #70757a; text-anchor: end; }
    .center-total-label { font-size: 14px; fill: #5f6368; text-anchor: middle; dominant-baseline: central; }
    .legend-item { font-size: 13px; fill: #5f6368; dominant-baseline: middle; }
    .donut-segment { filter: url(#shadow); }
    .center-total-value { font-family: 'Roboto', sans-serif; font-size: 28px; font-weight: bold; fill: #5f6368; text-anchor: middle; dominant-baseline: central; }
    .legend-percentage { font-family: 'Roboto', sans-serif; font-weight: 400; }
  </style>
  <g id="title-group"></g>
  <g id="donut-group"></g>
  <g id="legend-group"></g>
  <g id="source-group"></g>
  <script type="text/javascript">
    const svgNS = "http://www.w3.org/2000/svg";
    function createSVGElement(name, attributes, textContent) {
      const el = document.createElementNS(svgNS, name);
      for (const key in attributes) el.setAttribute(key, attributes[key]);
      if (textContent !== undefined && textContent !== null) el.textContent = textContent;
      return el;
    }
    function polarToCartesian(centerX, centerY, radius, angleInDegrees) {
      const angleInRadians = (angleInDegrees - 90) * Math.PI / 180.0;
      return {
        x: centerX + (radius * Math.cos(angleInRadians)),
        y: centerY + (radius * Math.sin(angleInRadians))
      };
    }
    function createDonutSegmentPath(x, y, radius, startAngle, endAngle) {
      if (endAngle - startAngle >= 360) endAngle = 359.99;
      const start = polarToCartesian(x, y, radius, endAngle);
      const end = polarToCartesian(x, y, radius, startAngle);
      const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
      return 'M ' + x + ' ' + y + ' L ' + start.x + ' ' + start.y + ' A ' + radius + ' ' + radius + ' 0 ' + largeArcFlag + ' 0 ' + end.x + ' ' + end.y + ' Z';
    }
    function hexToRgb(hex) {
      let shorthandRegex = /^#?([a-f0-9])([a-f0-9])([a-f0-9])$/i;
      hex = hex.replace(shorthandRegex, function(m, r, g, b) {
        return r + r + g + g + b + b;
      });
      let result = /^#?([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})$/i.exec(hex);
      return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
      } : null;
    }
    function rgbToHsl(r, g, b) {
      r /= 255; g /= 255; b /= 255;
      let max = Math.max(r, g, b), min = Math.min(r, g, b);
      let h, s, l = (max + min) / 2;
      l = Math.round(l * 100);
      if (max == min) {
        h = s = 0;
      } else {
        let d = max - min;
        s = l > 50 ? d / (2 - max - min) : d / (max + min);
        s = Math.round(s * 100);
        switch (max) {
          case r: h = (g - b) / d + (g < b ? 6 : 0); break;
          case g: h = (b - r) / d + 2; break;
          case b: h = (r - g) / d + 4; break;
        }
        h /= 6;
      }
      h = Math.round(h * 360);
      return [h, s, l];
    }
    function hslToRgb(h, s, l) {
      let r, g, b;
      h /= 360;
      s /= 100;
      l /= 100;
      if (s == 0) {
        r = g = b = l;
      } else {
        function hue2rgb(p, q, t) {
          if (t < 0) t += 1;
          if (t > 1) t -= 1;
          if (t < 1/6) return p + (q - p) * 6 * t;
          if (t < 1/2) return q;
          if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
          return p;
        }
        let q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        let p = 2 * l - q;
        r = hue2rgb(p, q, h + 1/3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1/3);
      }
      return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
    }
    function rgbToHex(r, g, b) {
      function componentToHex(c) {
        let hex = c.toString(16);
        return hex.length == 1 ? "0" + hex : hex;
      }
      return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
    }
    function linearInterpolateColor(hex1, hex2, t) {
      const rgb1 = hexToRgb(hex1);
      const rgb2 = hexToRgb(hex2);
      const r = Math.round(rgb1.r + (rgb2.r - rgb1.r) * t);
      const g = Math.round(rgb1.g + (rgb2.g - rgb1.g) * t);
      const b = Math.round(rgb1.b + (rgb2.b - rgb1.b) * t);
      return rgbToHex(r, g, b);
    }
    function generateGradientScale(gradientStops, numSteps) {
      const palette = [];
      if (numSteps <= 0) return [];
      if (numSteps === 1) return [gradientStops[0]];
      const numSegments = gradientStops.length - 1;
      for (let i = 0; i < numSteps; i++) {
        const t_global = (numSteps === 1) ? 0 : (i / (numSteps - 1));
        const segmentProgress = t_global * numSegments;
        const stopIndex1 = Math.floor(segmentProgress);
        const stopIndex2 = Math.min(stopIndex1 + 1, numSegments); 
        const t_local = segmentProgress - stopIndex1;
        const hex1 = gradientStops[stopIndex1];
        const hex2 = gradientStops[stopIndex2];
        palette.push(linearInterpolateColor(hex1, hex2, t_local));
      }
      return palette;
    }
    function generateMonochromaticScale(primaryHex, colorIds) {
      const numSteps = colorIds.length;
      if (numSteps === 0) return [];
      const rgb = hexToRgb(primaryHex);
      if (!rgb) return [];
      const [h, s, l_mid_original] = rgbToHsl(rgb.r, rgb.g, rgb.b);
      const gradientDarken = 7;
      const L_MAX = 80; 
      const L_MIN = 45; 
      const l_step_amount = numSteps > 1 ? (L_MAX - L_MIN) / (numSteps - 1) : 0;
      const l_steps_calc = [];
      for (let i = 0; i < numSteps; i++) {
        const l_val = L_MAX - (i * l_step_amount);
        l_steps_calc.push(l_val);
      }
      const l_steps_light = [];
      const l_steps_dark = [];
      const mid_index = Math.ceil(numSteps / 2);
      for (let i = 0; i < mid_index; i++) {
        l_steps_light.push(l_steps_calc[i]);
      }
      for (let i = mid_index; i < numSteps; i++) {
        l_steps_dark.push(l_steps_calc[i]);
      }
      const l_steps_start = [];
      for (let i = 0; i < numSteps; i++) {
        if (i % 2 === 0) {
          l_steps_start.push(l_steps_light[i / 2]);
        } else {
          l_steps_start.push(l_steps_dark[(i - 1) / 2]);
        }
      }
      const colors = [];
      for (let i = 0; i < numSteps; i++) {
        const l_start_step = l_steps_start[i];
        let l_start = Math.max(0, Math.min(100, l_start_step));
        let l_end = Math.max(0, Math.min(100, l_start - gradientDarken));
        let s_adjusted = s;
        if (l_start > 85) s_adjusted = s * (1 - (l_start - 85) / 15);
        if (l_start < 20) s_adjusted = s * (l_start / 20);
        s_adjusted = Math.max(0, Math.min(100, s_adjusted));
        const rgb_start = hslToRgb(h, s_adjusted, l_start);
        const hex_start = rgbToHex(rgb_start[0], rgb_start[1], rgb_start[2]);
        const rgb_end = hslToRgb(h, s_adjusted, l_end);
        const hex_end = rgbToHex(rgb_end[0], rgb_end[1], rgb_end[2]);
        colors.push({
          id: colorIds[i],
          start: hex_start,
          end: hex_end
        });
      }
      return colors;
    }
    function buildChart(config) {
      const animProgress = (config.animation == null) 
        ? 1.0 
        : Math.max(0, Math.min(1, Number(config.animation)));
      const total = config.items.reduce((sum, item) => sum + item.value, 0);
      if (total === 0) return;
      const layout = {
        centerY: 260,
        outerRadius: 150,
        innerRadius: 75,
        legendItemHeight: 30,
        legendColorBoxSize: 12,
        legendTextOffsetX: 18
      };
      const svgElement = document.getElementById('donut-chart-svg');
      const defs = svgElement.querySelector('defs');
      const existingGrads = defs.querySelectorAll('linearGradient');
      existingGrads.forEach(grad => grad.remove());
      config.colors.forEach(color => {
        const gradId = 'grad' + color.id;
        let gradient = document.getElementById(gradId);
        if(gradient) gradient.parentNode.removeChild(gradient);
        gradient = createSVGElement('linearGradient', { id: gradId, x1:"0", x2:"0", y1:"0", y2:"1" });
        gradient.appendChild(createSVGElement('stop', { offset: '0%', 'stop-color': color.start }));
        gradient.appendChild(createSVGElement('stop', { offset: '100%', 'stop-color': color.end }));
        defs.appendChild(gradient);
      });
      const groups = {
        title: document.getElementById('title-group'),
        donut: document.getElementById('donut-group'),
        legend: document.getElementById('legend-group'),
        source: document.getElementById('source-group')
      };
      for (const key in groups) {
        while (groups[key].firstChild) groups[key].removeChild(groups[key].firstChild);
      }
      const viewBoxWidth = svgElement.viewBox.baseVal.width;
      const centerX = viewBoxWidth / 4;
      const legendX = viewBoxWidth / 2 + 30;
      config.items.forEach((item, index) => {
        const percentage = ((item.value / total) * 100).toFixed(0);
        const legendItemGroup = createSVGElement('g', {
          class: 'legend-item-group',
          transform: 'translate(0, ' + (index * layout.legendItemHeight) + ')'
        });
        const colorData = config.colors.find(c => c.id === item.id);
        if (!colorData) {
            console.error("Color not found for ID:", item.id);
            return; 
        }
        const color = colorData.end;
        legendItemGroup.appendChild(createSVGElement('rect', { 
          x: 0, y: 0, 
          width: layout.legendColorBoxSize, height: layout.legendColorBoxSize, 
          rx: 2, fill: color, filter: 'url(#shadow)' 
        }));
        const textEl = createSVGElement('text', { 
          x: layout.legendTextOffsetX, 
          y: layout.legendColorBoxSize / 2, 
          class: 'legend-item' 
        });
        textEl.textContent = item.label + ' (';
        const percentageSpan = createSVGElement('tspan', { class: 'legend-percentage' }, percentage + '%');
        textEl.appendChild(percentageSpan);
        textEl.appendChild(document.createTextNode(')'));
        legendItemGroup.appendChild(textEl);
        groups.legend.appendChild(legendItemGroup);
      });
      const legendHeight = config.items.length * layout.legendItemHeight;
      const legendY = layout.centerY - (legendHeight / 2);
      groups.legend.setAttribute('transform', 'translate(' + legendX + ', ' + legendY + ')');
      const titleX = viewBoxWidth / 2;
      groups.title.appendChild(createSVGElement('text', { x: titleX, y: 50, class: 'title' }, config.title));
      groups.title.appendChild(createSVGElement('text', { x: titleX, y: 75, class: 'subtitle' }, config.subtitle));
      groups.source.appendChild(createSVGElement('text', { x: viewBoxWidth - 10, y: 440, class: 'source' }, config.source));
    let finalStartAngle = 0;
    const finalAngles = config.items.map(item => {
        const sweep = (item.value / total) * 360;
        const start = finalStartAngle;
        finalStartAngle += sweep;
        return { start: start, end: finalStartAngle };
    });
    const numItems = config.items.length;
    const equalSweep = 360 / numItems;
    const equalAngles = config.items.map((item, i) => {
        const start = i * equalSweep;
        const end = (i + 1) * equalSweep;
        return { start: start, end: end };
    });
      let startAngle = 0;
      config.items.forEach((item, index) => {
    const eqStart = equalAngles[index].start;
    const eqEnd = equalAngles[index].end;
    const finStart = finalAngles[index].start;
    const finEnd = finalAngles[index].end;
    const animatedStartAngle = eqStart + (finStart - eqStart) * animProgress;
    const animatedEndAngle = eqEnd + (finEnd - eqEnd) * animProgress;
const pathData = createDonutSegmentPath(centerX, layout.centerY, layout.outerRadius, animatedStartAngle, animatedEndAngle);
    groups.donut.appendChild(createSVGElement('path', {
        d: pathData,
        fill: 'url(#grad' + item.id + ')',
class: 'donut-segment'
    }));
});
      groups.donut.appendChild(createSVGElement('circle', { cx: centerX, cy: layout.centerY, r: layout.innerRadius, fill: 'white' }));
      groups.donut.appendChild(createSVGElement('text', { x: centerX, y: layout.centerY - 12, class: 'center-total-value' }, total));
      groups.donut.appendChild(createSVGElement('text', { x: centerX, y: layout.centerY + 15, class: 'center-total-label' }, config.centerLabel));
    }
    try {
      const configJson = document.getElementById('chart-json-data').textContent;
      if (!configJson || configJson.trim() === '') {
        throw new Error("JSON data is empty.");
      }
      let fullConfig;
      try {
        fullConfig = JSON.parse(configJson);
      } catch (parseError) {
        throw new Error('JSON parsing error: ' + parseError.message);
      }
      if (!fullConfig || typeof fullConfig !== 'object' || !fullConfig.data) {
        throw new Error("Invalid JSON structure: Missing 'data' key.");
      }
      const config = fullConfig.data;
      if (config.colors && typeof config.colors === 'string') {
        if (!config.items || !Array.isArray(config.items)) {
            throw new Error("Invalid JSON: 'items' array is required to generate colors.");
        }
        const colorIds = config.items.map(item => item.id);
        const numColors = colorIds.length;
        if (numColors === 0) {
             throw new Error("Invalid JSON: 'items' array is empty.");
        }
        let generatedColors;
        const colorKeyword = config.colors.toLowerCase();
        const geminiGradientDarken = 15;
        if (colorKeyword === '#gemini') {
          const GEMINI_STOPS = ['#E68A9C', '#9F63D0', '#616AD8'];
          const hexPalette = generateGradientScale(GEMINI_STOPS, numColors);
          generatedColors = [];
          for (let i = 0; i < numColors; i++) {
            const hex_start = hexPalette[i];
            const rgb_start = hexToRgb(hex_start);
            const [h, s, l] = rgbToHsl(rgb_start.r, rgb_start.g, rgb_start.b);
            const l_end = Math.max(0, Math.min(100, l - geminiGradientDarken));
            const rgb_end = hslToRgb(h, s, l_end);
            const hex_end = rgbToHex(rgb_end[0], rgb_end[1], rgb_end[2]);
            generatedColors.push({
              id: colorIds[i],
              start: hex_start,
              end: hex_end
            });
          }
        } else if (colorKeyword === '#night') {
          const NIGHT_COLORS = [
            { start: '#8d58d3', end: '#704ca4' },
            { start: '#8f83d6', end: '#7469a8' },
            { start: '#7891dd', end: '#5c6fbb' },
            { start: '#50a89d', end: '#258a7f' },
            { start: '#54c5d5', end: '#34a0b0' }
          ];
          generatedColors = [];
          for (let i = 0; i < numColors; i++) {
            const colorPair = NIGHT_COLORS[i] || NIGHT_COLORS[NIGHT_COLORS.length - 1];
            generatedColors.push({
              id: colorIds[i],
              start: colorPair.start,
              end: colorPair.end
            });
          }
        } else {
          const primaryColor = config.colors;
          generatedColors = generateMonochromaticScale(primaryColor, colorIds);
          if (generatedColors.length === 0) {
              throw new Error("Invalid HEX color string in 'colors': " + primaryColor);
          }
        }
        if (generatedColors && generatedColors.length > 0) {
          config.colors = generatedColors;
        } else {
            throw new Error("Failed to generate colors from keyword: " + config.colors);
        }
      } else if (!config.colors || !Array.isArray(config.colors)) {
          throw new Error("'colors' must be an array of color objects or a single primary color string.");
      }
      if (!config || typeof config !== 'object' || !config.items || !Array.isArray(config.items)) {
        throw new Error("Invalid JSON structure: Missing required fields like 'items' inside 'data'.");
      }
      buildChart(config);
    } catch (e) {
      console.error("Chart Error:", e);
      const svgRoot = document.getElementById('donut-chart-svg');
      const textEl = createSVGElement('text', { x: 300, y: 230, 'text-anchor': 'middle', fill: 'red' });
      textEl.textContent = e.message.includes('JSON') ? 'JSONデータにエラーがあります。' : 'グラフ描画エラーが発生しました。';
      if (e.message.includes('JSON parsing error')) {
        textEl.textContent += ' (' + e.message.split(': ')[1] + ')';
      }
      svgRoot.appendChild(textEl);
    }
  </script>
</svg>
  `;