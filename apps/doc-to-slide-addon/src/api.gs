var global = this;
(function() {
  "use strict";
  class Presentation {
    constructor(title) {
      this.title = title;
      this._slides = [];
    }
    addSlide(slide) {
      this._slides.push(slide);
    }
    get slides() {
      return [...this._slides];
    }
  }
  class Slide {
    constructor(title, content, layout = "CONTENT", subtitle, notes, rawData) {
      this.title = title;
      this.content = content;
      this.layout = layout;
      this.subtitle = subtitle;
      this.notes = notes;
      this.rawData = rawData;
    }
  }
  class SlideTitle {
    constructor(value) {
      this.value = value;
    }
  }
  class SlideContent {
    constructor(items) {
      this.items = items;
    }
  }
  class PresentationApplicationService {
    constructor(slideRepository) {
      this.slideRepository = slideRepository;
    }
    createPresentation(request) {
      const presentation = new Presentation(request.title);
      for (const slideData of request.slides) {
        const title = new SlideTitle(slideData.title);
        const content = new SlideContent(slideData.content);
        let layout = slideData.layout;
        if (!layout && slideData.type) {
          const typeUpper = String(slideData.type).toUpperCase();
          if (["TITLE", "AGENDA", "SECTION", "CONCLUSION"].includes(typeUpper)) {
            layout = typeUpper;
          } else {
            layout = "CONTENT";
          }
        }
        if (!layout) {
          layout = "CONTENT";
        }
        const slide = new Slide(title, content, layout, slideData.subtitle, slideData.notes, slideData);
        presentation.addSlide(slide);
      }
      return this.slideRepository.createPresentation(presentation, request.templateId, request.destinationId, request.settings);
    }
  }
  const DEFAULT_THEME = {
    basePx: {
      width: 960,
      height: 540
    },
    fonts: {
      // Noto Sans JP is good, but let's assume we can use different weights via styles
      family: "Meiryo UI",
      sizes: {
        title: 48,
        date: 14,
        sectionTitle: 52,
        contentTitle: 28,
        subhead: 20,
        // Increased visibility
        body: 18,
        // "18pt minimum" rule
        footer: 11,
        chip: 12,
        laneTitle: 16,
        small: 12,
        processStep: 16,
        axis: 12,
        ghostNum: 250
      }
    },
    colors: {
      primary: "#8FB130",
      // Brand Main (Vibrant Leaf Green)
      deepPrimary: "#526717",
      // Brand Deep (Moss Green)
      textPrimary: "#333333",
      // Soft Black
      textSmallFont: "#545454",
      // Slightly warmer gray for text
      backgroundWhite: "#FFFFFF",
      cardBg: "#FFFFFF",
      backgroundGray: "#F8F9FA",
      // Keep neutral background for logos
      faintGray: "#F8F9FA",
      ghostGray: "#E2E4E6",
      // 
      tableHeaderBg: "#E8EAE6",
      // Hint of green in gray
      laneBorder: "#E0E0E0",
      cardBorder: "#E0E0E0",
      neutralGray: "#9E9E9E",
      processArrow: "#526717"
      // Use Deep color for contrast
    },
    diagram: {
      laneGapPx: 30,
      // Wider gaps
      lanePadPx: 20,
      // More padding
      laneTitleHeightPx: 40,
      cardGapPx: 20,
      // Airy layout
      cardMinHeightPx: 60,
      cardMaxHeightPx: 90,
      arrowHeightPx: 8,
      // Thinner, elegant arrows
      arrowGapPx: 12
    },
    logos: {
      header: "",
      closing: ""
    },
    footerText: "",
    positions: {
      // Updated positions for "Editorial" look - wider margins (60px side margins)
      titleSlide: {
        logo: { left: 60, top: 60, width: 150 },
        title: { left: 60, top: 200, width: 840, height: 120 },
        date: { left: 60, top: 480, width: 300, height: 40 }
      },
      contentSlide: {
        headerLogo: { right: 30, top: 30, width: 80 },
        title: { left: 60, top: 30, width: 840, height: 70 },
        titleUnderline: { left: 60, top: 100, width: 80, height: 3 },
        // Short elegant underline
        subhead: { left: 60, top: 110, width: 840, height: 30 },
        body: { left: 60, top: 160, width: 840, height: 320 },
        twoColLeft: { left: 60, top: 160, width: 400, height: 320 },
        // 40px gap
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
        ghostNum: { left: 400, top: 100, width: 550, height: 350 },
        // Adjusted
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
    },
    backgroundImages: {
      title: "",
      closing: "",
      section: "",
      main: ""
    }
  };
  class LayoutManager {
    /**
     * Create a new LayoutManager
     * @param pageW_pt - Actual page width in points
     * @param pageH_pt - Actual page height in points
     * @param theme - Theme configuration (defaults to DEFAULT_THEME)
     */
    constructor(pageW_pt, pageH_pt, theme = DEFAULT_THEME) {
      this.pxToPtRatio = 0.75;
      this.theme = theme;
      this.pageW_pt = pageW_pt;
      this.pageH_pt = pageH_pt;
      const baseW_pt = this.pxToPt(theme.basePx.width);
      const baseH_pt = this.pxToPt(theme.basePx.height);
      this.scaleX = pageW_pt / baseW_pt;
      this.scaleY = pageH_pt / baseH_pt;
    }
    /**
     * Convert pixels to points
     */
    pxToPt(px) {
      return px * this.pxToPtRatio;
    }
    /**
     * Get position from a dot-notation path (e.g., "contentSlide.title")
     */
    getPositionFromPath(path) {
      return path.split(".").reduce(
        (obj, key) => obj && obj[key],
        this.theme.positions
      );
    }
    /**
     * Get a rectangle specification converted to points and scaled
     * @param spec - Either a dot-notation path string or a position object
     */
    getRect(spec) {
      const pos = typeof spec === "string" ? this.getPositionFromPath(spec) : spec;
      if (!pos) return { left: 0, top: 0, width: 0, height: 0 };
      let left_px = pos.left;
      if (pos.right !== void 0 && pos.left === void 0) {
        left_px = this.theme.basePx.width - pos.right - pos.width;
      }
      if (left_px === void 0 && pos.right === void 0) {
        left_px = 0;
      }
      return {
        left: left_px !== void 0 ? this.pxToPt(left_px) * this.scaleX : 0,
        top: pos.top !== void 0 ? this.pxToPt(pos.top) * this.scaleY : 0,
        width: pos.width !== void 0 ? this.pxToPt(pos.width) * this.scaleX : 0,
        height: pos.height !== void 0 ? this.pxToPt(pos.height) * this.scaleY : 0
      };
    }
    /**
     * Get the current theme
     */
    getTheme() {
      return this.theme;
    }
  }
  class GasTitleSlideGenerator {
    constructor(creditImageBlob) {
      this.creditImageBlob = creditImageBlob;
    }
    generate(slide, data, layout, pageNum, settings, imageUpdateOption = "update") {
      const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
      if (titlePlaceholder) {
        if (data.title) {
          try {
            titlePlaceholder.asShape().getText().setText(data.title);
          } catch (e) {
            Logger.log(`Warning: Title placeholder found but text could not be set. ${e}`);
          }
        } else {
          titlePlaceholder.remove();
        }
      } else {
        Logger.log("Title Slide: WARNING - No Title Placeholder found on this layout.");
      }
      const subtitlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
      if (subtitlePlaceholder) {
        if (data.date) {
          try {
            subtitlePlaceholder.asShape().getText().setText(data.date);
          } catch (e) {
          }
        } else {
          try {
            subtitlePlaceholder.asShape().getText().setText("");
          } catch (e) {
          }
        }
      } else {
        const bodyPlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
        if (bodyPlaceholder && data.date) {
          try {
            bodyPlaceholder.asShape().getText().setText(data.date);
          } catch (e) {
          }
        }
      }
    }
  }
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
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
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
  function hslToHex(h, s, l) {
    s /= 100;
    l /= 100;
    const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(h / 60 % 2 - 1)), m = l - c / 2;
    let r = 0, g = 0, b = 0;
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
    if (!rgb) return "#F8F9FA";
    const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
    return hslToHex(hsl.h, saturation, lightness);
  }
  function lightenColor(color, amount) {
    const rgb = hexToRgb(color);
    if (!rgb) return color;
    const lighten = (c) => Math.min(255, Math.round(c + (255 - c) * amount));
    const newR = lighten(rgb.r);
    const newG = lighten(rgb.g);
    const newB = lighten(rgb.b);
    return `#${newR.toString(16).padStart(2, "0")}${newG.toString(16).padStart(2, "0")}${newB.toString(16).padStart(2, "0")}`;
  }
  function darkenColor(color, amount) {
    const rgb = hexToRgb(color);
    if (!rgb) return color;
    const darken = (c) => Math.max(0, Math.round(c * (1 - amount)));
    const newR = darken(rgb.r);
    const newG = darken(rgb.g);
    const newB = darken(rgb.b);
    return `#${newR.toString(16).padStart(2, "0")}${newG.toString(16).padStart(2, "0")}${newB.toString(16).padStart(2, "0")}`;
  }
  function generatePyramidColors(baseColor, levels) {
    const colors = [];
    for (let i = 0; i < levels; i++) {
      const lightenAmount = i / Math.max(1, levels - 1) * 0.6;
      colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
  }
  function generateStepUpColors(baseColor, steps) {
    const colors = [];
    for (let i = 0; i < steps; i++) {
      const lightenAmount = 0.6 * (1 - i / Math.max(1, steps - 1));
      colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
  }
  function generateProcessColors(baseColor, steps) {
    const colors = [];
    for (let i = 0; i < steps; i++) {
      const lightenAmount = 0.5 * (1 - i / Math.max(1, steps - 1));
      colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
  }
  function generateTimelineCardColors(baseColor, milestones) {
    const colors = [];
    for (let i = 0; i < milestones; i++) {
      const lightenAmount = 0.4 * (1 - i / Math.max(1, milestones - 1));
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
  const ColorUtils = /* @__PURE__ */ Object.freeze(/* @__PURE__ */ Object.defineProperty({
    __proto__: null,
    darkenColor,
    generateCompareColors,
    generateProcessColors,
    generatePyramidColors,
    generateStepUpColors,
    generateTimelineCardColors,
    generateTintedGray,
    hexToRgb,
    hslToHex,
    lightenColor,
    rgbToHsl
  }, Symbol.toStringTag, { value: "Module" }));
  function getThemeFromLayout(layout) {
    if (layout && typeof layout.getTheme === "function") {
      return layout.getTheme();
    }
    return DEFAULT_THEME;
  }
  function applyTextStyle(textRange, opt, theme = DEFAULT_THEME) {
    const style = textRange.getTextStyle();
    let defaultColor;
    if (opt.fontType === "large") {
      defaultColor = theme.colors.textPrimary;
    } else {
      defaultColor = theme.colors.textSmallFont;
    }
    style.setFontFamily(theme.fonts.family).setForegroundColor(opt.color || defaultColor).setFontSize(opt.size || theme.fonts.sizes.body).setBold(opt.bold || false);
    if (opt.align) {
      try {
        textRange.getParagraphs().forEach((p) => {
          p.getRange().getParagraphStyle().setParagraphAlignment(opt.align);
        });
      } catch (e) {
      }
    }
  }
  function setStyledText(shapeOrCell, rawText, baseOpt) {
    const parsed = parseInlineStyles(rawText || "");
    const tr = shapeOrCell.getText();
    tr.setText(parsed.output);
    applyTextStyle(tr, baseOpt || {});
    applyStyleRanges(tr, parsed.ranges);
  }
  function setBulletsWithInlineStyles(shape, points, theme = DEFAULT_THEME) {
    const joiner = "\n\n";
    let combined = "";
    const ranges = [];
    (points || []).forEach((pt, idx) => {
      const parsed = parseInlineStyles(String(pt || ""));
      const bullet = parsed.output;
      if (idx > 0) combined += joiner;
      const start = combined.length;
      combined += bullet;
      parsed.ranges.forEach((r) => ranges.push({
        start: start + r.start,
        end: start + r.end,
        bold: r.bold,
        color: r.color
      }));
    });
    const tr = shape.getText();
    tr.setText(combined || "—");
    applyTextStyle(tr, {
      size: theme.fonts.sizes.body
    }, theme);
    try {
      tr.getParagraphs().forEach((p) => {
        p.getRange().getParagraphStyle().setLineSpacing(130).setSpaceBelow(12);
      });
    } catch (e) {
    }
    applyStyleRanges(tr, ranges);
  }
  function checkSpacing(s, out, i, nextCharIndex) {
    let prefix = "";
    let suffix = "";
    if (out.length > 0 && !/\s$/.test(out)) {
      prefix = " ";
    }
    if (nextCharIndex < s.length && !/^\s/.test(s[nextCharIndex])) {
      suffix = " ";
    }
    return { prefix, suffix };
  }
  function parseInlineStyles(s) {
    const ranges = [];
    let out = "";
    let i = 0;
    while (i < s.length) {
      if (s[i] === "*" && s[i + 1] === "*" && s[i + 2] === "[" && s[i + 3] === "[") {
        const contentStart = i + 4;
        const close = s.indexOf("]]**", contentStart);
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
            color: DEFAULT_THEME.colors.primary
          };
          ranges.push(rangeObj);
          i = nextCharIndex;
          continue;
        }
      }
      if (s[i] === "[" && s[i + 1] === "[") {
        const close = s.indexOf("]]", i + 2);
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
            color: DEFAULT_THEME.colors.primary
          };
          ranges.push(rangeObj);
          i = nextCharIndex;
          continue;
        }
      }
      if (s[i] === "*" && s[i + 1] === "*") {
        const close = s.indexOf("**", i + 2);
        if (close !== -1) {
          let content = s.substring(i + 2, close);
          if (content.indexOf("[[") === -1) {
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
              bold: true
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
  function applyStyleRanges(textRange, ranges) {
    ranges.forEach((r) => {
      try {
        const sub = textRange.getRange(r.start, r.end);
        if (sub.isEmpty()) return;
        const st = sub.getTextStyle();
        if (r.bold) st.setBold(true);
        if (r.color) st.setForegroundColor(r.color);
        if (r.size) st.setFontSize(r.size);
      } catch (e) {
      }
    });
  }
  function insertImageFromUrlOrFileId(urlOrFileId) {
    if (!urlOrFileId || typeof urlOrFileId !== "string") return null;
    function extractFileIdFromUrl(url) {
      if (!url || typeof url !== "string") return null;
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
    } else if (urlOrFileId.startsWith("data:image/")) {
      try {
        const parts = urlOrFileId.split(",");
        if (parts.length !== 2) throw new Error("Invalid Base64 format.");
        const mimeMatch = parts[0].match(/:(.*?);/);
        if (!mimeMatch) return null;
        const mimeType = mimeMatch[1];
        const base64Data = parts[1];
        const decodedData = Utilities.base64Decode(base64Data);
        return Utilities.newBlob(decodedData, mimeType);
      } catch (e) {
        return null;
      }
    } else {
      try {
        if (urlOrFileId.startsWith("http")) {
          const response = UrlFetchApp.fetch(urlOrFileId);
          return response.getBlob();
        }
      } catch (e) {
      }
      return null;
    }
  }
  function normalizeImages(arr) {
    return (arr || []).map((v) => typeof v === "string" ? {
      url: v
    } : v && v.url ? v : null).filter(Boolean).slice(0, 6);
  }
  function renderImagesInArea(slide, layout, area, images, imageUpdateOption = "update") {
    if (!images || !images.length) return;
    if (imageUpdateOption !== "update") {
      return;
    }
    const n = Math.min(6, images.length);
    let cols = n === 1 ? 1 : n <= 4 ? 2 : 3;
    const rows = Math.ceil(n / cols);
    const gap = layout.pxToPt(10);
    const cellW = (area.width - gap * (cols - 1)) / cols, cellH = (area.height - gap * (rows - 1)) / rows;
    for (let i = 0; i < n; i++) {
      const r = Math.floor(i / cols), c = i % cols;
      try {
        const img = slide.insertImage(images[i].url);
        const scale = Math.min(cellW / img.getWidth(), cellH / img.getHeight());
        const w = img.getWidth() * scale, h = img.getHeight() * scale;
        img.setWidth(w).setHeight(h).setLeft(area.left + c * (cellW + gap) + (cellW - w) / 2).setTop(area.top + r * (cellH + gap) + (cellH - h) / 2);
      } catch (e) {
      }
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
  function addFooter(slide, layout, pageNum, settings, creditImageBlob) {
    const theme = getThemeFromLayout(layout);
    if (theme.footerText && theme.footerText.trim() !== "") {
      const leftRect = layout.getRect("footer.leftText");
      const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftRect.left, leftRect.top, leftRect.width, leftRect.height);
      const tr = leftShape.getText();
      tr.setText(theme.footerText);
      applyTextStyle(tr, {
        size: theme.fonts.sizes.footer,
        fontType: "large"
      });
      try {
        leftShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {
      }
    }
    if (pageNum > 0 && settings && settings.showPageNumber) {
      const rightRect = layout.getRect("footer.rightPage");
      const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightRect.left, rightRect.top, rightRect.width, rightRect.height);
      const tr = rightShape.getText();
      tr.setText(String(pageNum));
      applyTextStyle(tr, {
        size: theme.fonts.sizes.footer,
        color: theme.colors.primary,
        align: SlidesApp.ParagraphAlignment.END
      });
      try {
        rightShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {
      }
    }
  }
  function safeGetRect(layout, path) {
    try {
      const rect = layout.getRect(path);
      if (rect && (typeof rect.left === "number" || rect.left === void 0) && typeof rect.top === "number" && typeof rect.width === "number" && typeof rect.height === "number") {
        if (rect.left === void 0) {
          return null;
        }
        return rect;
      }
      return null;
    } catch (e) {
      return null;
    }
  }
  function drawSubheadIfAny(slide, layout, key, subhead, preCalculatedWidthPt = null) {
    if (!subhead) return 0;
    const theme = getThemeFromLayout(layout);
    const rect = safeGetRect(layout, `${key}.subhead`);
    if (!rect) {
      return 0;
    }
    const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left, rect.top, rect.width, rect.height);
    setStyledText(box, subhead, {
      size: theme.fonts.sizes.subhead,
      fontType: "large"
    });
    return layout.pxToPt(36);
  }
  function createContentCushion(slide, area, settings, layout) {
    const theme = getThemeFromLayout(layout);
    const cushion = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, area.width, area.height);
    cushion.getFill().setSolidFill(theme.colors.cardBg);
    try {
      const border = cushion.getBorder();
      border.setTransparent();
    } catch (e) {
    }
  }
  function offsetRect(rect, dx, dy) {
    return {
      left: rect.left + dx,
      top: rect.top + dy,
      width: rect.width,
      height: rect.height
    };
  }
  function adjustAreaForSubhead(area, subhead, layout) {
    return area;
  }
  class GasSectionSlideGenerator {
    constructor(creditImageBlob) {
      this.creditImageBlob = creditImageBlob;
      this.sectionCounter = 0;
    }
    generate(slide, data, layout, pageNum, settings, imageUpdateOption = "update") {
      const theme = layout.getTheme();
      this.sectionCounter++;
      const parsedNum = (() => {
        if (Number.isFinite(data.sectionNo)) {
          return Number(data.sectionNo);
        }
        const m = String(data.title || "").match(/^\s*(\d+)[\.．]/);
        return m ? Number(m[1]) : this.sectionCounter;
      })();
      const num = String(parsedNum).padStart(2, "0");
      const ghostRect = layout.getRect("sectionSlide.ghostNum");
      let ghostImageInserted = false;
      if (imageUpdateOption === "update") {
        if (data.ghostImageBase64 && ghostRect) {
          try {
            const imageData = insertImageFromUrlOrFileId(data.ghostImageBase64);
            if (imageData && typeof imageData !== "string") {
              const ghostImage = slide.insertImage(imageData);
              const imgWidth = ghostImage.getWidth();
              const imgHeight = ghostImage.getHeight();
              const scale = Math.min(ghostRect.width / imgWidth, ghostRect.height / imgHeight);
              const w = imgWidth * scale;
              const h = imgHeight * scale;
              ghostImage.setWidth(w).setHeight(h).setLeft(ghostRect.left + (ghostRect.width - w) / 2).setTop(ghostRect.top + (ghostRect.height - h) / 2);
              ghostImageInserted = true;
            }
          } catch (e) {
          }
        }
      }
      if (!ghostImageInserted && ghostRect && imageUpdateOption === "update") {
        const ghost = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, ghostRect.left, ghostRect.top, ghostRect.width, ghostRect.height);
        ghost.getText().setText(num);
        const ghostTextStyle = ghost.getText().getTextStyle();
        ghostTextStyle.setFontFamily(theme.fonts.family).setFontSize(theme.fonts.sizes.ghostNum).setBold(true);
        try {
          ghostTextStyle.setForegroundColor(theme.colors.ghostGray);
        } catch (e) {
          ghostTextStyle.setForegroundColor(theme.colors.ghostGray);
        }
        try {
          ghost.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
        ghost.sendToBack();
      }
      const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
      if (titlePlaceholder) {
        try {
          titlePlaceholder.asShape().getText().setText(data.title || "");
        } catch (e) {
          Logger.log(`Warning: Section Title placeholder found but text could not be set. ${e}`);
        }
      } else {
        const titleRect = layout.getRect("sectionSlide.title");
        const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, titleRect.top, titleRect.width, titleRect.height);
        titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        setStyledText(titleShape, data.title, {
          size: theme.fonts.sizes.sectionTitle,
          bold: true,
          align: SlidesApp.ParagraphAlignment.CENTER,
          fontType: "large"
        });
      }
      addFooter(slide, layout, pageNum, settings, this.creditImageBlob);
    }
  }
  class GasContentSlideGenerator {
    constructor(creditImageBlob) {
      this.creditImageBlob = creditImageBlob;
    }
    generate(slide, data, layout, pageNum, settings, imageUpdateOption = "update") {
      const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
      if (titlePlaceholder) {
        if (data.title) {
          try {
            titlePlaceholder.asShape().getText().setText(data.title);
          } catch (e) {
            Logger.log(`Warning: Content Title placeholder found but text could not be set. ${e}`);
          }
        } else {
          titlePlaceholder.remove();
        }
      }
      const subheadWidthPt = data && typeof data._subhead_widthPt === "number" ? data._subhead_widthPt : null;
      const dy = drawSubheadIfAny(slide, layout, "contentSlide", data.subhead, subheadWidthPt);
      let points = Array.isArray(data.points) ? data.points.slice(0) : [];
      const isAgenda = /(agenda|アジェンダ|目次|本日お伝えすること)/i.test(String(data.title || ""));
      if (isAgenda) {
        slide.getBackground().setSolidFill("#FFFFFF");
        if (points.length === 0) {
          points = ["本日の目的", "進め方", "次のアクション"];
        }
      }
      const hasImages = Array.isArray(data.images) && data.images.length > 0;
      const isTwo = !!(data.twoColumn || data.columns);
      const bodies = slide.getPlaceholders().filter((p) => {
        if (typeof p.getPlaceholderType === "function") {
          return p.getPlaceholderType() === SlidesApp.PlaceholderType.BODY;
        }
        try {
          const s = p.asShape();
          if (typeof s.getPlaceholderType === "function") {
            return s.getPlaceholderType() === SlidesApp.PlaceholderType.BODY;
          }
        } catch (e) {
        }
        return false;
      });
      if (dy > 0) {
        bodies.forEach((p) => {
          try {
            const s = p.asShape();
            s.setTop(s.getTop() + dy);
            s.setHeight(s.getHeight() - dy);
          } catch (e) {
            Logger.log("Could not adjust body placeholder: " + e);
          }
        });
      }
      if (isTwo && bodies.length >= 2) {
        let L = [], R = [];
        if (Array.isArray(data.columns) && data.columns.length === 2) {
          L = data.columns[0] || [];
          R = data.columns[1] || [];
        } else {
          const mid = Math.ceil(points.length / 2);
          L = points.slice(0, mid);
          R = points.slice(mid);
        }
        const sortedBodies = bodies.map((p) => p.asShape()).sort((a, b) => a.getLeft() - b.getLeft());
        setBulletsWithInlineStyles(sortedBodies[0], L);
        setBulletsWithInlineStyles(sortedBodies[1], R);
      } else if (isTwo) {
        if (Array.isArray(data.columns) && data.columns.length === 2) ;
        let L = [], R = [];
        if (Array.isArray(data.columns) && data.columns.length === 2) {
          L = data.columns[0] || [];
          R = data.columns[1] || [];
        } else {
          const mid = Math.ceil(points.length / 2);
          L = points.slice(0, mid);
          R = points.slice(mid);
        }
        const baseLeftRect = layout.getRect("contentSlide.twoColLeft");
        const baseRightRect = layout.getRect("contentSlide.twoColRight");
        const adjustedLeftRect = adjustAreaForSubhead(baseLeftRect, data.subhead);
        const adjustedRightRect = adjustAreaForSubhead(baseRightRect, data.subhead);
        const leftRect = offsetRect(adjustedLeftRect, 0, dy);
        const rightRect = offsetRect(adjustedRightRect, 0, dy);
        const padding = layout.pxToPt(20);
        const leftTextRect = { left: leftRect.left + padding, top: leftRect.top + padding, width: leftRect.width - padding * 2, height: leftRect.height - padding * 2 };
        const rightTextRect = { left: rightRect.left + padding, top: rightRect.top + padding, width: rightRect.width - padding * 2, height: rightRect.height - padding * 2 };
        const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftTextRect.left, leftTextRect.top, leftTextRect.width, leftTextRect.height);
        const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightTextRect.left, rightTextRect.top, rightTextRect.width, rightTextRect.height);
        setBulletsWithInlineStyles(leftShape, L);
        setBulletsWithInlineStyles(rightShape, R);
      } else {
        if (bodies.length > 0) {
          const bodyShape = bodies[0].asShape();
          setBulletsWithInlineStyles(bodyShape, points);
        } else {
          const baseBodyRect = layout.getRect("contentSlide.body");
          const adjustedBodyRect = adjustAreaForSubhead(baseBodyRect, data.subhead);
          const bodyRect = offsetRect(adjustedBodyRect, 0, dy);
          const padding = layout.pxToPt(20);
          const textRect = { left: bodyRect.left + padding, top: bodyRect.top + padding, width: bodyRect.width - padding * 2, height: bodyRect.height - padding * 2 };
          const bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
          setBulletsWithInlineStyles(bodyShape, points);
        }
      }
      if (hasImages && !points.length && !isTwo) {
        const baseArea = layout.getRect("contentSlide.body");
        const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead);
        const area = offsetRect(adjustedArea, 0, dy);
        createContentCushion(slide, area, settings, layout);
        renderImagesInArea(slide, layout, area, normalizeImages(data.images), imageUpdateOption);
      }
      addFooter(slide, layout, pageNum, settings, this.creditImageBlob);
    }
  }
  class CardsDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const items = data.items || [];
      if (!items.length) return;
      const type = (data.type || "").toLowerCase();
      const hasHeader = type.includes("headercards");
      const cols = data.columns || Math.min(items.length, 3);
      const rows = Math.ceil(items.length / cols);
      const gap = layout.pxToPt(30);
      const cardW = (area.width - gap * (cols - 1)) / cols;
      const cardH = (area.height - gap * (rows - 1)) / rows;
      items.forEach((item, i) => {
        const r = Math.floor(i / cols);
        const c = i % cols;
        const x = area.left + c * (cardW + gap);
        const y = area.top + r * (cardH + gap);
        let title = "";
        let desc = "";
        if (typeof item === "string") {
          const lines = item.split("\n");
          title = lines[0] || "";
          desc = lines.slice(1).join("\n") || "";
        } else {
          title = item.title || item.label || "";
          desc = item.desc || item.description || item.text || "";
        }
        if (hasHeader) {
          const barH = layout.pxToPt(4);
          const bar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, cardW, barH);
          bar.getFill().setSolidFill(settings.primaryColor);
          bar.getBorder().setTransparent();
          const numStr = String(i + 1).padStart(2, "0");
          const numBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + layout.pxToPt(6), cardW, layout.pxToPt(20));
          setStyledText(numBox, numStr, {
            size: 14,
            bold: true,
            color: DEFAULT_THEME.colors.neutralGray,
            align: SlidesApp.ParagraphAlignment.END
          });
          const titleTop = y + layout.pxToPt(6);
          const titleH = layout.pxToPt(30);
          const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleTop, cardW, titleH);
          setStyledText(titleBox, title, {
            size: 18,
            bold: true,
            color: DEFAULT_THEME.colors.textPrimary,
            align: SlidesApp.ParagraphAlignment.START
          });
          try {
            titleBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
          } catch (e) {
          }
          const descTop = titleTop + titleH;
          const descH = cardH - (descTop - y);
          if (descH > 20) {
            const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, descTop, cardW, descH);
            setStyledText(descBox, desc, {
              size: 13,
              color: typeof DEFAULT_THEME.colors.textSmallFont === "string" ? DEFAULT_THEME.colors.textSmallFont : "#424242",
              align: SlidesApp.ParagraphAlignment.START
            });
            try {
              descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
            } catch (e) {
            }
          }
        } else {
          const dotSize = layout.pxToPt(6);
          const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y + layout.pxToPt(8), dotSize, dotSize);
          dot.getFill().setSolidFill(settings.primaryColor);
          dot.getBorder().setTransparent();
          const contentX = x + dotSize + layout.pxToPt(10);
          const contentW = cardW - (dotSize + layout.pxToPt(10));
          const titleH = layout.pxToPt(30);
          const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y, contentW, titleH);
          setStyledText(titleBox, title, {
            size: 16,
            bold: true,
            color: DEFAULT_THEME.colors.textPrimary,
            align: SlidesApp.ParagraphAlignment.START
          });
          try {
            titleBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
          } catch (e) {
          }
          const descTop = y + titleH - layout.pxToPt(5);
          const descH = cardH - (descTop - y);
          if (descH > 20) {
            const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, descTop, contentW, descH);
            setStyledText(descBox, desc, {
              size: 13,
              color: typeof DEFAULT_THEME.colors.textSmallFont === "string" ? DEFAULT_THEME.colors.textSmallFont : "#424242",
              align: SlidesApp.ParagraphAlignment.START
            });
            try {
              descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
            } catch (e) {
            }
          }
        }
      });
    }
  }
  class TimelineDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const milestones = data.milestones || data.items || [];
      if (!milestones.length) return;
      const inner = layout.pxToPt(80), baseY = area.top + area.height * 0.5;
      const leftX = area.left + inner, rightX = area.left + area.width - inner;
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, leftX, baseY, rightX, baseY);
      line.getLineFill().setSolidFill(DEFAULT_THEME.colors.neutralGray);
      line.setWeight(1);
      const dotR = layout.pxToPt(8);
      const gap = milestones.length > 1 ? (rightX - leftX) / (milestones.length - 1) : 0;
      const cardW_pt = layout.pxToPt(160);
      const connectorH = layout.pxToPt(20);
      const dateHeight = layout.pxToPt(20);
      const labelHeight = layout.pxToPt(50);
      const dateToLabelGap = layout.pxToPt(5);
      milestones.forEach((m, i) => {
        const x = leftX + gap * i;
        const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - dotR / 2, baseY - dotR / 2, dotR, dotR);
        dot.getFill().setSolidFill("#FFFFFF");
        dot.getBorder().getLineFill().setSolidFill(settings.primaryColor || DEFAULT_THEME.colors.primary);
        dot.getBorder().setWeight(2);
        const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, baseY - connectorH, x, baseY - dotR / 2);
        connector.getLineFill().setSolidFill(settings.primaryColor || DEFAULT_THEME.colors.primary);
        connector.setWeight(1.5);
        const dateTop = baseY - connectorH - dateHeight;
        const labelTop = dateTop - labelHeight - dateToLabelGap;
        const dateShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - cardW_pt / 2, dateTop, cardW_pt, dateHeight);
        dateShape.getFill().setTransparent();
        dateShape.getBorder().setTransparent();
        setStyledText(dateShape, String(m.date || ""), {
          size: 12,
          bold: true,
          color: settings.primaryColor || DEFAULT_THEME.colors.primary,
          align: SlidesApp.ParagraphAlignment.CENTER
        });
        try {
          dateShape.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
        } catch (e) {
        }
        const labelText = String(m.label || m.state || "");
        const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - cardW_pt / 2, labelTop, cardW_pt, labelHeight);
        labelShape.getFill().setTransparent();
        labelShape.getBorder().setTransparent();
        let bodyFontSize = 12;
        if (labelText.length > 50) bodyFontSize = 10;
        setStyledText(labelShape, labelText, {
          size: bodyFontSize,
          color: DEFAULT_THEME.colors.textPrimary,
          // Stark Black/Gray
          align: SlidesApp.ParagraphAlignment.CENTER
        });
        try {
          dateShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          labelShape.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
        } catch (e) {
        }
      });
    }
  }
  class ProcessDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const steps = data.steps || data.items || [];
      if (!steps.length) return;
      const n = steps.length;
      const startY = area.top + layout.pxToPt(20);
      let currentY = startY;
      const totalH = area.height;
      const itemH = totalH / n;
      const actualItemH = Math.min(itemH, layout.pxToPt(100));
      const margin = layout.pxToPt(20);
      const numColW = layout.pxToPt(50);
      const gapNumToText = layout.pxToPt(10);
      const textLeft = area.left + numColW + gapNumToText;
      const textWidth = area.width - (numColW + gapNumToText);
      steps.forEach((step, i) => {
        const cleanText = String(step || "").replace(/^\s*\d+[\.\s]*/, "");
        const numStr = String(i + 1).padStart(2, "0");
        const numShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, currentY, numColW, layout.pxToPt(40));
        setStyledText(numShape, numStr, {
          size: 28,
          // Slightly smaller but Bold
          bold: true,
          color: settings.primaryColor || DEFAULT_THEME.colors.primary,
          align: SlidesApp.ParagraphAlignment.END
          // Align right towards the text
        });
        try {
          numShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        } catch (e) {
        }
        const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textLeft, currentY + layout.pxToPt(6), textWidth, actualItemH - margin);
        setStyledText(textShape, cleanText, {
          size: 14,
          color: DEFAULT_THEME.colors.textPrimary,
          align: SlidesApp.ParagraphAlignment.START,
          bold: true
          // Title-like weight for the step content
        });
        try {
          textShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        } catch (e) {
        }
        if (i < n - 1) {
          const lineY = currentY + actualItemH - margin / 2;
          const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, textLeft, lineY, area.left + area.width, lineY);
          line.getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
          line.setWeight(0.5);
        }
        currentY += actualItemH;
      });
    }
  }
  class CycleDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const items = data.items || [];
      if (!items.length) return;
      const centerX = area.left + area.width / 2;
      const centerY = area.top + area.height / 2;
      const radius = Math.min(area.width, area.height) * 0.35;
      const circleParams = {
        left: centerX - radius,
        top: centerY - radius,
        width: radius * 2,
        height: radius * 2
      };
      const mainCircle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, circleParams.left, circleParams.top, circleParams.width, circleParams.height);
      mainCircle.getFill().setTransparent();
      mainCircle.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
      mainCircle.getBorder().setWeight(1);
      mainCircle.getBorder().setDashStyle(SlidesApp.DashStyle.DOT);
      if (data.centerText) {
        const centerW = radius * 1.2;
        const centerTextBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, centerX - centerW / 2, centerY - layout.pxToPt(20), centerW, layout.pxToPt(40));
        setStyledText(centerTextBox, data.centerText, {
          size: 18,
          bold: true,
          align: SlidesApp.ParagraphAlignment.CENTER,
          color: DEFAULT_THEME.colors.primary
        });
        try {
          centerTextBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
      }
      const count = items.length;
      const angleStep = 2 * Math.PI / count;
      const startAngle = -Math.PI / 2;
      items.forEach((item, i) => {
        const angle = startAngle + i * angleStep;
        const itemX = centerX + Math.cos(angle) * radius;
        const itemY = centerY + Math.sin(angle) * radius;
        const dotR = layout.pxToPt(10);
        const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, itemX - dotR / 2, itemY - dotR / 2, dotR, dotR);
        dot.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundWhite);
        dot.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.primary);
        dot.getBorder().setWeight(2);
        const isRight = itemX > centerX;
        const textW = layout.pxToPt(140);
        const textH = layout.pxToPt(60);
        const margin = layout.pxToPt(10);
        let textLeft = isRight ? itemX + margin : itemX - textW - margin;
        if (Math.abs(itemX - centerX) < 5) {
          textLeft = itemX - textW / 2;
        }
        const textTop = itemY - textH / 2;
        const labelStr = item.label || "";
        const subLabelStr = item.subLabel || `${String(i + 1).padStart(2, "0")}`;
        const align = Math.abs(itemX - centerX) < 5 ? SlidesApp.ParagraphAlignment.CENTER : isRight ? SlidesApp.ParagraphAlignment.START : SlidesApp.ParagraphAlignment.END;
        const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textLeft, textTop, textW, textH);
        setStyledText(textBox, `${subLabelStr}
${labelStr}`, {
          size: 14,
          bold: true,
          // Title is bold
          color: DEFAULT_THEME.colors.textPrimary,
          align
        });
        if (Math.abs(itemX - centerX) > 5) {
          const lineStart = isRight ? itemX + dotR / 2 : itemX - dotR / 2;
          const lineEnd = isRight ? textLeft : textLeft + textW;
          const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lineStart, itemY, lineEnd, itemY);
          connector.getLineFill().setSolidFill(DEFAULT_THEME.colors.primary);
          connector.setWeight(1);
        }
      });
    }
  }
  class PyramidDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const levels = data.levels || data.items || [];
      if (!levels.length) return;
      const levelsToDraw = levels.slice(0, 5);
      const levelHeight = layout.pxToPt(60);
      const levelGap = layout.pxToPt(20);
      const totalHeight = levelHeight * levelsToDraw.length + levelGap * (levelsToDraw.length - 1);
      const startY = area.top + (area.height - totalHeight) / 2;
      const contentW = layout.pxToPt(500);
      const centerX = area.left + area.width / 2;
      const startX = centerX - contentW / 2;
      levelsToDraw.forEach((level, index) => {
        const y = startY + index * (levelHeight + levelGap);
        index === levelsToDraw.length - 1;
        const numStr = String(index + 1).padStart(2, "0");
        const numW = layout.pxToPt(50);
        const numBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, startX, y, numW, levelHeight);
        setStyledText(numBox, numStr, {
          size: 32,
          bold: true,
          color: settings.primaryColor,
          align: SlidesApp.ParagraphAlignment.END
          // Right align number towards content
        });
        try {
          numBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        } catch (e) {
        }
        const lineX = startX + numW + layout.pxToPt(10);
        const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lineX, y + layout.pxToPt(5), lineX, y + levelHeight - layout.pxToPt(5));
        line.getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
        line.setWeight(1);
        const contentX = lineX + layout.pxToPt(15);
        const textW = contentW - (contentX - startX);
        const levelTitle = level.title || `Level ${index + 1}`;
        const levelDesc = level.description || "";
        const titleH = layout.pxToPt(25);
        const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y, textW, titleH);
        setStyledText(titleBox, levelTitle.toUpperCase(), {
          size: 14,
          bold: true,
          color: DEFAULT_THEME.colors.textPrimary,
          align: SlidesApp.ParagraphAlignment.START
        });
        if (levelDesc) {
          const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y + titleH, textW, levelHeight - titleH);
          setStyledText(descBox, levelDesc, {
            size: 12,
            color: DEFAULT_THEME.colors.neutralGray,
            align: SlidesApp.ParagraphAlignment.START
          });
          try {
            descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
          } catch (e) {
          }
        }
      });
    }
  }
  class TriangleDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const items = data.items || [];
      if (!items.length) return;
      const itemsToDraw = items.slice(0, 3);
      const centerX = area.left + area.width / 2;
      const centerY = area.top + area.height / 2;
      const radius = Math.min(area.width, area.height) / 3.2;
      const positions = [
        { x: centerX, y: centerY - radius },
        { x: centerX + radius * 0.866, y: centerY + radius * 0.5 },
        { x: centerX - radius * 0.866, y: centerY + radius * 0.5 }
      ];
      const circleSize = layout.pxToPt(160);
      const trianglePath = slide.insertShape(SlidesApp.ShapeType.TRIANGLE, centerX - radius / 2, centerY - radius / 2, radius, radius);
      trianglePath.setRotation(0);
      trianglePath.getFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
      trianglePath.getBorder().setTransparent();
      positions.slice(0, itemsToDraw.length).forEach((pos, i) => {
        const item = itemsToDraw[i];
        const title = item.title || item.label || "";
        const desc = item.desc || item.subLabel || "";
        const x = pos.x - circleSize / 2;
        const y = pos.y - circleSize / 2;
        const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, circleSize, circleSize);
        circle.getFill().setSolidFill(settings.primaryColor);
        circle.getBorder().setTransparent();
        setStyledText(circle, `${title}
${desc}`, {
          size: 14,
          bold: true,
          color: DEFAULT_THEME.colors.backgroundGray,
          align: SlidesApp.ParagraphAlignment.CENTER
        });
        try {
          circle.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          const textRange = circle.getText();
          if (title.length > 0 && desc.length > 0) {
          }
        } catch (e) {
        }
      });
    }
  }
  class ComparisonDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const leftTitle = data.leftTitle || "プランA";
      const rightTitle = data.rightTitle || "プランB";
      const leftItems = data.leftItems || [];
      const rightItems = data.rightItems || [];
      const gap = layout.pxToPt(60);
      const colWidth = (area.width - gap) / 2;
      const compareColors = generateCompareColors(settings.primaryColor);
      const headerH = layout.pxToPt(50);
      const itemSpacing = layout.pxToPt(16);
      const drawColumn = (x, title, items, accentColor) => {
        const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, area.top, layout.pxToPt(60), layout.pxToPt(4));
        line.getFill().setSolidFill(accentColor);
        line.getBorder().setTransparent();
        const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, area.top + layout.pxToPt(15), colWidth, headerH);
        setStyledText(titleBox, title, {
          size: 28,
          // Even Larger
          bold: true,
          color: DEFAULT_THEME.colors.textPrimary,
          align: SlidesApp.ParagraphAlignment.START
        });
        let currentY = area.top + headerH + layout.pxToPt(10);
        items.forEach((itemText) => {
          const bulletSize = layout.pxToPt(6);
          const bullet = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, currentY + layout.pxToPt(7), bulletSize, bulletSize);
          bullet.getFill().setSolidFill(DEFAULT_THEME.colors.neutralGray);
          bullet.getBorder().setTransparent();
          const textX = x + bulletSize + layout.pxToPt(8);
          const textW = colWidth - (bulletSize + layout.pxToPt(8));
          const itemH = layout.pxToPt(50);
          const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textX, currentY, textW, itemH);
          setStyledText(textBox, itemText, {
            size: 16,
            color: DEFAULT_THEME.colors.textPrimary,
            align: SlidesApp.ParagraphAlignment.START
          });
          try {
            textBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
          } catch (e) {
          }
          currentY += layout.pxToPt(25) + itemSpacing;
        });
      };
      drawColumn(area.left, leftTitle, leftItems, compareColors.left);
      drawColumn(area.left + colWidth + gap, rightTitle, rightItems, compareColors.right);
      const centerX = area.left + colWidth + gap / 2;
      const sepY = area.top + layout.pxToPt(20);
      const sepH = area.height - layout.pxToPt(40);
      const sepLine = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, centerX, sepY, centerX, sepY + sepH);
      sepLine.getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
      sepLine.setWeight(1);
    }
  }
  class StatsCompareDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const leftTitle = data.leftTitle || "導入前";
      const rightTitle = data.rightTitle || "導入後";
      const stats = data.stats || [];
      if (!stats.length) return;
      const compareColors = generateCompareColors(settings.primaryColor);
      const headerH = layout.pxToPt(45);
      const labelColW = area.width * 0.35;
      const valueColW = (area.width - labelColW) / 2;
      const leftHeaderX = area.left + labelColW;
      const leftHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftHeaderX, area.top, valueColW, headerH);
      leftHeader.getFill().setSolidFill(compareColors.left);
      leftHeader.getBorder().setTransparent();
      setStyledText(leftHeader, leftTitle, { size: 14, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });
      try {
        leftHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {
      }
      const rightHeaderX = area.left + labelColW + valueColW;
      const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightHeaderX, area.top, valueColW, headerH);
      rightHeader.getFill().setSolidFill(compareColors.right);
      rightHeader.getBorder().setTransparent();
      setStyledText(rightHeader, rightTitle, { size: 14, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });
      try {
        rightHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {
      }
      const availableHeight = area.height - headerH;
      const rowHeight = Math.min(layout.pxToPt(60), availableHeight / stats.length);
      let currentY = area.top + headerH;
      stats.forEach((stat, index) => {
        const label = stat.label || "";
        const leftValue = stat.leftValue || "";
        const rightValue = stat.rightValue || "";
        const trend = stat.trend || null;
        const rowBg = index % 2 === 0 ? DEFAULT_THEME.colors.backgroundGray : "#FFFFFF";
        const labelCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, currentY, labelColW, rowHeight);
        labelCell.getFill().setSolidFill(rowBg);
        labelCell.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
        setStyledText(labelCell, label, { size: DEFAULT_THEME.fonts.sizes.body, bold: true, align: SlidesApp.ParagraphAlignment.START });
        try {
          labelCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
        const leftCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftHeaderX, currentY, valueColW, rowHeight);
        leftCell.getFill().setSolidFill(rowBg);
        leftCell.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
        setStyledText(leftCell, leftValue, { size: DEFAULT_THEME.fonts.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER, color: compareColors.left });
        try {
          leftCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
        const rightCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightHeaderX, currentY, valueColW - (trend ? layout.pxToPt(40) : 0), rowHeight);
        rightCell.getFill().setSolidFill(rowBg);
        rightCell.getBorder().getLineFill().setSolidFill(DEFAULT_THEME.colors.faintGray);
        setStyledText(rightCell, rightValue, { size: DEFAULT_THEME.fonts.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER, color: compareColors.right });
        try {
          rightCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
        if (trend) {
          const trendX = rightHeaderX + valueColW - layout.pxToPt(35);
          const trendShape = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, trendX, currentY + rowHeight / 4, layout.pxToPt(25), layout.pxToPt(25));
          const isUp = trend.toLowerCase() === "up";
          const trendColor = isUp ? "#28a745" : "#dc3545";
          trendShape.getFill().setSolidFill(trendColor);
          trendShape.getBorder().setTransparent();
          setStyledText(trendShape, isUp ? "↑" : "↓", { size: 12, color: "#FFFFFF", bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
          try {
            trendShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {
          }
        }
        currentY += rowHeight;
      });
    }
  }
  class BarCompareDiagramRenderer {
    render(slide, data, area, settings, layout) {
      data.leftTitle || "導入前";
      data.rightTitle || "導入後";
      const stats = data.stats || [];
      if (!stats.length) return;
      generateCompareColors(settings.primaryColor);
      let maxValue = 0;
      stats.forEach((stat) => {
        const leftNum = parseFloat(String(stat.leftValue || "0").replace(/[^0-9.]/g, "")) || 0;
        const rightNum = parseFloat(String(stat.rightValue || "0").replace(/[^0-9.]/g, "")) || 0;
        maxValue = Math.max(maxValue, leftNum, rightNum);
      });
      if (maxValue === 0) maxValue = 100;
      const labelColW = area.width * 0.2;
      area.width * 0.6;
      area.width * 0.1;
      area.width * 0.1;
      const rowHeight = Math.min(layout.pxToPt(80), area.height / stats.length);
      const barHeight = layout.pxToPt(18);
      layout.pxToPt(4);
      let currentY = area.top;
      stats.forEach((stat, index) => {
        const label = stat.label || "";
        const leftValue = stat.leftValue || "";
        const rightValue = stat.rightValue || "";
        const trend = stat.trend || null;
        const leftNum = parseFloat(String(leftValue).replace(/[^0-9.]/g, "")) || 0;
        const rightNum = parseFloat(String(rightValue).replace(/[^0-9.]/g, "")) || 0;
        const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, currentY, labelColW, rowHeight);
        setStyledText(labelShape, label, {
          size: 14,
          bold: true,
          align: SlidesApp.ParagraphAlignment.START,
          color: DEFAULT_THEME.colors.textPrimary
        });
        try {
          labelShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
        const barLeft = area.left + labelColW + layout.pxToPt(20);
        const maxBarWidth = area.width - (labelColW + layout.pxToPt(80));
        const barGap2 = layout.pxToPt(2);
        const barY = currentY + (rowHeight - (barHeight * 2 + barGap2 * 2)) / 2;
        const leftBarW = leftNum / maxValue * maxBarWidth;
        if (leftBarW > 0) {
          const b = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barY, leftBarW, barHeight);
          b.getFill().setSolidFill(DEFAULT_THEME.colors.neutralGray);
          b.getBorder().setTransparent();
          const v = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + leftBarW + layout.pxToPt(5), barY - layout.pxToPt(2), layout.pxToPt(60), barHeight + layout.pxToPt(5));
          setStyledText(v, leftValue, { size: 10, color: DEFAULT_THEME.colors.neutralGray });
        }
        const rightBarW = rightNum / maxValue * maxBarWidth;
        if (rightBarW > 0) {
          const b = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barY + barHeight + barGap2, rightBarW, barHeight);
          b.getFill().setSolidFill(settings.primaryColor);
          b.getBorder().setTransparent();
          const v = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + rightBarW + layout.pxToPt(5), barY + barHeight + barGap2 - layout.pxToPt(2), layout.pxToPt(60), barHeight + layout.pxToPt(5));
          setStyledText(v, rightValue, { size: 10, bold: true, color: settings.primaryColor });
        }
        if (trend) {
          const trendX = area.left + area.width - layout.pxToPt(40);
          const isUp = trend.toLowerCase() === "up";
          const trendShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, trendX, currentY, layout.pxToPt(40), rowHeight);
          const color = isUp ? "#2E7D32" : "#C62828";
          const sym = isUp ? "↑" : "↓";
          setStyledText(trendShape, sym, {
            size: 20,
            bold: true,
            color,
            align: SlidesApp.ParagraphAlignment.CENTER
          });
          try {
            trendShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {
          }
        }
        if (index < stats.length - 1) {
          const lineY = currentY + rowHeight;
          const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, area.left, lineY, area.left + area.width, lineY);
          line.getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
          line.setWeight(0.5);
        }
        currentY += rowHeight;
      });
    }
  }
  class StepUpDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const items = data.items || [];
      if (!items.length) return;
      const count = items.length;
      const gap = layout.pxToPt(20);
      const validCount = Math.min(count, 5);
      const stepWidth = (area.width - gap * (validCount - 1)) / validCount;
      area.height / validCount;
      items.slice(0, validCount).forEach((item, i) => {
        const x = area.left + i * (stepWidth + gap);
        const totalRise = area.height * 0.6;
        const yStep = totalRise / (validCount - 1 || 1);
        const y = area.top + area.height - layout.pxToPt(100) - i * yStep;
        const lineW = stepWidth;
        const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, lineW, layout.pxToPt(4));
        line.getFill().setSolidFill(settings.primaryColor);
        line.getBorder().setTransparent();
        const numStr = String(i + 1).padEnd(2, "0");
        const numBoxH = layout.pxToPt(40);
        const numBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y - numBoxH + layout.pxToPt(5), lineW, numBoxH);
        setStyledText(numBox, numStr, {
          size: 32,
          bold: true,
          color: settings.primaryColor,
          align: SlidesApp.ParagraphAlignment.START
        });
        try {
          numBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
        } catch (e) {
        }
        const dropLine = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, y + layout.pxToPt(2), x, area.top + area.height);
        dropLine.getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
        dropLine.setDashStyle(SlidesApp.DashStyle.DOT);
        dropLine.setWeight(1);
        const titleContentY = y + layout.pxToPt(10);
        const textH = area.top + area.height - titleContentY;
        const titleHeight = layout.pxToPt(25);
        const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleContentY, stepWidth + gap / 2, titleHeight);
        const title = item.title || item.label || "";
        const desc = item.desc || item.description || item.text || "";
        setStyledText(textBox, `${title.toUpperCase()}`, {
          size: 14,
          bold: true,
          color: DEFAULT_THEME.colors.textPrimary
        });
        try {
          textBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        } catch (e) {
        }
        if (desc) {
          const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleContentY + titleHeight, stepWidth + gap / 2, textH - titleHeight);
          setStyledText(descBox, desc, {
            size: 12,
            color: DEFAULT_THEME.colors.textSmallFont
          });
          try {
            descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
          } catch (e) {
          }
        }
      });
    }
  }
  class LanesDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const lanes = data.lanes || [];
      const n = Math.max(1, lanes.length);
      const { laneGapPx, lanePadPx, laneTitleHeightPx, cardGapPx, cardMinHeightPx, cardMaxHeightPx, arrowHeightPx, arrowGapPx } = DEFAULT_THEME.diagram;
      const px = (p) => layout.pxToPt(p);
      const laneW = (area.width - px(laneGapPx) * (n - 1)) / n;
      const cardBoxes = [];
      for (let j = 0; j < n; j++) {
        const lane = lanes[j] || { title: "", items: [] };
        const left = area.left + j * (laneW + px(laneGapPx));
        const lt = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, area.top, laneW, px(laneTitleHeightPx));
        lt.getFill().setSolidFill(settings.primaryColor);
        lt.getBorder().getLineFill().setSolidFill(settings.primaryColor);
        setStyledText(lt, lane.title || "", { size: DEFAULT_THEME.fonts.sizes.laneTitle, bold: true, color: DEFAULT_THEME.colors.backgroundGray, align: SlidesApp.ParagraphAlignment.CENTER });
        try {
          lt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
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
          setStyledText(card, items[i] || "", { size: DEFAULT_THEME.fonts.sizes.body });
          try {
            card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {
          }
          cardBoxes[j][i] = {
            left: left + px(lanePadPx),
            top: cardTop,
            width: laneW - px(lanePadPx) * 2,
            height: cardH
          };
        }
      }
      const maxRows = Math.max(0, ...cardBoxes.map((a) => a ? a.length : 0));
      for (let j = 0; j < n - 1; j++) {
        for (let i = 0; i < maxRows; i++) {
          if (cardBoxes[j] && cardBoxes[j][i] && cardBoxes[j + 1] && cardBoxes[j + 1][i]) {
            drawArrowBetweenRects(slide, cardBoxes[j][i], cardBoxes[j + 1][i], px(arrowHeightPx), px(arrowGapPx), settings);
          }
        }
      }
    }
  }
  class FlowChartDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const steps = data.steps || data.items || [];
      if (!steps.length) return;
      const count = steps.length;
      const gap = 30;
      const boxWidth = (area.width - gap * (count - 1)) / count;
      const boxHeight = 80;
      const y = area.top + (area.height - boxHeight) / 2;
      steps.forEach((step, i) => {
        const x = area.left + i * (boxWidth + gap);
        const shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, boxWidth, boxHeight);
        shape.getFill().setSolidFill(DEFAULT_THEME.colors.backgroundGray);
        shape.getBorder().getLineFill().setSolidFill(settings.primaryColor);
        shape.getBorder().setWeight(2);
        setStyledText(shape, typeof step === "string" ? step : step.label || "", { size: DEFAULT_THEME.fonts.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER });
        try {
          shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
        if (i < count - 1) {
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
  }
  class KPIDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const items = data.items || [];
      if (!items.length) return;
      const cols = items.length > 4 ? 4 : items.length || 1;
      const gap = layout.pxToPt(40);
      const cardW = (area.width - gap * (cols - 1)) / cols;
      const cardH = layout.pxToPt(180);
      const y = area.top + (area.height - cardH) / 2 + layout.pxToPt(20);
      items.forEach((item, i) => {
        const x = area.left + i * (cardW + gap);
        if (i > 0) {
          const lineH = layout.pxToPt(100);
          const lineY = y + (cardH - lineH) / 2;
          const lineX = x - gap / 2;
          const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lineX, lineY, lineX, lineY + lineH);
          line.getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
          line.setWeight(1);
        }
        const labelH = layout.pxToPt(20);
        const labelBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + layout.pxToPt(5), cardW, labelH);
        setStyledText(labelBox, (item.label || "METRIC").toUpperCase(), {
          size: 11,
          // Slightly smaller
          color: DEFAULT_THEME.colors.neutralGray,
          align: SlidesApp.ParagraphAlignment.CENTER,
          bold: true
        });
        try {
          labelBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
        } catch (e) {
        }
        const valueH = layout.pxToPt(90);
        const valueBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + labelH, cardW, valueH);
        const valStr = String(item.value || "0");
        let fontSize = 72;
        if (valStr.length > 4) fontSize = 60;
        if (valStr.length > 6) fontSize = 48;
        if (valStr.length > 10) fontSize = 36;
        setStyledText(valueBox, valStr, {
          size: fontSize,
          bold: true,
          color: settings.primaryColor,
          align: SlidesApp.ParagraphAlignment.CENTER,
          fontType: "lato"
        });
        try {
          valueBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        } catch (e) {
        }
        if (item.change || item.status) {
          const statusH = layout.pxToPt(30);
          const statusBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + labelH + valueH, cardW, statusH);
          let color = DEFAULT_THEME.colors.neutralGray;
          let prefix = "";
          if (item.status === "good") {
            color = "#2E7D32";
            prefix = "↑ ";
          }
          if (item.status === "bad") {
            color = "#C62828";
            prefix = "↓ ";
          }
          setStyledText(statusBox, prefix + (item.change || ""), {
            size: 14,
            bold: true,
            color,
            align: SlidesApp.ParagraphAlignment.CENTER
          });
        }
      });
    }
  }
  class TableDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const headers = data.headers || [];
      const rows = data.rows || [];
      const numCols = Math.max(
        headers.length,
        rows.length > 0 && Array.isArray(rows[0]) ? rows[0].length : 0
      );
      const numRows = rows.length + (headers.length ? 1 : 0);
      Logger.log(`Table: ${numRows} rows x ${numCols} cols, headers: ${JSON.stringify(headers)}`);
      Logger.log(`Table area: left=${area.left}, top=${area.top}, width=${area.width}, height=${area.height}`);
      if (numRows === 0 || numCols === 0) return;
      const table = slide.insertTable(numRows, numCols);
      table.setLeft(area.left);
      table.setTop(area.top);
      table.setWidth(area.width);
      const theme = layout.getTheme();
      for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
          const cell = table.getCell(r, c);
          cell.getFill().setTransparent();
          cell.getBorderRight().setTransparent();
          cell.getBorderLeft().setTransparent();
          cell.getBorderTop().setTransparent();
          const borderBottom = cell.getBorderBottom();
          borderBottom.setSolidFill(theme.colors.ghostGray);
          borderBottom.setWeight(1);
        }
      }
      let rowIndex = 0;
      if (headers.length) {
        for (let c = 0; c < numCols; c++) {
          const cell = table.getCell(0, c);
          cell.getText().setText(headers[c] || "");
          const style = cell.getText().getTextStyle();
          style.setFontFamily(theme.fonts.family);
          style.setBold(true);
          style.setFontSize(14);
          style.setForegroundColor(settings.primaryColor);
          cell.getBorderBottom().setSolidFill(settings.primaryColor);
          cell.getBorderBottom().setWeight(3);
          try {
            cell.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
          } catch (e) {
          }
        }
        rowIndex++;
      }
      rows.forEach((row, rIdx) => {
        var _a;
        for (let c = 0; c < numCols; c++) {
          const cell = table.getCell(rowIndex, c);
          cell.getText().setText(String((_a = row[c]) != null ? _a : ""));
          const rowStyle = cell.getText().getTextStyle();
          rowStyle.setFontFamily(theme.fonts.family);
          rowStyle.setFontSize(12);
          rowStyle.setForegroundColor(theme.colors.textPrimary);
          try {
            cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {
          }
        }
        rowIndex++;
      });
    }
  }
  class FAQDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const items = data.items || data.points || [];
      const parsedItems = [];
      if (items.length && typeof items[0] === "string") {
        let currentQ = "";
        items.forEach((str) => {
          if (str.startsWith("Q:") || str.startsWith("Q.")) currentQ = str;
          else if (str.startsWith("A:") || str.startsWith("A.")) parsedItems.push({ q: currentQ, a: str });
        });
      } else {
        items.forEach((it) => parsedItems.push(it));
      }
      if (!parsedItems.length) return;
      const gap = layout.pxToPt(30);
      const itemH = (area.height - gap * (parsedItems.length - 1)) / parsedItems.length;
      parsedItems.forEach((item, i) => {
        const y = area.top + i * (itemH + gap);
        const qStr = (item.q || "").replace(/^[QA][:. ]+/, "");
        const aStr = (item.a || "").replace(/^[QA][:. ]+/, "");
        const qInd = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, y, layout.pxToPt(30), layout.pxToPt(30));
        setStyledText(qInd, "Q.", { size: 16, bold: true, color: settings.primaryColor });
        const qBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + layout.pxToPt(30), y, area.width - layout.pxToPt(30), layout.pxToPt(40));
        setStyledText(qBox, qStr, { size: 14, bold: true, color: DEFAULT_THEME.colors.textPrimary });
        const aY = y + layout.pxToPt(30);
        const aInd = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, aY, layout.pxToPt(30), layout.pxToPt(30));
        setStyledText(aInd, "A.", { size: 16, bold: true, color: DEFAULT_THEME.colors.neutralGray });
        const aBoxHasHeight = itemH - layout.pxToPt(40);
        if (aBoxHasHeight > 10) {
          const aBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + layout.pxToPt(30), aY, area.width - layout.pxToPt(30), aBoxHasHeight);
          setStyledText(aBox, aStr, { size: 12, color: DEFAULT_THEME.colors.textPrimary });
          try {
            aBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
          } catch (e) {
          }
        }
        if (i < parsedItems.length - 1) {
          const lineY = y + itemH + gap / 2;
          const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, area.left, lineY, area.left + area.width, lineY);
          line.getLineFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
          line.setWeight(0.5);
        }
      });
    }
  }
  class QuoteDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const text = data.text || data.points && data.points[0] || "";
      const author = data.author || data.points && data.points[1] || "";
      const quoteSize = layout.pxToPt(200);
      const quoteMark = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + layout.pxToPt(20), area.top - layout.pxToPt(40), quoteSize, quoteSize);
      setStyledText(quoteMark, "“", {
        size: 200,
        color: "#F0F0F0",
        // Very faint
        fontType: "georgia",
        // Serif
        bold: true
      });
      quoteMark.sendToBack();
      const contentW = area.width * 0.9;
      const contentX = area.left + (area.width - contentW) / 2;
      const textTop = area.top + layout.pxToPt(60);
      const quoteBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, textTop, contentW, layout.pxToPt(160));
      setStyledText(quoteBox, text, {
        size: 32,
        bold: false,
        color: DEFAULT_THEME.colors.textPrimary,
        align: SlidesApp.ParagraphAlignment.CENTER,
        fontType: "georgia"
      });
      try {
        quoteBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
      } catch (e) {
      }
      const lineW = layout.pxToPt(40);
      const lineX = area.left + (area.width - lineW) / 2;
      const lineY = textTop + layout.pxToPt(165);
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lineX, lineY, lineX + lineW, lineY);
      line.getLineFill().setSolidFill(settings.primaryColor);
      line.setWeight(2);
      if (author) {
        const authorBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, lineY + layout.pxToPt(5), contentW, layout.pxToPt(30));
        setStyledText(authorBox, author, {
          size: 14,
          align: SlidesApp.ParagraphAlignment.CENTER,
          color: DEFAULT_THEME.colors.neutralGray,
          bold: true
          // Small but bold
        });
        try {
          authorBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        } catch (e) {
        }
      }
    }
  }
  class ProgressDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const items = data.items || [];
      if (!items.length) return;
      const rowH = layout.pxToPt(50);
      const gap = layout.pxToPt(15);
      const startY = area.top + (area.height - items.length * (rowH + gap)) / 2;
      items.forEach((item, i) => {
        const y = startY + i * (rowH + gap);
        const labelW = layout.pxToPt(150);
        const barAreaW = area.width - labelW - layout.pxToPt(60);
        const labelBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, y, labelW, rowH);
        setStyledText(labelBox, item.label || "", { size: 14, bold: true, align: SlidesApp.ParagraphAlignment.END });
        try {
          labelBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
        const barBg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left + labelW + layout.pxToPt(20), y + rowH / 3, barAreaW, rowH / 3);
        barBg.getFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
        barBg.getBorder().setTransparent();
        const percent = Math.min(100, Math.max(0, parseInt(item.percent || 0)));
        if (percent > 0) {
          const barFg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left + labelW + layout.pxToPt(20), y + rowH / 3, barAreaW * (percent / 100), rowH / 3);
          barFg.getFill().setSolidFill(settings.primaryColor);
          barFg.getBorder().setTransparent();
        }
        const valBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + labelW + barAreaW + layout.pxToPt(30), y, layout.pxToPt(50), rowH);
        setStyledText(valBox, `${percent}%`, { size: 14, color: DEFAULT_THEME.colors.neutralGray });
        try {
          valBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {
        }
      });
    }
  }
  class ImageTextDiagramRenderer {
    render(slide, data, area, settings, layout) {
      const imageUrl = data.image;
      const points = data.points || [];
      const gap = layout.pxToPt(20);
      const halfW = (area.width - gap) / 2;
      const isImageLeft = data.imagePosition !== "right";
      const imgX = isImageLeft ? area.left : area.left + halfW + gap;
      const txtX = isImageLeft ? area.left + halfW + gap : area.left;
      if (imageUrl) {
        try {
          const blob = insertImageFromUrlOrFileId(imageUrl);
          let img;
          if (blob) {
            img = slide.insertImage(blob);
          } else if (typeof imageUrl === "string" && imageUrl.startsWith("http")) {
            img = slide.insertImage(imageUrl);
          }
          if (img) {
            const scale = Math.min(halfW / img.getWidth(), area.height / img.getHeight());
            const w = img.getWidth() * scale;
            const h = img.getHeight() * scale;
            img.setWidth(w).setHeight(h).setLeft(imgX + (halfW - w) / 2).setTop(area.top + (area.height - h) / 2);
          } else {
            throw new Error("Image insert failed");
          }
        } catch (e) {
          Logger.log("Image insert failed: " + e);
          const ph = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, imgX, area.top, halfW, area.height);
          ph.getFill().setSolidFill(DEFAULT_THEME.colors.ghostGray);
          setStyledText(ph, "Image Placeholder", { align: SlidesApp.ParagraphAlignment.CENTER });
        }
      }
      const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, txtX, area.top, halfW, area.height);
      const textContent = points.join("\n");
      setStyledText(textBox, textContent, { size: DEFAULT_THEME.fonts.sizes.body });
    }
  }
  class DiagramRendererFactory {
    static getRenderer(type) {
      const normalizedType = type.toLowerCase();
      Logger.log(`[Factory] Checking renderer for type: '${type}' (normalized: '${normalizedType}')`);
      if (normalizedType.includes("cards") || normalizedType.includes("headercards")) {
        Logger.log("[Factory] Matched CardsDiagramRenderer");
        return new CardsDiagramRenderer();
      }
      if (normalizedType.includes("timeline")) {
        Logger.log("[Factory] Matched TimelineDiagramRenderer");
        return new TimelineDiagramRenderer();
      }
      if (normalizedType.includes("process")) {
        Logger.log("[Factory] Matched ProcessDiagramRenderer");
        return new ProcessDiagramRenderer();
      }
      if (normalizedType.includes("cycle")) {
        Logger.log("[Factory] Matched CycleDiagramRenderer");
        return new CycleDiagramRenderer();
      }
      if (normalizedType.includes("pyramid")) {
        Logger.log("[Factory] Matched PyramidDiagramRenderer");
        return new PyramidDiagramRenderer();
      }
      if (normalizedType.includes("triangle")) {
        Logger.log("[Factory] Matched TriangleDiagramRenderer");
        return new TriangleDiagramRenderer();
      }
      if (normalizedType.includes("compare") || normalizedType.includes("comparison") || normalizedType.includes("kaizen")) {
        if (normalizedType.includes("stats")) {
          Logger.log("[Factory] Matched StatsCompareDiagramRenderer");
          return new StatsCompareDiagramRenderer();
        }
        if (normalizedType.includes("bar")) {
          Logger.log("[Factory] Matched BarCompareDiagramRenderer");
          return new BarCompareDiagramRenderer();
        }
        Logger.log("[Factory] Matched ComparisonDiagramRenderer");
        return new ComparisonDiagramRenderer();
      }
      if (normalizedType.includes("stepup") || normalizedType.includes("stair")) {
        Logger.log("[Factory] Matched StepUpDiagramRenderer");
        return new StepUpDiagramRenderer();
      }
      if (normalizedType.includes("lanes") || normalizedType.includes("diagram")) {
        Logger.log("[Factory] Matched LanesDiagramRenderer");
        return new LanesDiagramRenderer();
      }
      if (normalizedType.includes("flow")) {
        Logger.log("[Factory] Matched FlowChartDiagramRenderer");
        return new FlowChartDiagramRenderer();
      }
      if (normalizedType.includes("kpi")) {
        Logger.log("[Factory] Matched KPIDiagramRenderer");
        return new KPIDiagramRenderer();
      }
      if (normalizedType.includes("table")) {
        Logger.log("[Factory] Matched TableDiagramRenderer");
        return new TableDiagramRenderer();
      }
      if (normalizedType.includes("faq")) {
        Logger.log("[Factory] Matched FAQDiagramRenderer");
        return new FAQDiagramRenderer();
      }
      if (normalizedType.includes("quote")) {
        Logger.log("[Factory] Matched QuoteDiagramRenderer");
        return new QuoteDiagramRenderer();
      }
      if (normalizedType.includes("progress")) {
        Logger.log("[Factory] Matched ProgressDiagramRenderer");
        return new ProgressDiagramRenderer();
      }
      if (normalizedType.includes("image") || normalizedType.includes("imagetext")) {
        Logger.log("[Factory] Matched ImageTextDiagramRenderer");
        return new ImageTextDiagramRenderer();
      }
      Logger.log("[Factory] No matching renderer found.");
      return null;
    }
  }
  class GasDiagramSlideGenerator {
    constructor(creditImageBlob) {
      this.creditImageBlob = creditImageBlob;
    }
    generate(slide, data, layout, pageNum, settings, imageUpdateOption = "update") {
      Logger.log(`[GasDiagramSlideGenerator] Generating Diagram Slide: ${data.layout || data.type}`);
      Logger.log(`[Debug] ColorUtils loaded: ${!!ColorUtils}`);
      const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) || slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
      if (titlePlaceholder) {
        try {
          titlePlaceholder.asShape().getText().setText(data.title || "");
        } catch (e) {
          Logger.log(`Warning: Title placeholder found but text could not be set. ${e}`);
        }
      }
      const type = (data.type || data.layout || "").toLowerCase();
      Logger.log("Generating Diagram Slide: " + type);
      const placeholders = slide.getPlaceholders();
      const getPlaceholderTypeSafe = (p) => {
        try {
          const shape = p.asShape();
          if (shape) {
            return shape.getPlaceholderType();
          }
        } catch (e) {
        }
        return null;
      };
      const targetPlaceholder = placeholders.find((p) => getPlaceholderTypeSafe(p) === SlidesApp.PlaceholderType.BODY) || placeholders.find((p) => getPlaceholderTypeSafe(p) === SlidesApp.PlaceholderType.OBJECT) || placeholders.find((p) => getPlaceholderTypeSafe(p) === SlidesApp.PlaceholderType.PICTURE);
      let workArea;
      if (targetPlaceholder) {
        workArea = {
          left: targetPlaceholder.getLeft(),
          top: targetPlaceholder.getTop(),
          width: targetPlaceholder.getWidth(),
          height: targetPlaceholder.getHeight()
        };
      } else {
        const typeKey = `${type}Slide`;
        const areaRect = layout.getRect(`${typeKey}.area`) || layout.getRect("contentSlide.body");
        workArea = areaRect;
      }
      Logger.log(`WorkArea for ${type}: left=${workArea.left}, top=${workArea.top}, width=${workArea.width}, height=${workArea.height}`);
      if (targetPlaceholder) {
        try {
          targetPlaceholder.remove();
        } catch (e) {
          Logger.log("Warning: Could not remove target placeholder: " + e);
        }
      }
      const elementsBefore = slide.getPageElements().map((e) => e.getObjectId());
      try {
        const renderer = DiagramRendererFactory.getRenderer(type);
        if (renderer) {
          renderer.render(slide, data, workArea, settings, layout);
        } else {
          Logger.log("Diagram logic not implemented for type: " + type);
        }
      } catch (e) {
        Logger.log(`ERROR in drawing ${type}: ${e}`);
      }
      const newElements = slide.getPageElements().filter((e) => {
        if (elementsBefore.includes(e.getObjectId())) return false;
        try {
          const shape = e.asShape();
          if (shape) {
            const placeholderType = shape.getPlaceholderType();
            if (placeholderType === SlidesApp.PlaceholderType.TITLE || placeholderType === SlidesApp.PlaceholderType.SUBTITLE || placeholderType === SlidesApp.PlaceholderType.CENTERED_TITLE) {
              return false;
            }
          }
        } catch (e2) {
        }
        return true;
      });
      let generatedGroup = null;
      if (newElements.length > 1) {
        try {
          generatedGroup = slide.group(newElements);
          Logger.log(`Grouped ${newElements.length} content elements for ${type}`);
        } catch (e) {
          Logger.log(`Warning: Could not group elements: ${e}`);
        }
      } else if (newElements.length === 1) {
        generatedGroup = newElements[0];
      }
      if (generatedGroup) {
        const groupWidth = generatedGroup.getWidth();
        const groupHeight = generatedGroup.getHeight();
        const workAreaCenterX = workArea.left + workArea.width / 2;
        const workAreaCenterY = workArea.top + workArea.height / 2;
        const newLeft = workAreaCenterX - groupWidth / 2;
        const newTop = workAreaCenterY - groupHeight / 2;
        generatedGroup.setLeft(newLeft);
        generatedGroup.setTop(newTop);
        Logger.log(`Centered Group: left=${newLeft}, top=${newTop} (Area Center: ${workAreaCenterX}, ${workAreaCenterY})`);
      }
      addFooter(slide, layout, pageNum, settings, this.creditImageBlob);
    }
  }
  class GasSlideRepository {
    createPresentation(presentation, templateId, destinationId, settingsOverride) {
      const slidesApp = SlidesApp;
      const driveApp = DriveApp;
      let pres;
      if (destinationId) {
        try {
          Logger.log(`Using existing destination ID: ${destinationId}`);
          pres = slidesApp.openById(destinationId);
        } catch (e) {
          Logger.log(`Error opening destination: ${e.toString()}`);
          throw new Error(`Failed to open destination presentation with ID: ${destinationId}. Details: ${e.message}`);
        }
      } else if (templateId) {
        try {
          Logger.log(`Attempting to access template ID: ${templateId}`);
          const templateFile = driveApp.getFileById(templateId);
          const newFile = templateFile.makeCopy(presentation.title);
          Logger.log(`Template copied. New File ID: ${newFile.getId()}`);
          pres = slidesApp.openById(newFile.getId());
        } catch (e) {
          Logger.log(`Error accessing/copying template: ${e.toString()}`);
          throw new Error(`Failed to access or copy template with ID: ${templateId}. Ensure the ID is correct and the script has permission to access it. Details: ${e.message}`);
        }
      } else {
        pres = slidesApp.create(presentation.title);
      }
      const templateSlides = pres.getSlides();
      const pageWidth = pres.getPageWidth();
      const pageHeight = pres.getPageHeight();
      const layoutManager = new LayoutManager(pageWidth, pageHeight);
      const titleGenerator = new GasTitleSlideGenerator(null);
      const sectionGenerator = new GasSectionSlideGenerator(null);
      const contentGenerator = new GasContentSlideGenerator(null);
      const diagramGenerator = new GasDiagramSlideGenerator(null);
      const theme = layoutManager.getTheme();
      const settings = {
        primaryColor: theme.colors.primary,
        enableGradient: false,
        showTitleUnderline: true,
        showBottomBar: true,
        showDateColumn: true,
        showPageNumber: true,
        primary_color: theme.colors.primary,
        text_primary: theme.colors.textPrimary,
        background_gray: theme.colors.backgroundGray,
        card_bg: theme.colors.cardBg,
        ghost_gray: theme.colors.ghostGray,
        ...settingsOverride && settingsOverride.colors ? {
          primaryColor: settingsOverride.colors.primary,
          primary_color: settingsOverride.colors.primary,
          text_primary: settingsOverride.colors.text
        } : {}
      };
      presentation.slides.forEach((slideModel, index) => {
        var _a;
        const commonData = {
          title: slideModel.title.value,
          subtitle: slideModel.subtitle,
          date: (/* @__PURE__ */ new Date()).toLocaleDateString(),
          points: slideModel.content.items,
          // Map content items to points for Content/Agenda
          content: slideModel.content.items,
          ...slideModel.rawData
          // Merge all extra data from JSON
        };
        let slideLayout;
        const layouts = pres.getLayouts();
        if (index === 0) {
          Logger.log("Available Layouts: " + layouts.map((l) => l.getLayoutName()).join(", "));
        }
        const findLayout = (name) => {
          return layouts.find((l) => l.getLayoutName().toUpperCase() === name.toUpperCase());
        };
        const layoutType = (slideModel.layout || "content").toUpperCase();
        if (layoutType === "TITLE") {
          slideLayout = findLayout("TITLE") || layouts[0];
        } else if (layoutType === "SECTION") {
          slideLayout = findLayout("SECTION_HEADER") || findLayout("SECTION ONLY") || findLayout("SECTION TITLE_AND_DESCRIPTION") || layouts[1];
        } else if (layoutType === "CONTENT" || layoutType === "AGENDA") {
          slideLayout = findLayout("TITLE_AND_BODY") || findLayout("TITLE_AND_TWO_COLUMNS") || layouts[2];
        } else {
          slideLayout = findLayout("TITLE_AND_BODY") || findLayout("TITLE_ONLY") || layouts[layouts.length - 1];
        }
        if (!slideLayout) {
          slideLayout = layouts[0];
        }
        Logger.log(`Slide ${index + 1} (${slideModel.layout}): Using Layout '${slideLayout.getLayoutName()}'`);
        const slide = pres.appendSlide(slideLayout);
        const rawType = (((_a = slideModel.rawData) == null ? void 0 : _a.type) || "").toLowerCase();
        Logger.log(`Dispatching Slide ${index + 1}: LayoutType=${layoutType}, RawType=${rawType}`);
        if (layoutType === "TITLE") {
          titleGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
        } else if (layoutType === "SECTION" || rawType === "section") {
          sectionGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
        } else if (rawType.includes("timeline") || rawType.includes("process") || rawType.includes("cycle") || rawType.includes("triangle") || rawType.includes("pyramid") || rawType.includes("diagram") || rawType.includes("compare") || rawType.includes("stepup") || rawType.includes("flowchart") || rawType.includes("cards") || rawType.includes("kpi") || rawType.includes("table") || rawType.includes("faq") || rawType.includes("progress") || rawType.includes("quote") || rawType.includes("imagetext")) {
          diagramGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
        } else {
          contentGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
        }
        if (slideModel.notes) {
          try {
            const notesPage = slide.getNotesPage();
            notesPage.getSpeakerNotesShape().getText().setText(slideModel.notes);
          } catch (e) {
            Logger.log(`Warning: Could not set speaker notes. ${e}`);
          }
        }
      });
      const initialCount = templateSlides.length;
      if (templateId) {
        for (let i = 0; i < initialCount; i++) {
          if (pres.getSlides().length > presentation.slides.length) {
            pres.getSlides()[0].remove();
          }
        }
      } else {
        if (pres.getSlides().length > presentation.slides.length) {
          pres.getSlides()[0].remove();
        }
      }
      return pres.getUrl();
    }
  }
  function generateSlides(data) {
    try {
      Logger.log("Library Call: generateSlides with data: " + JSON.stringify(data));
      const request = {
        title: data.title || "無題のプレゼンテーション",
        templateId: data.templateId,
        // Optional ID for template
        destinationId: data.destinationId,
        // Optional ID for existing destination
        settings: data.settings,
        slides: data.slides.map((s) => ({
          ...s,
          // Spread all valid properties from source
          title: s.title,
          subtitle: s.subtitle || s.subhead,
          // Handle alias if inconsistent
          content: s.content || s.points || [],
          // Partial alias support
          layout: s.layout || s.type,
          // Map type to layout
          notes: s.notes
        }))
      };
      const repository = new GasSlideRepository();
      const service = new PresentationApplicationService(repository);
      const slideUrl = service.createPresentation(request);
      return {
        success: true,
        url: slideUrl
      };
    } catch (error) {
      Logger.log("Library Error: " + error.toString());
      return {
        success: false,
        error: error.message || error.toString()
      };
    }
  }
  function doGet(e) {
    return ContentService.createTextOutput("Slide Generator API is running. Authorization successful.");
  }
  function doPost(e) {
    try {
      const postData = JSON.parse(e.postData.contents);
      if (postData.action === "test") {
        return createJsonResponse({
          success: true,
          message: "POST Connection Successful",
          received: postData
        });
      }
      const data = postData.json || postData;
      Logger.log("Incoming Request Data: " + JSON.stringify(data));
      const result = generateSlides(data);
      return createJsonResponse(result);
    } catch (error) {
      Logger.log("API Error: " + error.toString());
      return createJsonResponse({
        success: false,
        error: error.message || error.toString()
      });
    }
  }
  function createJsonResponse(data) {
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  }
  global.generateSlides = generateSlides;
  global.doPost = doPost;
  global.doGet = doGet;
})();
