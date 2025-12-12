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
  const BASE_POSITIONS = {
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
    positions: BASE_POSITIONS,
    backgroundImages: {
      title: "",
      closing: "",
      section: "",
      main: ""
    }
  };
  const THEME_GREEN = {
    ...BASE_THEME_OPTS,
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
    }
  };
  const THEME_BLUE = {
    ...BASE_THEME_OPTS,
    colors: {
      primary: "#3498DB",
      // Standard Blue
      deepPrimary: "#2C3E50",
      // Dark Blue
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
  const THEME_RED = {
    ...BASE_THEME_OPTS,
    colors: {
      primary: "#E74C3C",
      // Red
      deepPrimary: "#C0392B",
      // Dark Red
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
  const THEME_GRAYSCALE = {
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
  const DEFAULT_THEME = THEME_GREEN;
  const AVAILABLE_THEMES = {
    "Green": THEME_GREEN,
    "Blue": THEME_BLUE,
    "Red": THEME_RED,
    "Grayscale": THEME_GRAYSCALE
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
  class RequestFactory {
    static createSlide(pageId, layoutId = "TITLE_AND_BODY", isPredefined = true) {
      return {
        createSlide: {
          objectId: pageId,
          slideLayoutReference: isPredefined ? { predefinedLayout: layoutId } : { layoutId }
        }
      };
    }
    static createTextBox(pageId, objectId, text, leftOpt, topOpt, widthOpt, heightOpt) {
      const requests = [];
      requests.push({
        createShape: {
          objectId,
          shapeType: "TEXT_BOX",
          elementProperties: {
            pageObjectId: pageId,
            size: {
              width: { magnitude: widthOpt, unit: "PT" },
              height: { magnitude: heightOpt, unit: "PT" }
            },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: leftOpt,
              translateY: topOpt,
              unit: "PT"
            }
          }
        }
      });
      if (text) {
        requests.push({
          insertText: {
            objectId,
            text
          }
        });
      }
      return requests;
    }
    static updateTextStyle(objectId, style, selectionStart, selectionEnd, rowIndex, colIndex) {
      const fields = [];
      const textStyle = {};
      if (style.bold !== void 0) {
        textStyle.bold = style.bold;
        fields.push("bold");
      }
      if (style.fontSize !== void 0) {
        textStyle.fontSize = { magnitude: style.fontSize, unit: "PT" };
        fields.push("fontSize");
      }
      if (style.color !== void 0) {
        textStyle.foregroundColor = {
          opaqueColor: { themeColor: style.color.startsWith("#") ? "DARK1" : style.color }
        };
        const rgb = RequestFactory.toRgbColor(style.color);
        if (rgb) {
          textStyle.foregroundColor = {
            opaqueColor: { rgbColor: rgb }
          };
        }
        fields.push("foregroundColor");
      }
      if (style.fontFamily !== void 0) {
        textStyle.fontFamily = style.fontFamily;
        fields.push("fontFamily");
      }
      const request = {
        updateTextStyle: {
          objectId,
          style: textStyle,
          fields: fields.join(",")
        }
      };
      if (selectionStart !== void 0 && selectionEnd !== void 0) {
        request.updateTextStyle.textRange = {
          type: "FIXED_RANGE",
          startIndex: selectionStart,
          endIndex: selectionStart + selectionEnd
        };
      }
      if (rowIndex !== void 0 && colIndex !== void 0) {
        request.updateTextStyle.cellLocation = {
          rowIndex,
          columnIndex: colIndex
        };
      }
      return request;
    }
    static createShape(pageId, objectId, shapeType, left, top, width, height) {
      return {
        createShape: {
          objectId,
          shapeType,
          elementProperties: {
            pageObjectId: pageId,
            size: {
              width: { magnitude: width, unit: "PT" },
              height: { magnitude: height, unit: "PT" }
            },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: left,
              translateY: top,
              unit: "PT"
            }
          }
        }
      };
    }
    static createLine(pageId, objectId, x1, y1, x2, y2) {
      const width = Math.abs(x2 - x1) || 1;
      const height = Math.abs(y2 - y1) || 1;
      const left = Math.min(x1, x2);
      const top = Math.min(y1, y2);
      return {
        createLine: {
          objectId,
          lineCategory: "STRAIGHT",
          elementProperties: {
            pageObjectId: pageId,
            size: {
              width: { magnitude: width, unit: "PT" },
              height: { magnitude: height, unit: "PT" }
            },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: left,
              translateY: top,
              unit: "PT"
            }
          }
        }
      };
    }
    static updateShapeProperties(objectId, fillHex, borderHex, borderWeightPt, contentAlignment) {
      const fields = [];
      const shapeProperties = {};
      if (fillHex !== void 0 && fillHex !== null) {
        if (fillHex === "TRANSPARENT") {
          shapeProperties.shapeBackgroundFill = { propertyState: "NOT_RENDERED" };
        } else {
          const rgb = RequestFactory.toRgbColor(fillHex);
          if (rgb) {
            shapeProperties.shapeBackgroundFill = {
              solidFill: { color: { rgbColor: rgb } }
            };
          }
        }
        fields.push("shapeBackgroundFill");
      }
      if (borderHex !== void 0 || borderWeightPt !== void 0) {
        const outline = {};
        if (borderHex === "TRANSPARENT") {
          outline.propertyState = "NOT_RENDERED";
        } else {
          const rgb = RequestFactory.toRgbColor(borderHex);
          if (rgb) {
            outline.outlineFill = { solidFill: { color: { rgbColor: rgb } } };
          }
        }
        if (borderWeightPt !== void 0) {
          outline.weight = { magnitude: borderWeightPt, unit: "PT" };
        }
        shapeProperties.outline = outline;
        fields.push("outline");
      }
      if (contentAlignment) {
        shapeProperties.contentAlignment = contentAlignment;
        fields.push("contentAlignment");
      }
      return {
        updateShapeProperties: {
          objectId,
          shapeProperties,
          fields: fields.join(",")
        }
      };
    }
    static updateParagraphStyle(objectId, alignment, rowIndex, colIndex) {
      const req = {
        updateParagraphStyle: {
          objectId,
          style: {
            alignment
          },
          fields: "alignment"
        }
      };
      if (rowIndex !== void 0 && colIndex !== void 0) {
        req.updateParagraphStyle.cellLocation = {
          rowIndex,
          columnIndex: colIndex
        };
      }
      return req;
    }
    static createTable(pageId, objectId, rows, cols, left, top, width, height) {
      return {
        createTable: {
          objectId,
          rows,
          columns: cols,
          elementProperties: {
            pageObjectId: pageId,
            size: {
              width: { magnitude: width, unit: "PT" },
              height: { magnitude: height, unit: "PT" }
            },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: left,
              translateY: top,
              unit: "PT"
            }
          }
        }
      };
    }
    static insertText(objectId, text, rowIndex, colIndex) {
      const req = {
        insertText: {
          objectId,
          text
        }
      };
      if (rowIndex !== void 0 && colIndex !== void 0) {
        req.insertText.cellLocation = {
          rowIndex,
          columnIndex: colIndex
        };
      }
      return req;
    }
    static createImage(pageId, objectId, url, left, top, width, height) {
      return {
        createImage: {
          objectId,
          url,
          elementProperties: {
            pageObjectId: pageId,
            size: {
              width: { magnitude: width, unit: "PT" },
              height: { magnitude: height, unit: "PT" }
            },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: left,
              translateY: top,
              unit: "PT"
            }
          }
        }
      };
    }
    static toRgbColor(hex) {
      if (!hex || typeof hex !== "string" || !hex.startsWith("#") || hex.length < 7) {
        return void 0;
      }
      const r = parseInt(hex.substring(1, 3), 16) / 255;
      const g = parseInt(hex.substring(3, 5), 16) / 255;
      const b = parseInt(hex.substring(5, 7), 16) / 255;
      return { red: r, green: g, blue: b };
    }
    static groupObjects(pageId, groupObjectId, childrenObjectIds) {
      return {
        groupObjects: {
          groupObjectId,
          childrenObjectIds
        }
      };
    }
  }
  class GasTitleSlideGenerator {
    constructor(creditImageBlob) {
      this.creditImageBlob = creditImageBlob;
    }
    // API Version
    generate(slideId, data, layout, pageNum, settings) {
      const requests = [];
      const theme = layout.getTheme();
      if (data.title) {
        const titleId = slideId + "_TITLE";
        const titleRect = layout.getRect("titleSlide.title") || { left: 50, top: 200, width: 860, height: 100 };
        requests.push(RequestFactory.createShape(
          slideId,
          titleId,
          "TEXT_BOX",
          titleRect.left,
          titleRect.top,
          titleRect.width,
          titleRect.height
        ));
        requests.push({ insertText: { objectId: titleId, text: data.title } });
        requests.push(RequestFactory.updateTextStyle(titleId, {
          fontSize: 48,
          bold: true,
          color: settings.primaryColor || "#000000",
          fontFamily: theme.fonts.family
          // Use theme font if available
        }));
        requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: "CENTER" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: "BOTTOM" }, fields: "contentAlignment" } });
      }
      const text = data.subtitle || data.date;
      if (text) {
        const subId = slideId + "_SUBTITLE";
        const subRect = layout.getRect("titleSlide.subtitle") || { left: 50, top: 320, width: 860, height: 50 };
        requests.push(RequestFactory.createShape(
          slideId,
          subId,
          "TEXT_BOX",
          subRect.left,
          subRect.top,
          subRect.width,
          subRect.height
        ));
        requests.push({ insertText: { objectId: subId, text } });
        requests.push(RequestFactory.updateTextStyle(subId, {
          fontSize: 24,
          color: theme.colors.neutralGray || "#666666",
          fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: subId, style: { alignment: "CENTER" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: subId, shapeProperties: { contentAlignment: "TOP" }, fields: "contentAlignment" } });
      }
      return requests;
    }
  }
  class GasSectionSlideGenerator {
    constructor(creditImageBlob) {
      this.creditImageBlob = creditImageBlob;
      this.sectionCounter = 0;
    }
    // API Version
    generate(slideId, data, layout, pageNum, settings) {
      const requests = [];
      const theme = layout.getTheme();
      this.sectionCounter++;
      const parsedNum = (() => {
        if (Number.isFinite(data.sectionNo)) return Number(data.sectionNo);
        const m = String(data.title || "").match(/^\s*(\d+)[\.．]/);
        return m ? Number(m[1]) : this.sectionCounter;
      })();
      const num = String(parsedNum).padStart(2, "0");
      const ghostRect = layout.getRect("sectionSlide.ghostNum") || { left: 680, top: 320, width: 280, height: 200 };
      const ghostId = slideId + "_GHOST";
      requests.push(RequestFactory.createShape(
        slideId,
        ghostId,
        "TEXT_BOX",
        ghostRect.left,
        ghostRect.top,
        ghostRect.width,
        ghostRect.height
      ));
      requests.push({ insertText: { objectId: ghostId, text: num } });
      requests.push(RequestFactory.updateTextStyle(ghostId, {
        bold: true,
        fontSize: theme.fonts.sizes.ghostNum || 250,
        color: theme.colors.ghostGray,
        fontFamily: theme.fonts.family
      }));
      requests.push({ updateParagraphStyle: { objectId: ghostId, style: { alignment: "CENTER" }, fields: "alignment" } });
      const titleId = slideId + "_TITLE";
      if (data.title) {
        const titleRect = layout.getRect("sectionSlide.title") || { left: 40, top: 200, width: 700, height: 100 };
        requests.push(RequestFactory.createShape(
          slideId,
          titleId,
          "TEXT_BOX",
          titleRect.left,
          titleRect.top,
          titleRect.width,
          titleRect.height
        ));
        requests.push({ insertText: { objectId: titleId, text: data.title } });
        requests.push(RequestFactory.updateTextStyle(titleId, {
          bold: true,
          fontSize: theme.fonts.sizes.sectionTitle || 52,
          color: theme.colors.primary,
          fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: "START" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: "BOTTOM" }, fields: "contentAlignment" } });
      }
      return requests;
    }
  }
  class GasContentSlideGenerator {
    constructor(creditImageBlob) {
      this.creditImageBlob = creditImageBlob;
    }
    generate(slideId, data, layout, pageNum, settings) {
      const requests = [];
      const theme = layout.getTheme();
      if (data.title) {
        const titleId = slideId + "_TITLE";
        const titleRect = layout.getRect("contentSlide.title") || { left: 30, top: 20, width: 900, height: 60 };
        requests.push(RequestFactory.createShape(slideId, titleId, "TEXT_BOX", titleRect.left, titleRect.top, titleRect.width, titleRect.height));
        requests.push({ insertText: { objectId: titleId, text: data.title } });
        requests.push(RequestFactory.updateTextStyle(titleId, {
          fontSize: 32,
          bold: true,
          color: settings.primaryColor,
          fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: "START" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: "BOTTOM" }, fields: "contentAlignment" } });
      }
      if (data.subhead) {
        const subId = slideId + "_SUBHEAD";
        const subRect = layout.getRect("contentSlide.subhead") || { left: 30, top: 90, width: 900, height: 30 };
        requests.push(RequestFactory.createShape(slideId, subId, "TEXT_BOX", subRect.left, subRect.top, subRect.width, subRect.height));
        requests.push({ insertText: { objectId: subId, text: data.subhead } });
        requests.push(RequestFactory.updateTextStyle(subId, {
          fontSize: 18,
          bold: true,
          color: theme.colors.neutralGray,
          fontFamily: theme.fonts.family
        }));
        subRect.height;
      }
      const points = Array.isArray(data.points) ? data.points : [];
      const columns = data.columns || [];
      const isTwoColumn = columns.length === 2 || data.twoColumn;
      const bodyRect = layout.getRect("contentSlide.body") || { left: 30, top: 130, width: 900, height: 380 };
      const topY = bodyRect.top + (data.subhead ? 10 : 0);
      if (isTwoColumn) {
        const gap = 30;
        const colW = (bodyRect.width - gap) / 2;
        const leftPoints = columns[0] || points.slice(0, Math.ceil(points.length / 2));
        const rightPoints = columns[1] || points.slice(Math.ceil(points.length / 2));
        requests.push(...this.createBulletRequests(slideId, slideId + "_BODY_0", leftPoints, bodyRect.left, topY, colW, bodyRect.height, theme));
        requests.push(...this.createBulletRequests(slideId, slideId + "_BODY_1", rightPoints, bodyRect.left + colW + gap, topY, colW, bodyRect.height, theme));
      } else {
        requests.push(...this.createBulletRequests(slideId, slideId + "_BODY", points, bodyRect.left, topY, bodyRect.width, bodyRect.height, theme));
      }
      return requests;
    }
    createBulletRequests(slideId, objectId, points, x, y, w, h, theme) {
      const reqs = [];
      if (!points || points.length === 0) return reqs;
      reqs.push(RequestFactory.createShape(slideId, objectId, "TEXT_BOX", x, y, w, h));
      const fullText = points.join("\n");
      reqs.push({ insertText: { objectId, text: fullText } });
      reqs.push(RequestFactory.updateTextStyle(objectId, {
        fontSize: 18,
        color: theme.colors.textPrimary,
        fontFamily: theme.fonts.family
      }));
      reqs.push({
        createParagraphBullets: {
          objectId,
          textRange: { type: "ALL" },
          bulletPreset: "BULLET_DISC_CIRCLE_SQUARE"
        }
      });
      reqs.push({ updateParagraphStyle: { objectId, style: { alignment: "START" }, fields: "alignment" } });
      reqs.push({ updateShapeProperties: { objectId, shapeProperties: { contentAlignment: "TOP" }, fields: "contentAlignment" } });
      return reqs;
    }
  }
  class BatchTextStyleUtils {
    /**
     * Replicates the behavior of setStyledText but triggers Requests.
     * Supports basic Markdown '**bold**' parsing.
     */
    static setText(slideId, objectId, text, baseStyle, theme) {
      const requests = [];
      const parts = (text || "").split("**");
      let cleanText = "";
      const boldRanges = [];
      let isBold = false;
      parts.forEach((part) => {
        if (isBold && part.length > 0) {
          boldRanges.push({ start: cleanText.length, length: part.length });
        }
        cleanText += part;
        isBold = !isBold;
      });
      requests.push(RequestFactory.insertText(objectId, cleanText));
      const styleReq = {
        bold: baseStyle.bold || false,
        fontSize: baseStyle.size || baseStyle.fontSize || theme.fonts.sizes.body,
        color: baseStyle.color || theme.colors.textPrimary,
        // Default color
        fontFamily: theme.fonts.family
      };
      requests.push(RequestFactory.updateTextStyle(objectId, styleReq, 0, cleanText.length));
      boldRanges.forEach((r) => {
        requests.push(RequestFactory.updateTextStyle(objectId, { bold: true }, r.start, r.length));
      });
      if (baseStyle.align) {
        let alignStr = "START";
        const align = String(baseStyle.align).toUpperCase();
        if (align.includes("CENTER") || align.includes("MIDDLE")) alignStr = "CENTER";
        else if (align.includes("END") || align.includes("RIGHT")) alignStr = "END";
        else if (align.includes("JUSTIFY")) alignStr = "JUSTIFIED";
        requests.push(RequestFactory.updateParagraphStyle(objectId, alignStr));
      }
      return requests;
    }
  }
  class CardsDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const items = data.items || [];
      if (!items.length) return requests;
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
        const baseId = slideId + "_CARD_" + i;
        if (hasHeader) {
          const barH = layout.pxToPt(4);
          const barId = baseId + "_BAR";
          requests.push(RequestFactory.createShape(slideId, barId, "RECTANGLE", x, y, cardW, barH));
          requests.push(RequestFactory.updateShapeProperties(barId, settings.primaryColor, "TRANSPARENT"));
          const numStr = String(i + 1).padStart(2, "0");
          const numId = baseId + "_NUM";
          requests.push(RequestFactory.createShape(slideId, numId, "TEXT_BOX", x, y + layout.pxToPt(6), cardW, layout.pxToPt(20)));
          requests.push(...BatchTextStyleUtils.setText(slideId, numId, numStr, {
            size: 14,
            bold: true,
            color: theme.colors.neutralGray,
            align: "END"
          }, theme));
          const titleTop = y + layout.pxToPt(6);
          const titleH = layout.pxToPt(30);
          const titleId = baseId + "_TITLE";
          requests.push(RequestFactory.createShape(slideId, titleId, "TEXT_BOX", x, titleTop, cardW, titleH));
          requests.push(RequestFactory.updateShapeProperties(titleId, null, null, null, "TOP"));
          requests.push(...BatchTextStyleUtils.setText(slideId, titleId, title, {
            size: 18,
            bold: true,
            color: theme.colors.textPrimary,
            align: "START"
          }, theme));
          const descTop = titleTop + titleH;
          const descH = cardH - (descTop - y);
          if (descH > 20) {
            const descId = baseId + "_DESC";
            requests.push(RequestFactory.createShape(slideId, descId, "TEXT_BOX", x, descTop, cardW, descH));
            requests.push(RequestFactory.updateShapeProperties(descId, null, null, null, "TOP"));
            requests.push(...BatchTextStyleUtils.setText(slideId, descId, desc, {
              size: 13,
              color: typeof theme.colors.textSmallFont === "string" ? theme.colors.textSmallFont : "#424242",
              align: "START"
            }, theme));
          }
        } else {
          const dotSize = layout.pxToPt(6);
          const dotId = baseId + "_DOT";
          const dotY = y + layout.pxToPt(8);
          requests.push(RequestFactory.createShape(slideId, dotId, "ELLIPSE", x, dotY, dotSize, dotSize));
          requests.push(RequestFactory.updateShapeProperties(dotId, settings.primaryColor, "TRANSPARENT"));
          const contentX = x + dotSize + layout.pxToPt(10);
          const contentW = cardW - (dotSize + layout.pxToPt(10));
          const titleH = layout.pxToPt(30);
          const titleId = baseId + "_TITLE";
          requests.push(RequestFactory.createShape(slideId, titleId, "TEXT_BOX", contentX, y, contentW, titleH));
          requests.push(RequestFactory.updateShapeProperties(titleId, null, null, null, "TOP"));
          requests.push(...BatchTextStyleUtils.setText(slideId, titleId, title, {
            size: 16,
            bold: true,
            color: theme.colors.textPrimary,
            align: "START"
          }, theme));
          const descTop = y + titleH - layout.pxToPt(5);
          const descH = cardH - (descTop - y);
          if (descH > 20) {
            const descId = baseId + "_DESC";
            requests.push(RequestFactory.createShape(slideId, descId, "TEXT_BOX", contentX, descTop, contentW, descH));
            requests.push(RequestFactory.updateShapeProperties(descId, null, null, null, "TOP"));
            requests.push(...BatchTextStyleUtils.setText(slideId, descId, desc, {
              size: 13,
              color: typeof theme.colors.textSmallFont === "string" ? theme.colors.textSmallFont : "#424242",
              align: "START"
            }, theme));
          }
        }
      });
      return requests;
    }
  }
  class TimelineDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const milestones = data.milestones || data.items || [];
      if (milestones.length === 0) return requests;
      const centerY = area.top + area.height / 2;
      const startX = area.left + 50;
      const endX = area.left + area.width - 50;
      const axisId = slideId + "_TL_AXIS";
      requests.push(RequestFactory.createLine(slideId, axisId, startX, centerY, endX, centerY));
      requests.push({
        updateLineProperties: {
          objectId: axisId,
          lineProperties: {
            weight: { magnitude: 4, unit: "PT" },
            lineFill: { solidFill: { color: { rgbColor: { red: 0.8, green: 0.8, blue: 0.8 } } } },
            // neutralGray #CCCCCC
            dashStyle: "SOLID"
          },
          fields: "weight,lineFill,dashStyle"
        }
      });
      const count = milestones.length;
      const usableWidth = endX - startX;
      const gap = usableWidth / (count + 1);
      milestones.forEach((m, i) => {
        const x = startX + gap * (i + 1);
        const isTop = i % 2 === 0;
        const cardH = 80;
        const cardW = 140;
        const connLen = 50;
        const dotR = 24;
        const unique = `_${i}`;
        const dotId = slideId + "_TL_DOT" + unique;
        const connId = slideId + "_TL_CONN" + unique;
        const cardId = slideId + "_TL_CARD" + unique;
        const connY1 = centerY;
        const connY2 = isTop ? centerY - connLen : centerY + connLen;
        requests.push(RequestFactory.createShape(slideId, dotId, "ELLIPSE", x - dotR / 2, centerY - dotR / 2, dotR, dotR));
        requests.push(RequestFactory.updateShapeProperties(dotId, settings.primaryColor, "#FFFFFF", 3));
        requests.push(RequestFactory.createLine(slideId, connId, x, connY1, x, connY2));
        requests.push({
          updateLineProperties: {
            objectId: connId,
            lineProperties: {
              weight: { magnitude: 2, unit: "PT" },
              lineFill: { solidFill: { color: { rgbColor: { red: 0.7, green: 0.7, blue: 0.7 } } } },
              dashStyle: "DOT"
            },
            fields: "weight,lineFill,dashStyle"
          }
        });
        const cardY = isTop ? connY2 - cardH : connY2;
        const cardX = x - cardW / 2;
        requests.push(RequestFactory.createShape(slideId, cardId, "ROUND_RECTANGLE", cardX, cardY, cardW, cardH));
        requests.push(RequestFactory.updateShapeProperties(cardId, "#FFFFFF", theme.colors.ghostGray));
        const date = m.date || m.year || "";
        const label = m.label || m.title || m.text || "";
        const fullText = `${date}
${label}`;
        requests.push({ insertText: { objectId: cardId, text: fullText } });
        if (date) {
          requests.push(RequestFactory.updateTextStyle(cardId, {
            bold: true,
            fontSize: 14,
            color: settings.primaryColor,
            fontFamily: theme.fonts.family
          }, 0, date.length));
        }
        if (label) {
          const start = date.length + 1;
          requests.push(RequestFactory.updateTextStyle(cardId, {
            bold: false,
            fontSize: 12,
            color: theme.colors.textPrimary || "#333333",
            fontFamily: theme.fonts.family
          }, start, label.length));
        }
        requests.push({ updateParagraphStyle: { objectId: cardId, style: { alignment: "CENTER" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: cardId, shapeProperties: { contentAlignment: "MIDDLE" }, fields: "contentAlignment" } });
      });
      return requests;
    }
  }
  class ProcessDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const steps = data.steps || data.items || [];
      if (!steps.length) return requests;
      const count = steps.length;
      const gap = -15;
      const itemWidth = (area.width - gap * (count - 1)) / count;
      const itemHeight = Math.min(area.height, 150);
      const y = area.top + (area.height - itemHeight) / 2;
      steps.forEach((step, i) => {
        const x = area.left + i * (itemWidth + gap);
        const unique = `_${i}`;
        const shapeId = slideId + "_PR_CHEV_" + unique;
        requests.push(RequestFactory.createShape(slideId, shapeId, "CHEVRON", x, y, itemWidth, itemHeight));
        const bgColor = i % 2 === 0 ? theme.colors.primary : theme.colors.deepPrimary || theme.colors.primary;
        requests.push({
          updateShapeProperties: {
            objectId: shapeId,
            shapeProperties: {
              shapeBackgroundFill: { solidFill: { color: { rgbColor: RequestFactory.toRgbColor(bgColor) || { red: 0, green: 0, blue: 0 } } } },
              outline: { propertyState: "NOT_RENDERED" },
              contentAlignment: "MIDDLE"
            },
            fields: "shapeBackgroundFill,outline,contentAlignment"
          }
        });
        const cleanText = String(step || "").replace(/^\s*\d+[\.\s]*/, "");
        const fullText = `${i + 1}
${cleanText}`;
        requests.push({ insertText: { objectId: shapeId, text: fullText } });
        requests.push(RequestFactory.updateTextStyle(shapeId, {
          fontSize: 32,
          bold: true,
          color: "#FFFFFF",
          fontFamily: theme.fonts.family
        }, 0, String(i + 1).length));
        requests.push(RequestFactory.updateTextStyle(shapeId, {
          fontSize: 14,
          bold: false,
          color: "#FFFFFF",
          // White text on colored background
          fontFamily: theme.fonts.family
        }, String(i + 1).length + 1, cleanText.length + 1));
        requests.push({
          updateParagraphStyle: {
            objectId: shapeId,
            style: { alignment: "CENTER" },
            fields: "alignment"
          }
        });
      });
      return requests;
    }
  }
  class CycleDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const items = data.items || [];
      if (!items.length) return requests;
      const centerX = area.left + area.width / 2;
      const centerY = area.top + area.height / 2;
      const orbitRadius = Math.min(area.width, area.height) * 0.35;
      const count = items.length;
      const angleStep = 2 * Math.PI / count;
      const startAngle = -Math.PI / 2;
      items.forEach((item, i) => {
        const angle = startAngle + i * angleStep;
        const x = centerX + Math.cos(angle) * orbitRadius;
        const y = centerY + Math.sin(angle) * orbitRadius;
        const connId = slideId + `_CYC_CONN_${i}`;
        requests.push(RequestFactory.createLine(slideId, connId, centerX, centerY, x, y));
        requests.push({
          updateLineProperties: {
            objectId: connId,
            lineProperties: {
              lineFill: { solidFill: { color: { rgbColor: { red: 0.8, green: 0.8, blue: 0.8 } } } },
              // Light Gray
              weight: { magnitude: 4, unit: "PT" },
              // Thicker
              dashStyle: "SOLID"
            },
            fields: "lineFill,weight,dashStyle"
          }
        });
      });
      const centerR = 140;
      const centerId = slideId + "_CYC_CENTER";
      const shadowId = slideId + "_CYC_CENTER_SHADOW";
      requests.push(RequestFactory.createShape(slideId, shadowId, "ELLIPSE", centerX - centerR / 2 + 3, centerY - centerR / 2 + 3, centerR, centerR));
      requests.push(RequestFactory.updateShapeProperties(shadowId, "#DDDDDD", "TRANSPARENT"));
      requests.push(RequestFactory.createShape(slideId, centerId, "ELLIPSE", centerX - centerR / 2, centerY - centerR / 2, centerR, centerR));
      requests.push(RequestFactory.updateShapeProperties(centerId, settings.primaryColor, "#FFFFFF", 3));
      const centerText = data.centerText || data.title || "CYCLE";
      requests.push({ insertText: { objectId: centerId, text: centerText } });
      requests.push(RequestFactory.updateTextStyle(centerId, {
        fontSize: 20,
        bold: true,
        color: "#FFFFFF",
        fontFamily: theme.fonts.family
      }));
      requests.push({ updateParagraphStyle: { objectId: centerId, style: { alignment: "CENTER" }, fields: "alignment" } });
      requests.push({ updateShapeProperties: { objectId: centerId, shapeProperties: { contentAlignment: "MIDDLE" }, fields: "contentAlignment" } });
      items.forEach((item, i) => {
        const angle = startAngle + i * angleStep;
        const x = centerX + Math.cos(angle) * orbitRadius;
        const y = centerY + Math.sin(angle) * orbitRadius;
        const unique = `_CYC_${i}`;
        const bubbleId = slideId + unique + "_BUBBLE";
        const labelId = slideId + unique + "_LABEL";
        const bubbleR = 60;
        const bShadowId = bubbleId + "_SHADOW";
        requests.push(RequestFactory.createShape(slideId, bShadowId, "ELLIPSE", x - bubbleR / 2 + 2, y - bubbleR / 2 + 2, bubbleR, bubbleR));
        requests.push(RequestFactory.updateShapeProperties(bShadowId, "#DDDDDD", "TRANSPARENT"));
        requests.push(RequestFactory.createShape(slideId, bubbleId, "ELLIPSE", x - bubbleR / 2, y - bubbleR / 2, bubbleR, bubbleR));
        requests.push(RequestFactory.updateShapeProperties(bubbleId, "#FFFFFF", settings.primaryColor, 2));
        const numStr = String(i + 1).padStart(2, "0");
        requests.push({ insertText: { objectId: bubbleId, text: numStr } });
        requests.push(RequestFactory.updateTextStyle(bubbleId, {
          fontSize: 24,
          bold: true,
          color: settings.primaryColor,
          fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: bubbleId, style: { alignment: "CENTER" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: bubbleId, shapeProperties: { contentAlignment: "MIDDLE" }, fields: "contentAlignment" } });
        const isRight = x > centerX;
        const textW = 140;
        const textH = 70;
        const offsetX = isRight ? bubbleR / 2 + 5 : -bubbleR / 2 - textW - 5;
        const textX = x + offsetX;
        const textY = y - textH / 2;
        const label = typeof item === "string" ? item : item.label || item.title || "";
        requests.push(RequestFactory.createShape(slideId, labelId, "TEXT_BOX", textX, textY, textW, textH));
        requests.push({ insertText: { objectId: labelId, text: label } });
        requests.push(RequestFactory.updateTextStyle(labelId, {
          fontSize: 14,
          bold: true,
          color: theme.colors.textPrimary || "#333333",
          fontFamily: theme.fonts.family
        }));
        requests.push({
          updateParagraphStyle: {
            objectId: labelId,
            style: { alignment: isRight ? "START" : "END" },
            fields: "alignment"
          }
        });
        requests.push({ updateShapeProperties: { objectId: labelId, shapeProperties: { contentAlignment: "MIDDLE" }, fields: "contentAlignment" } });
      });
      return requests;
    }
  }
  class PyramidDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const levels = data.levels || data.items || [];
      if (!levels.length) return requests;
      const count = levels.length;
      const pyramidH = Math.min(area.height, 400);
      const pyramidW = Math.min(area.width * 0.6, 500);
      const centerX = area.left + area.width / 2;
      const topY = area.top + (area.height - pyramidH) / 2;
      const triId = slideId + "_PYR_MAIN";
      requests.push(RequestFactory.createShape(slideId, triId, "TRIANGLE", centerX - pyramidW / 2, topY, pyramidW, pyramidH));
      requests.push(RequestFactory.updateShapeProperties(triId, settings.primaryColor, "TRANSPARENT"));
      const gap = 4;
      const levelHeight = (pyramidH - gap * (count - 1)) / count;
      for (let i = 1; i < count; i++) {
        const lineY = topY + i * (pyramidH / count);
        const ratio = i / count;
        const widthAtY = pyramidW * ratio;
        const lineId = slideId + `_PYR_LINE_${i}`;
        const lineW = widthAtY - 4;
        const lineH = 2;
        const lineLeft = centerX - lineW / 2;
        requests.push(RequestFactory.createShape(slideId, lineId, "RECTANGLE", lineLeft, lineY, lineW, lineH));
        requests.push(RequestFactory.updateShapeProperties(lineId, "#FFFFFF", "TRANSPARENT"));
      }
      levels.forEach((level, index) => {
        const y = topY + index * (pyramidH / count) + levelHeight / 2 - 10;
        const depthRatio = (index + 0.5) / count;
        const widthAtDepth = pyramidW * depthRatio;
        const pyramidEdgeX = centerX + widthAtDepth / 2;
        const lineStartX = pyramidEdgeX + 10;
        const lineEndX = lineStartX + 30;
        const textX = lineEndX + 10;
        const connId = slideId + `_PYR_CONN_${index}`;
        requests.push(RequestFactory.createLine(slideId, connId, lineStartX, y + 10, lineEndX, y + 10));
        requests.push({
          updateLineProperties: {
            objectId: connId,
            lineProperties: {
              lineFill: { solidFill: { color: { rgbColor: { red: 0.2, green: 0.2, blue: 0.2 } } } },
              // Dark Gray
              weight: { magnitude: 1, unit: "PT" }
            },
            fields: "lineFill,weight"
          }
        });
        const textId = slideId + `_PYR_TXT_${index}`;
        const title = level.label || level.title || "";
        const sub = level.subLabel || level.desc || "";
        const fullText = `${title}
${sub}`;
        requests.push(RequestFactory.createShape(slideId, textId, "TEXT_BOX", textX, y - 25, 200, 60));
        requests.push({ insertText: { objectId: textId, text: fullText } });
        if (title) {
          requests.push(RequestFactory.updateTextStyle(textId, {
            bold: true,
            fontSize: 16,
            color: settings.text_primary || "#333333",
            fontFamily: layout.getTheme().fonts.family
          }, 0, title.length));
        }
        if (sub) {
          const subStart = title.length + 1;
          requests.push(RequestFactory.updateTextStyle(textId, {
            bold: false,
            fontSize: 12,
            color: settings.ghost_gray || "#666666",
            fontFamily: layout.getTheme().fonts.family
          }, subStart, sub.length));
        }
        requests.push({
          updateShapeProperties: {
            objectId: textId,
            shapeProperties: { contentAlignment: "MIDDLE" },
            fields: "contentAlignment"
          }
        });
      });
      return requests;
    }
  }
  class TriangleDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const items = data.items || [];
      if (!items.length) return requests;
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
      const triangleId = slideId + "_TRI_CENTER";
      requests.push(RequestFactory.createShape(slideId, triangleId, "TRIANGLE", centerX - radius, centerY - radius, radius * 2, radius * 2));
      requests.push(RequestFactory.updateShapeProperties(triangleId, theme.colors.faintGray || "#EEEEEE", "TRANSPARENT"));
      itemsToDraw.forEach((item, i) => {
        const pos = positions[i];
        const title = item.title || item.label || "";
        const desc = item.desc || item.subLabel || "";
        const x = pos.x - circleSize / 2;
        const y = pos.y - circleSize / 2;
        const circleId = slideId + "_TRI_ITEM_" + i;
        requests.push(RequestFactory.createShape(slideId, circleId, "ELLIPSE", x, y, circleSize, circleSize));
        requests.push(RequestFactory.updateShapeProperties(circleId, settings.primaryColor, "TRANSPARENT", 0, "MIDDLE"));
        const fullText = `${title}
${desc}`;
        requests.push(RequestFactory.insertText(circleId, fullText));
        if (title.length > 0) {
          requests.push(RequestFactory.updateTextStyle(circleId, {
            fontSize: 14,
            bold: true,
            color: theme.colors.backgroundGray,
            fontFamily: theme.fonts.family
          }, 0, title.length));
        }
        if (desc.length > 0) {
          requests.push(RequestFactory.updateTextStyle(circleId, {
            fontSize: 12,
            bold: false,
            color: theme.colors.backgroundGray,
            fontFamily: theme.fonts.family
          }, title.length + 1, desc.length));
        }
        requests.push(RequestFactory.updateParagraphStyle(circleId, "CENTER"));
      });
      return requests;
    }
  }
  class ComparisonDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const leftTitle = data.leftTitle || "Plan A";
      const rightTitle = data.rightTitle || "Plan B";
      const leftItems = data.leftItems || [];
      const rightItems = data.rightItems || [];
      const gap = 60;
      const colWidth = (area.width - gap) / 2;
      const headerH = 70;
      const drawColumn = (x, title, items, suffix) => {
        const headerId = slideId + "_CMP_HEAD_" + suffix;
        const bodyId = slideId + "_CMP_BODY_" + suffix;
        requests.push(RequestFactory.createShape(slideId, headerId, "ROUND_RECTANGLE", x, area.top, colWidth, headerH));
        requests.push(RequestFactory.updateShapeProperties(headerId, settings.primaryColor, "TRANSPARENT"));
        requests.push({ insertText: { objectId: headerId, text: title } });
        requests.push(RequestFactory.updateTextStyle(headerId, {
          fontSize: 24,
          bold: true,
          color: "#FFFFFF",
          fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: headerId, style: { alignment: "CENTER" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: headerId, shapeProperties: { contentAlignment: "MIDDLE" }, fields: "contentAlignment" } });
        const bodyH = area.height - headerH - 10;
        const bodyY = area.top + headerH + 10;
        requests.push(RequestFactory.createShape(slideId, bodyId, "ROUND_RECTANGLE", x, bodyY, colWidth, bodyH));
        requests.push(RequestFactory.updateShapeProperties(bodyId, theme.colors.faintGray || "#F8F9FA", "TRANSPARENT"));
        let currentY = bodyY + 20;
        const itemGap = 15;
        items.forEach((itemText, i) => {
          const unique = `_${suffix}_${i}`;
          const iconId = slideId + "_CMP_ICON" + unique;
          const textId = slideId + "_CMP_TXT" + unique;
          const iconSize = 24;
          requests.push(RequestFactory.createShape(slideId, iconId, "TEXT_BOX", x + 20, currentY, iconSize, iconSize));
          requests.push({ insertText: { objectId: iconId, text: "✔" } });
          requests.push(RequestFactory.updateTextStyle(iconId, {
            fontSize: 18,
            bold: true,
            color: settings.primaryColor,
            fontFamily: theme.fonts.family
            // Use symbol font if needed, but standard usually has check
          }));
          requests.push({ updateParagraphStyle: { objectId: iconId, style: { alignment: "CENTER" }, fields: "alignment" } });
          const textX = x + 20 + iconSize + 10;
          const textW = colWidth - (20 + iconSize + 10 + 20);
          const itemH = 40;
          requests.push(RequestFactory.createShape(slideId, textId, "TEXT_BOX", textX, currentY, textW, itemH));
          requests.push({ insertText: { objectId: textId, text: itemText } });
          requests.push(RequestFactory.updateTextStyle(textId, {
            fontSize: 16,
            color: theme.colors.textPrimary || "#333333",
            fontFamily: theme.fonts.family
          }));
          requests.push({ updateParagraphStyle: { objectId: textId, style: { alignment: "START" }, fields: "alignment" } });
          currentY += itemH + itemGap;
        });
      };
      drawColumn(area.left, leftTitle, leftItems, "L");
      drawColumn(area.left + colWidth + gap, rightTitle, rightItems, "R");
      return requests;
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
  function darkenColor(color, amount) {
    const rgb = hexToRgb(color);
    if (!rgb) return color;
    const darken = (c) => Math.max(0, Math.round(c * (1 - amount)));
    const newR = darken(rgb.r);
    const newG = darken(rgb.g);
    const newB = darken(rgb.b);
    return `#${newR.toString(16).padStart(2, "0")}${newG.toString(16).padStart(2, "0")}${newB.toString(16).padStart(2, "0")}`;
  }
  function generateCompareColors(baseColor) {
    return {
      left: darkenColor(baseColor, 0.3),
      right: baseColor
    };
  }
  class StatsCompareDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const leftTitle = data.leftTitle || "導入前";
      const rightTitle = data.rightTitle || "導入後";
      const stats = data.stats || [];
      if (!stats.length) return requests;
      const compareColors = generateCompareColors(settings.primaryColor);
      const headerH = layout.pxToPt(45);
      const labelColW = area.width * 0.35;
      const valueColW = (area.width - labelColW) / 2;
      const leftHeaderX = area.left + labelColW;
      const lhId = slideId + "_STAT_LH";
      requests.push(RequestFactory.createShape(slideId, lhId, "RECTANGLE", leftHeaderX, area.top, valueColW, headerH));
      requests.push(RequestFactory.updateShapeProperties(lhId, compareColors.left, "TRANSPARENT", 0, "MIDDLE"));
      requests.push(...BatchTextStyleUtils.setText(slideId, lhId, leftTitle, { size: 14, bold: true, color: theme.colors.backgroundGray, align: "CENTER" }, theme));
      const rightHeaderX = area.left + labelColW + valueColW;
      const rhId = slideId + "_STAT_RH";
      requests.push(RequestFactory.createShape(slideId, rhId, "RECTANGLE", rightHeaderX, area.top, valueColW, headerH));
      requests.push(RequestFactory.updateShapeProperties(rhId, compareColors.right, "TRANSPARENT", 0, "MIDDLE"));
      requests.push(...BatchTextStyleUtils.setText(slideId, rhId, rightTitle, { size: 14, bold: true, color: theme.colors.backgroundGray, align: "CENTER" }, theme));
      const availableHeight = area.height - headerH;
      const rowHeight = Math.min(layout.pxToPt(60), availableHeight / stats.length);
      let currentY = area.top + headerH;
      stats.forEach((stat, index) => {
        const label = stat.label || "";
        const leftValue = stat.leftValue || "";
        const rightValue = stat.rightValue || "";
        const trend = stat.trend || null;
        const rowBg = index % 2 === 0 ? theme.colors.backgroundGray : "#FFFFFF";
        const baseId = slideId + "_STAT_R" + index;
        const lId = baseId + "_L";
        requests.push(RequestFactory.createShape(slideId, lId, "RECTANGLE", area.left, currentY, labelColW, rowHeight));
        requests.push(RequestFactory.updateShapeProperties(lId, rowBg, theme.colors.faintGray, 1, "MIDDLE"));
        requests.push(...BatchTextStyleUtils.setText(slideId, lId, label, { size: theme.fonts.sizes.body, bold: true, align: "START" }, theme));
        const lvId = baseId + "_LV";
        requests.push(RequestFactory.createShape(slideId, lvId, "RECTANGLE", leftHeaderX, currentY, valueColW, rowHeight));
        requests.push(RequestFactory.updateShapeProperties(lvId, rowBg, theme.colors.faintGray, 1, "MIDDLE"));
        requests.push(...BatchTextStyleUtils.setText(slideId, lvId, leftValue, { size: theme.fonts.sizes.body, align: "CENTER", color: compareColors.left }, theme));
        const rvId = baseId + "_RV";
        requests.push(RequestFactory.createShape(slideId, rvId, "RECTANGLE", rightHeaderX, currentY, valueColW - (trend ? layout.pxToPt(40) : 0), rowHeight));
        requests.push(RequestFactory.updateShapeProperties(rvId, rowBg, theme.colors.faintGray, 1, "MIDDLE"));
        requests.push(...BatchTextStyleUtils.setText(slideId, rvId, rightValue, { size: theme.fonts.sizes.body, align: "CENTER", color: compareColors.right }, theme));
        if (trend) {
          const trendX = rightHeaderX + valueColW - layout.pxToPt(35);
          const tId = baseId + "_TR";
          const isUp = String(trend).toLowerCase() === "up";
          const trendColor = isUp ? "#28a745" : "#dc3545";
          requests.push(RequestFactory.createShape(slideId, tId, "ELLIPSE", trendX, currentY + rowHeight / 4, layout.pxToPt(25), layout.pxToPt(25)));
          requests.push(RequestFactory.updateShapeProperties(tId, trendColor, "TRANSPARENT", 0, "MIDDLE"));
          requests.push(...BatchTextStyleUtils.setText(slideId, tId, isUp ? "↑" : "↓", { size: 12, color: "#FFFFFF", bold: true, align: "CENTER" }, theme));
        }
        currentY += rowHeight;
      });
      return requests;
    }
  }
  class BarCompareDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      data.leftTitle || "導入前";
      data.rightTitle || "導入後";
      const stats = data.stats || [];
      if (!stats.length) return requests;
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
      const rowHeight = Math.min(layout.pxToPt(80), area.height / stats.length);
      const barHeight = layout.pxToPt(18);
      let currentY = area.top;
      stats.forEach((stat, index) => {
        const label = stat.label || "";
        const leftValue = stat.leftValue || "";
        const rightValue = stat.rightValue || "";
        const trend = stat.trend || null;
        const leftNum = parseFloat(String(leftValue).replace(/[^0-9.]/g, "")) || 0;
        const rightNum = parseFloat(String(rightValue).replace(/[^0-9.]/g, "")) || 0;
        const baseId = slideId + "_BAR_R" + index;
        const labelId = baseId + "_LBL";
        requests.push(RequestFactory.createShape(slideId, labelId, "TEXT_BOX", area.left, currentY, labelColW, rowHeight));
        requests.push(...BatchTextStyleUtils.setText(slideId, labelId, label, {
          size: 14,
          bold: true,
          align: "START",
          color: theme.colors.textPrimary
        }, theme));
        requests.push(RequestFactory.updateShapeProperties(labelId, null, null, null, "MIDDLE"));
        const barLeft = area.left + labelColW + layout.pxToPt(20);
        const maxBarWidth = area.width - (labelColW + layout.pxToPt(80));
        const barGap = layout.pxToPt(2);
        const barY = currentY + (rowHeight - (barHeight * 2 + barGap * 2)) / 2;
        const leftBarW = leftNum / maxValue * maxBarWidth;
        if (leftBarW > 0) {
          const lbId = baseId + "_LB";
          requests.push(RequestFactory.createShape(slideId, lbId, "RECTANGLE", barLeft, barY, leftBarW, barHeight));
          requests.push(RequestFactory.updateShapeProperties(lbId, theme.colors.neutralGray, "TRANSPARENT"));
          const lbvId = baseId + "_LBV";
          requests.push(RequestFactory.createShape(slideId, lbvId, "TEXT_BOX", barLeft + leftBarW + layout.pxToPt(5), barY - layout.pxToPt(2), layout.pxToPt(60), barHeight + layout.pxToPt(5)));
          requests.push(...BatchTextStyleUtils.setText(slideId, lbvId, leftValue, { size: 10, color: theme.colors.neutralGray }, theme));
        }
        const rightBarW = rightNum / maxValue * maxBarWidth;
        if (rightBarW > 0) {
          const rbId = baseId + "_RB";
          requests.push(RequestFactory.createShape(slideId, rbId, "RECTANGLE", barLeft, barY + barHeight + barGap, rightBarW, barHeight));
          requests.push(RequestFactory.updateShapeProperties(rbId, settings.primaryColor, "TRANSPARENT"));
          const rbvId = baseId + "_RBV";
          requests.push(RequestFactory.createShape(slideId, rbvId, "TEXT_BOX", barLeft + rightBarW + layout.pxToPt(5), barY + barHeight + barGap - layout.pxToPt(2), layout.pxToPt(60), barHeight + layout.pxToPt(5)));
          requests.push(...BatchTextStyleUtils.setText(slideId, rbvId, rightValue, { size: 10, bold: true, color: settings.primaryColor }, theme));
        }
        if (trend) {
          const trendX = area.left + area.width - layout.pxToPt(40);
          const isUp = String(trend).toLowerCase() === "up";
          const trId = baseId + "_TR";
          requests.push(RequestFactory.createShape(slideId, trId, "TEXT_BOX", trendX, currentY, layout.pxToPt(40), rowHeight));
          const color = isUp ? "#2E7D32" : "#C62828";
          const sym = isUp ? "↑" : "↓";
          requests.push(...BatchTextStyleUtils.setText(slideId, trId, sym, { size: 20, bold: true, color, align: "CENTER" }, theme));
          requests.push(RequestFactory.updateShapeProperties(trId, null, null, null, "MIDDLE"));
        }
        if (index < stats.length - 1) {
          const lineY = currentY + rowHeight;
          const lineId = baseId + "_LN";
          requests.push(RequestFactory.createLine(slideId, lineId, area.left, lineY, area.left + area.width, lineY));
          requests.push({
            updateLineProperties: {
              objectId: lineId,
              lineProperties: {
                lineFill: { solidFill: { color: { rgbColor: RequestFactory.toRgbColor(theme.colors.ghostGray) || { red: 0.9, green: 0.9, blue: 0.9 } } } },
                weight: { magnitude: 0.5, unit: "PT" }
              },
              fields: "lineFill,weight"
            }
          });
        }
        currentY += rowHeight;
      });
      return requests;
    }
  }
  class StepUpDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const items = data.items || [];
      if (!items.length) return requests;
      const count = items.length;
      const validCount = Math.min(count, 5);
      const gap = 20;
      const barWidth = (area.width - gap * (validCount - 1)) / validCount;
      const maxBarH = area.height * 0.7;
      const minBarH = area.height * 0.3;
      items.slice(0, validCount).forEach((item, i) => {
        const ratio = i / (validCount - 1 || 1);
        const barH = minBarH + (maxBarH - minBarH) * ratio;
        const x = area.left + i * (barWidth + gap);
        const y = area.top + area.height - barH;
        const unique = `_STEP_${i}`;
        const barId = slideId + unique + "_BAR";
        const numId = slideId + unique + "_NUM";
        const titleId = slideId + unique + "_TITLE";
        const descId = slideId + unique + "_DESC";
        requests.push(RequestFactory.createShape(slideId, barId, "RECTANGLE", x, y, barWidth, barH));
        requests.push(RequestFactory.updateShapeProperties(barId, settings.primaryColor, "TRANSPARENT"));
        const numStr = String(i + 1).padStart(2, "0");
        requests.push(RequestFactory.createShape(slideId, numId, "TEXT_BOX", x, y + barH - 50, barWidth, 40));
        requests.push({ insertText: { objectId: numId, text: numStr } });
        requests.push(RequestFactory.updateTextStyle(numId, {
          fontSize: 32,
          bold: true,
          color: "#FFFFFF",
          // High contrast
          fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: numId, style: { alignment: "CENTER" }, fields: "alignment" } });
        const titleContent = typeof item === "string" ? item : item.title || item.label || "";
        requests.push(RequestFactory.createShape(slideId, titleId, "TEXT_BOX", x, y - 60, barWidth, 50));
        requests.push({ insertText: { objectId: titleId, text: titleContent } });
        requests.push(RequestFactory.updateTextStyle(titleId, {
          fontSize: 16,
          bold: true,
          color: theme.colors.textPrimary || "#333333",
          fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: "CENTER" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: "BOTTOM" }, fields: "contentAlignment" } });
        const desc = typeof item === "object" ? item.desc || item.description : null;
        if (desc) {
          const descBoxH = Math.min(barH - 60, 100);
          if (descBoxH > 30) {
            requests.push(RequestFactory.createShape(slideId, descId, "TEXT_BOX", x + 5, y + 10, barWidth - 10, descBoxH));
            requests.push({ insertText: { objectId: descId, text: desc } });
            requests.push(RequestFactory.updateTextStyle(descId, {
              fontSize: 12,
              color: "#FFFFFF",
              fontFamily: theme.fonts.family
            }));
            requests.push({ updateParagraphStyle: { objectId: descId, style: { alignment: "CENTER" }, fields: "alignment" } });
            requests.push({ updateShapeProperties: { objectId: descId, shapeProperties: { contentAlignment: "TOP" }, fields: "contentAlignment" } });
          }
        }
      });
      return requests;
    }
  }
  class AgendaDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const items = data.items || [];
      if (items.length === 0) return requests;
      const COLUMN_COUNT = 2;
      const ROW_COUNT = Math.ceil(items.length / COLUMN_COUNT);
      const GAP_X = 40;
      const GAP_Y = 25;
      const itemWidth = (area.width - GAP_X * (COLUMN_COUNT - 1)) / COLUMN_COUNT;
      const itemHeight = Math.min((area.height - GAP_Y * (ROW_COUNT - 1)) / ROW_COUNT, 80);
      items.forEach((itemText, index) => {
        const col = index % COLUMN_COUNT;
        const row = Math.floor(index / COLUMN_COUNT);
        const x = area.left + col * (itemWidth + GAP_X);
        const y = area.top + row * (itemHeight + GAP_Y);
        const uniqueId = `_AG_${index}`;
        const numShapeId = slideId + uniqueId + "_NUM";
        const cardShapeId = slideId + uniqueId + "_CARD";
        const numberSize = itemHeight;
        requests.push(RequestFactory.createShape(slideId, numShapeId, "RECTANGLE", x, y, numberSize, numberSize));
        requests.push(RequestFactory.updateShapeProperties(numShapeId, settings.primaryColor, "TRANSPARENT"));
        const numStr = (index + 1).toString().padStart(2, "0");
        requests.push({ insertText: { objectId: numShapeId, text: numStr } });
        requests.push(RequestFactory.updateTextStyle(numShapeId, {
          fontSize: 28,
          bold: true,
          color: "#FFFFFF",
          fontFamily: theme.fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: numShapeId, style: { alignment: "CENTER" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: numShapeId, shapeProperties: { contentAlignment: "MIDDLE" }, fields: "contentAlignment" } });
        const textWidth = itemWidth - numberSize;
        const textX = x + numberSize;
        requests.push(RequestFactory.createShape(slideId, cardShapeId, "RECTANGLE", textX, y, textWidth, itemHeight));
        requests.push(RequestFactory.updateShapeProperties(cardShapeId, settings.card_bg || "#F5F5F5", "TRANSPARENT"));
        requests.push({ insertText: { objectId: cardShapeId, text: itemText } });
        requests.push(RequestFactory.updateTextStyle(cardShapeId, {
          fontSize: 18,
          color: settings.text_primary || "#333333",
          fontFamily: theme.fonts.family
        }));
        requests.push({
          updateShapeProperties: {
            objectId: cardShapeId,
            shapeProperties: { contentAlignment: "MIDDLE" },
            fields: "contentAlignment"
          }
        });
      });
      return requests;
    }
  }
  class LanesDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const lanes = data.lanes || [];
      const n = Math.max(1, lanes.length);
      const { laneGapPx, lanePadPx, laneTitleHeightPx, cardGapPx, cardMinHeightPx, cardMaxHeightPx, arrowHeightPx, arrowGapPx } = theme.diagram;
      const px = (p) => layout.pxToPt(p);
      const laneW = (area.width - px(laneGapPx) * (n - 1)) / n;
      const cardBoxes = [];
      for (let j = 0; j < n; j++) {
        const lane = lanes[j] || { title: "", items: [] };
        const left = area.left + j * (laneW + px(laneGapPx));
        const laneIdPrefix = slideId + `_LANE_${j}`;
        const ltId = laneIdPrefix + "_HEAD";
        requests.push(RequestFactory.createShape(slideId, ltId, "RECTANGLE", left, area.top, laneW, px(laneTitleHeightPx)));
        requests.push(RequestFactory.updateShapeProperties(ltId, settings.primaryColor, settings.primaryColor, 1));
        requests.push(...BatchTextStyleUtils.setText(slideId, ltId, lane.title || "", {
          size: theme.fonts.sizes.laneTitle,
          bold: true,
          color: theme.colors.backgroundGray,
          align: "CENTER"
        }, theme));
        requests.push(RequestFactory.updateShapeProperties(ltId, null, null, null, "MIDDLE"));
        const laneBodyTop = area.top + px(laneTitleHeightPx);
        const laneBodyHeight = area.height - px(laneTitleHeightPx);
        const lbId = laneIdPrefix + "_BODY";
        requests.push(RequestFactory.createShape(slideId, lbId, "RECTANGLE", left, laneBodyTop, laneW, laneBodyHeight));
        requests.push(RequestFactory.updateShapeProperties(lbId, theme.colors.backgroundGray, theme.colors.laneBorder));
        const items = Array.isArray(lane.items) ? lane.items : [];
        const rows = Math.max(1, items.length);
        const availH = laneBodyHeight - px(lanePadPx) * 2;
        const idealH = (availH - px(cardGapPx) * (rows - 1)) / rows;
        const cardH = Math.max(px(cardMinHeightPx), Math.min(px(cardMaxHeightPx), idealH));
        const firstTop = laneBodyTop + px(lanePadPx) + Math.max(0, (availH - (cardH * rows + px(cardGapPx) * (rows - 1))) / 2);
        cardBoxes[j] = [];
        for (let i = 0; i < rows; i++) {
          const cardTop = firstTop + i * (cardH + px(cardGapPx));
          const cardId = laneIdPrefix + `_CARD_${i}`;
          requests.push(RequestFactory.createShape(slideId, cardId, "ROUND_RECTANGLE", left + px(lanePadPx), cardTop, laneW - px(lanePadPx) * 2, cardH));
          requests.push(RequestFactory.updateShapeProperties(cardId, theme.colors.backgroundGray, theme.colors.cardBorder));
          requests.push(...BatchTextStyleUtils.setText(slideId, cardId, items[i] || "", {
            size: theme.fonts.sizes.body
          }, theme));
          requests.push(RequestFactory.updateShapeProperties(cardId, null, null, null, "MIDDLE"));
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
            const boxA = cardBoxes[j][i];
            const boxB = cardBoxes[j + 1][i];
            const startX = boxA.left + boxA.width;
            const startY = boxA.top + boxA.height / 2;
            const endX = boxB.left;
            const endY = boxB.top + boxB.height / 2;
            const arrowId = slideId + `_ARR_${j}_${i}`;
            requests.push(RequestFactory.createLine(slideId, arrowId, startX, startY, endX, endY));
            const arrowColor = RequestFactory.toRgbColor(settings.primaryColor) || { red: 0, green: 0, blue: 0 };
            requests.push({
              updateLineProperties: {
                objectId: arrowId,
                lineProperties: {
                  lineFill: { solidFill: { color: { rgbColor: arrowColor } } },
                  weight: { magnitude: 2, unit: "PT" },
                  endArrow: "FILL_ARROW"
                },
                fields: "lineFill,weight,endArrow"
              }
            });
          }
        }
      }
      return requests;
    }
  }
  class FlowChartDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const steps = data.steps || data.items || [];
      if (!steps.length) return requests;
      const count = steps.length;
      const gap = 30;
      const boxWidth = (area.width - gap * (count - 1)) / count;
      const boxHeight = 80;
      const y = area.top + (area.height - boxHeight) / 2;
      for (let i = 0; i < count; i++) {
        const step = steps[i];
        const x = area.left + i * (boxWidth + gap);
        const baseId = slideId + `_FLOW_${i}`;
        const boxId = baseId + "_BOX";
        const arrowId = baseId + "_ARR";
        requests.push(RequestFactory.createShape(slideId, boxId, "ROUND_RECTANGLE", x, y, boxWidth, boxHeight));
        requests.push(RequestFactory.updateShapeProperties(boxId, theme.colors.backgroundGray, settings.primaryColor, 2, "MIDDLE"));
        const label = typeof step === "string" ? step : step.label || "";
        requests.push(...BatchTextStyleUtils.setText(slideId, boxId, label, {
          size: theme.fonts.sizes.body,
          align: "CENTER",
          color: theme.colors.textPrimary
        }, theme));
        if (i < count - 1) {
          const ax = x + boxWidth;
          const ay = y + boxHeight / 2;
          const bx = x + boxWidth + gap;
          const by = ay;
          requests.push(RequestFactory.createLine(slideId, arrowId, ax, ay, bx, by));
          requests.push({
            updateLineProperties: {
              objectId: arrowId,
              lineProperties: {
                lineFill: { solidFill: { color: { rgbColor: { red: 0.6, green: 0.6, blue: 0.6 } } } },
                // Neutral Gray (approx)
                weight: { magnitude: 1, unit: "PT" },
                endArrow: "ARROW"
              },
              fields: "lineFill,weight,endArrow"
            }
          });
        }
      }
      return requests;
    }
  }
  class KPIDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const items = data.items || [];
      if (!items.length) return requests;
      const cols = items.length > 4 ? 4 : items.length || 1;
      const gap = layout.pxToPt(40);
      const cardW = (area.width - gap * (cols - 1)) / cols;
      const cardH = layout.pxToPt(180);
      const y = area.top + (area.height - cardH) / 2 + layout.pxToPt(20);
      items.forEach((item, i) => {
        const x = area.left + i * (cardW + gap);
        const baseId = slideId + `_KPI_${i}`;
        if (i > 0) {
          const lineH = layout.pxToPt(100);
          const lineY = y + (cardH - lineH) / 2;
          const lineX = x - gap / 2;
          const lineId = baseId + "_LN";
          requests.push(RequestFactory.createLine(slideId, lineId, lineX, lineY, lineX, lineY + lineH));
          requests.push({
            updateLineProperties: {
              objectId: lineId,
              lineProperties: {
                lineFill: { solidFill: { color: { rgbColor: { red: 0.9, green: 0.9, blue: 0.95 } } } },
                // Ghost Gray approx
                weight: { magnitude: 1, unit: "PT" }
              },
              fields: "lineFill,weight"
            }
          });
        }
        const labelH = layout.pxToPt(20);
        const labelId = baseId + "_LBL";
        requests.push(RequestFactory.createShape(slideId, labelId, "TEXT_BOX", x, y + layout.pxToPt(5), cardW, labelH));
        requests.push(RequestFactory.updateShapeProperties(labelId, null, null, null, "BOTTOM"));
        requests.push(...BatchTextStyleUtils.setText(slideId, labelId, (item.label || "METRIC").toUpperCase(), {
          size: 11,
          color: theme.colors.neutralGray,
          align: "CENTER",
          bold: true
        }, theme));
        const valueH = layout.pxToPt(90);
        const valueId = baseId + "_VAL";
        requests.push(RequestFactory.createShape(slideId, valueId, "TEXT_BOX", x, y + labelH, cardW, valueH));
        requests.push(RequestFactory.updateShapeProperties(valueId, null, null, null, "TOP"));
        const valStr = String(item.value || "0");
        let fontSize = 72;
        if (valStr.length > 4) fontSize = 60;
        if (valStr.length > 6) fontSize = 48;
        if (valStr.length > 10) fontSize = 36;
        requests.push(...BatchTextStyleUtils.setText(slideId, valueId, valStr, {
          size: fontSize,
          bold: true,
          color: settings.primaryColor,
          align: "CENTER",
          fontType: "lato"
        }, theme));
        if (item.change || item.status) {
          const statusH = layout.pxToPt(30);
          const statusId = baseId + "_STS";
          requests.push(RequestFactory.createShape(slideId, statusId, "TEXT_BOX", x, y + labelH + valueH, cardW, statusH));
          let color = theme.colors.neutralGray;
          let prefix = "";
          if (item.status === "good") {
            color = "#2E7D32";
            prefix = "↑ ";
          }
          if (item.status === "bad") {
            color = "#C62828";
            prefix = "↓ ";
          }
          requests.push(...BatchTextStyleUtils.setText(slideId, statusId, prefix + (item.change || ""), {
            size: 14,
            bold: true,
            color,
            align: "CENTER"
          }, theme));
        }
      });
      return requests;
    }
  }
  class TableDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const headers = data.headers || [];
      const rows = data.rows || [];
      const theme = layout.getTheme();
      const numCols = Math.max(headers.length, rows.length > 0 && Array.isArray(rows[0]) ? rows[0].length : 0);
      const numRows = rows.length + (headers.length ? 1 : 0);
      if (numRows === 0 || numCols === 0) return requests;
      const tableId = slideId + "_TBL";
      requests.push({
        createTable: {
          objectId: tableId,
          elementProperties: {
            pageObjectId: slideId,
            size: { width: { magnitude: area.width, unit: "PT" }, height: { magnitude: area.height, unit: "PT" } },
            transform: { scaleX: 1, scaleY: 1, translateX: area.left, translateY: area.top, unit: "PT" }
          },
          rows: numRows,
          columns: numCols
        }
      });
      requests.push({
        updateTableBorderProperties: {
          objectId: tableId,
          tableRange: { location: { rowIndex: 0, columnIndex: 0 }, rowSpan: numRows, columnSpan: numCols },
          borderPosition: "ALL",
          tableBorderProperties: {
            tableBorderFill: { solidFill: { color: { rgbColor: { red: 0, green: 0, blue: 0 } } } },
            // Should accept transparent?
            // Actually API handles transparency via propertyState='NOT_RENDERED' or fill alpha?
            // Let's set fill to NOT_RENDERED doesn't exist for TableBorderFill? 
            // Actually it does: "solidFill: { alpha: 0 }" or similar.
            // Or set weight to 0.
            weight: { magnitude: 0, unit: "PT" }
          },
          fields: "weight"
        }
      });
      let rowIndex = 0;
      if (headers.length) {
        headers.forEach((header, c) => {
          const text = header || "";
          if (text) {
            requests.push({ insertText: { objectId: tableId, cellLocation: { rowIndex: 0, columnIndex: c }, text } });
            requests.push(RequestFactory.updateTextStyle(tableId, {
              bold: true,
              fontSize: 14,
              color: "#FFFFFF",
              fontFamily: theme.fonts.family
            }));
            requests.push({
              updateTextStyle: {
                objectId: tableId,
                cellLocation: { rowIndex: 0, columnIndex: c },
                style: {
                  bold: true,
                  fontSize: { magnitude: 14, unit: "PT" },
                  foregroundColor: { opaqueColor: { themeColor: "BACKGROUND1" } },
                  fontFamily: theme.fonts.family
                },
                fields: "bold,fontSize,foregroundColor,fontFamily"
              }
            });
          }
          requests.push({
            updateTableCellProperties: {
              objectId: tableId,
              tableRange: { location: { rowIndex: 0, columnIndex: c }, rowSpan: 1, columnSpan: 1 },
              tableCellProperties: {
                tableCellBackgroundFill: { solidFill: { color: { themeColor: "ACCENT1" } } },
                contentAlignment: "MIDDLE"
              },
              fields: "tableCellBackgroundFill,contentAlignment"
            }
          });
        });
        rowIndex++;
      }
      rows.forEach((row, rIdx) => {
        var _a;
        const actualRowIndex = rowIndex + rIdx;
        const isEven = actualRowIndex % 2 === 0;
        const bgFill = isEven ? { solidFill: { color: { rgbColor: { red: 0.97, green: 0.97, blue: 0.98 } } } } : { solidFill: { color: { themeColor: "BACKGROUND1" } } };
        for (let c = 0; c < numCols; c++) {
          const cellText = String((_a = row[c]) != null ? _a : "");
          if (cellText) {
            requests.push({ insertText: { objectId: tableId, cellLocation: { rowIndex: actualRowIndex, columnIndex: c }, text: cellText } });
            requests.push({
              updateTextStyle: {
                objectId: tableId,
                cellLocation: { rowIndex: actualRowIndex, columnIndex: c },
                style: {
                  fontSize: { magnitude: 12, unit: "PT" },
                  foregroundColor: { opaqueColor: { themeColor: "TEXT1" } },
                  fontFamily: theme.fonts.family
                },
                fields: "fontSize,foregroundColor,fontFamily"
              }
            });
          }
          requests.push({
            updateTableCellProperties: {
              objectId: tableId,
              tableRange: { location: { rowIndex: actualRowIndex, columnIndex: c }, rowSpan: 1, columnSpan: 1 },
              tableCellProperties: {
                tableCellBackgroundFill: bgFill,
                contentAlignment: "MIDDLE"
              },
              fields: "tableCellBackgroundFill,contentAlignment"
            }
          });
          requests.push({
            updateTableBorderProperties: {
              objectId: tableId,
              tableRange: { location: { rowIndex: actualRowIndex, columnIndex: c }, rowSpan: 1, columnSpan: 1 },
              borderPosition: "BOTTOM",
              tableBorderProperties: {
                tableBorderFill: { solidFill: { color: { rgbColor: { red: 0.9, green: 0.9, blue: 0.9 } } } },
                weight: { magnitude: 1, unit: "PT" },
                dashStyle: "SOLID"
              },
              fields: "tableBorderFill,weight,dashStyle"
            }
          });
        }
      });
      return requests;
    }
  }
  class FAQDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
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
      if (!parsedItems.length) return requests;
      const gap = layout.pxToPt(30);
      const itemH = (area.height - gap * (parsedItems.length - 1)) / parsedItems.length;
      parsedItems.forEach((item, i) => {
        const y = area.top + i * (itemH + gap);
        const qStr = (item.q || "").replace(/^[QA][:. ]+/, "");
        const aStr = (item.a || "").replace(/^[QA][:. ]+/, "");
        const baseId = slideId + `_FAQ_${i}`;
        const qIndId = baseId + "_QI";
        requests.push(RequestFactory.createShape(slideId, qIndId, "TEXT_BOX", area.left, y, layout.pxToPt(30), layout.pxToPt(30)));
        requests.push(...BatchTextStyleUtils.setText(slideId, qIndId, "Q.", { size: 16, bold: true, color: settings.primaryColor }, theme));
        const qBoxId = baseId + "_QC";
        requests.push(RequestFactory.createShape(slideId, qBoxId, "TEXT_BOX", area.left + layout.pxToPt(30), y, area.width - layout.pxToPt(30), layout.pxToPt(40)));
        requests.push(...BatchTextStyleUtils.setText(slideId, qBoxId, qStr, { size: 14, bold: true, color: theme.colors.textPrimary }, theme));
        const aY = y + layout.pxToPt(30);
        const aIndId = baseId + "_AI";
        requests.push(RequestFactory.createShape(slideId, aIndId, "TEXT_BOX", area.left, aY, layout.pxToPt(30), layout.pxToPt(30)));
        requests.push(...BatchTextStyleUtils.setText(slideId, aIndId, "A.", { size: 16, bold: true, color: theme.colors.neutralGray }, theme));
        const aBoxHasHeight = itemH - layout.pxToPt(40);
        if (aBoxHasHeight > 10) {
          const aBoxId = baseId + "_AC";
          requests.push(RequestFactory.createShape(slideId, aBoxId, "TEXT_BOX", area.left + layout.pxToPt(30), aY, area.width - layout.pxToPt(30), aBoxHasHeight));
          requests.push(...BatchTextStyleUtils.setText(slideId, aBoxId, aStr, { size: 12, color: theme.colors.textPrimary }, theme));
          requests.push(RequestFactory.updateShapeProperties(aBoxId, null, null, null, "TOP"));
        }
        if (i < parsedItems.length - 1) {
          const lineY = y + itemH + gap / 2;
          const lineId = baseId + "_LN";
          requests.push(RequestFactory.createLine(slideId, lineId, area.left, lineY, area.left + area.width, lineY));
          requests.push({
            updateLineProperties: {
              objectId: lineId,
              lineProperties: {
                lineFill: { solidFill: { color: { rgbColor: { red: 0.9, green: 0.9, blue: 0.9 } } } },
                // Approx Ghost Gray
                weight: { magnitude: 0.5, unit: "PT" }
              },
              fields: "lineFill,weight"
            }
          });
        }
      });
      return requests;
    }
  }
  class QuoteDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const text = data.text || data.points && data.points[0] || "";
      const author = data.author || data.points && data.points[1] || "";
      const quoteSize = layout.pxToPt(200);
      const quoteId = slideId + "_QUOTE_MARK";
      requests.push(RequestFactory.createShape(slideId, quoteId, "TEXT_BOX", area.left + layout.pxToPt(20), area.top - layout.pxToPt(40), quoteSize, quoteSize));
      requests.push(...BatchTextStyleUtils.setText(slideId, quoteId, "“", {
        size: 200,
        color: "#F0F0F0",
        fontFamily: "Georgia",
        bold: true
      }, theme));
      const contentW = area.width * 0.9;
      const contentX = area.left + (area.width - contentW) / 2;
      const textTop = area.top + layout.pxToPt(60);
      const textId = slideId + "_QUOTE_TEXT";
      requests.push(RequestFactory.createShape(slideId, textId, "TEXT_BOX", contentX, textTop, contentW, layout.pxToPt(160)));
      requests.push(...BatchTextStyleUtils.setText(slideId, textId, text, {
        size: 32,
        bold: false,
        // Explicit false
        color: theme.colors.textPrimary,
        align: "CENTER",
        fontFamily: "Georgia"
      }, theme));
      requests.push(RequestFactory.updateShapeProperties(textId, null, null, null, "BOTTOM"));
      const lineW = layout.pxToPt(40);
      const lineX = area.left + (area.width - lineW) / 2;
      const lineY = textTop + layout.pxToPt(165);
      const lineId = slideId + "_QUOTE_LINE";
      requests.push(RequestFactory.createLine(slideId, lineId, lineX, lineY, lineX + lineW, lineY));
      requests.push({
        updateLineProperties: {
          objectId: lineId,
          lineProperties: {
            lineFill: { solidFill: { color: { rgbColor: RequestFactory.toRgbColor(settings.primaryColor) || { red: 0, green: 0, blue: 0 } } } },
            weight: { magnitude: 2, unit: "PT" }
          },
          fields: "lineFill,weight"
        }
      });
      if (author) {
        const authId = slideId + "_QUOTE_AUTH";
        requests.push(RequestFactory.createShape(slideId, authId, "TEXT_BOX", contentX, lineY + layout.pxToPt(5), contentW, layout.pxToPt(30)));
        requests.push(...BatchTextStyleUtils.setText(slideId, authId, author, {
          size: 14,
          align: "CENTER",
          color: theme.colors.neutralGray,
          bold: true
        }, theme));
        requests.push(RequestFactory.updateShapeProperties(authId, null, null, null, "TOP"));
      }
      return requests;
    }
  }
  class ProgressDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const items = data.items || [];
      if (!items.length) return requests;
      const rowH = layout.pxToPt(50);
      const gap = layout.pxToPt(15);
      const startY = area.top + (area.height - items.length * (rowH + gap)) / 2;
      items.forEach((item, i) => {
        const y = startY + i * (rowH + gap);
        const labelW = layout.pxToPt(150);
        const barAreaW = area.width - labelW - layout.pxToPt(60);
        const baseId = slideId + `_PROG_${i}`;
        const labelId = baseId + "_LBL";
        requests.push(RequestFactory.createShape(slideId, labelId, "TEXT_BOX", area.left, y, labelW, rowH));
        requests.push(...BatchTextStyleUtils.setText(slideId, labelId, item.label || "", {
          size: 14,
          bold: true,
          align: "END"
        }, theme));
        requests.push(RequestFactory.updateShapeProperties(labelId, null, null, null, "MIDDLE"));
        const barBgId = baseId + "_BG";
        requests.push(RequestFactory.createShape(slideId, barBgId, "ROUND_RECTANGLE", area.left + labelW + layout.pxToPt(20), y + rowH / 3, barAreaW, rowH / 3));
        requests.push(RequestFactory.updateShapeProperties(barBgId, theme.colors.ghostGray, "TRANSPARENT"));
        const percent = Math.min(100, Math.max(0, parseInt(item.percent || 0)));
        if (percent > 0) {
          const barFgId = baseId + "_FG";
          const fgW = barAreaW * (percent / 100);
          requests.push(RequestFactory.createShape(slideId, barFgId, "ROUND_RECTANGLE", area.left + labelW + layout.pxToPt(20), y + rowH / 3, fgW, rowH / 3));
          requests.push(RequestFactory.updateShapeProperties(barFgId, settings.primaryColor, "TRANSPARENT"));
        }
        const valId = baseId + "_VAL";
        requests.push(RequestFactory.createShape(slideId, valId, "TEXT_BOX", area.left + labelW + barAreaW + layout.pxToPt(30), y, layout.pxToPt(50), rowH));
        requests.push(...BatchTextStyleUtils.setText(slideId, valId, `${percent}%`, {
          size: 14,
          color: theme.colors.neutralGray
        }, theme));
        requests.push(RequestFactory.updateShapeProperties(valId, null, null, null, "MIDDLE"));
      });
      return requests;
    }
  }
  class ImageTextDiagramRenderer {
    render(slideId, data, area, settings, layout) {
      const requests = [];
      const theme = layout.getTheme();
      const imageUrl = data.image;
      const points = data.points || [];
      const gap = layout.pxToPt(20);
      const halfW = (area.width - gap) / 2;
      const isImageLeft = data.imagePosition !== "right";
      const imgX = isImageLeft ? area.left : area.left + halfW + gap;
      const txtX = isImageLeft ? area.left + halfW + gap : area.left;
      const imgId = slideId + "_IMG";
      if (imageUrl && typeof imageUrl === "string" && imageUrl.startsWith("http")) {
        requests.push(RequestFactory.createImage(slideId, imgId, imageUrl, imgX, area.top, halfW, area.height));
      } else {
        const phId = slideId + "_IMG_PH";
        requests.push(RequestFactory.createShape(slideId, phId, "RECTANGLE", imgX, area.top, halfW, area.height));
        requests.push(RequestFactory.updateShapeProperties(phId, theme.colors.ghostGray, "TRANSPARENT"));
        requests.push(...BatchTextStyleUtils.setText(slideId, phId, "Image Placeholder (URL required)", {
          size: 14,
          align: "CENTER",
          color: theme.colors.neutralGray
        }, theme));
        requests.push(RequestFactory.updateShapeProperties(phId, null, null, null, "MIDDLE"));
      }
      const txtId = slideId + "_TXT";
      const textContent = Array.isArray(points) ? points.join("\n") : String(points);
      requests.push(RequestFactory.createShape(slideId, txtId, "TEXT_BOX", txtX, area.top, halfW, area.height));
      requests.push(...BatchTextStyleUtils.setText(slideId, txtId, textContent, {
        size: theme.fonts.sizes.body,
        color: theme.colors.textPrimary
      }, theme));
      return requests;
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
      if (normalizedType.includes("agenda")) {
        Logger.log("[Factory] Matched AgendaDiagramRenderer");
        return new AgendaDiagramRenderer();
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
    generate(slideId, data, layout, pageNum, settings) {
      const requests = [];
      if (data.title) {
        const titleId = slideId + "_TITLE";
        const titleRect = layout.getRect("contentSlide.title") || { left: 30, top: 20, width: 900, height: 60 };
        requests.push(RequestFactory.createShape(slideId, titleId, "TEXT_BOX", titleRect.left, titleRect.top, titleRect.width, titleRect.height));
        requests.push({ insertText: { objectId: titleId, text: data.title } });
        requests.push(RequestFactory.updateTextStyle(titleId, {
          fontSize: 32,
          bold: true,
          color: settings.primaryColor,
          fontFamily: layout.getTheme().fonts.family
        }));
        requests.push({ updateParagraphStyle: { objectId: titleId, style: { alignment: "START" }, fields: "alignment" } });
        requests.push({ updateShapeProperties: { objectId: titleId, shapeProperties: { contentAlignment: "BOTTOM" }, fields: "contentAlignment" } });
      }
      const type = (data.type || data.layout || "").toLowerCase();
      const renderer = DiagramRendererFactory.getRenderer(type);
      if (renderer) {
        const typeKey = `${type}Slide`;
        const areaRect = layout.getRect(`${typeKey}.area`) || layout.getRect("contentSlide.body");
        const workArea = {
          left: areaRect.left,
          top: areaRect.top,
          width: areaRect.width,
          height: areaRect.height
        };
        const diagramReqs = renderer.render(slideId, data, workArea, settings, layout);
        requests.push(...diagramReqs);
      } else {
        Logger.log("Diagram logic not implemented for type: " + type);
      }
      return requests;
    }
  }
  class GasSlideRepository {
    // Note: All generators (Title, Section, Content, Diagram) are refactored to use Batch API.
    createPresentation(presentation, templateId, destinationId, settingsOverride) {
      const slidesApp = SlidesApp;
      const driveApp = DriveApp;
      let pres;
      let presId;
      if (destinationId) {
        pres = slidesApp.openById(destinationId);
      } else if (templateId) {
        const templateFile = driveApp.getFileById(templateId);
        const newFile = templateFile.makeCopy(presentation.title);
        pres = slidesApp.openById(newFile.getId());
      } else {
        pres = slidesApp.create(presentation.title);
      }
      presId = pres ? pres.getId() : "mock-presentation-id";
      const pageWidth = pres.getPageWidth();
      const pageHeight = pres.getPageHeight();
      const themeName = settingsOverride && settingsOverride.theme ? settingsOverride.theme : "Green";
      const selectedTheme = AVAILABLE_THEMES[themeName] || DEFAULT_THEME;
      const layoutManager = new LayoutManager(pageWidth, pageHeight, selectedTheme);
      const settings = { primaryColor: selectedTheme.colors.primary, ...settingsOverride };
      const layoutIdMap = /* @__PURE__ */ new Map();
      try {
        const presData = Slides.Presentations.get(presId);
        if (presData.layouts) {
          presData.layouts.forEach((layout) => {
            const props = layout.layoutProperties;
            if (props) {
              if (props.displayName) {
                layoutIdMap.set(props.displayName.toUpperCase(), layout.objectId);
              }
              if (props.name) {
                layoutIdMap.set(props.name.toUpperCase(), layout.objectId);
              }
            }
          });
        }
      } catch (e) {
        Logger.log("Warning: Failed to fetch presentation details via Advanced API. Layout mapping may be limited. " + e);
      }
      Logger.log(`Mapped Layouts (Display Name -> ID): ${Array.from(layoutIdMap.keys()).join(", ")}`);
      const titleGenerator = new GasTitleSlideGenerator(null);
      const sectionGenerator = new GasSectionSlideGenerator(null);
      const contentGenerator = new GasContentSlideGenerator(null);
      const diagramGenerator = new GasDiagramSlideGenerator(null);
      let allRequests = [];
      let cleanupSlideId = null;
      if (!templateId && !destinationId) {
        const slides = pres.getSlides();
        if (slides.length > 0) {
          cleanupSlideId = slides[0].getObjectId();
        }
      }
      presentation.slides.forEach((slideModel, index) => {
        var _a;
        const layoutType = (slideModel.layout || "content").toUpperCase();
        const rawType = (((_a = slideModel.rawData) == null ? void 0 : _a.type) || "").toLowerCase();
        Logger.log(`[Slide ${index + 1}] Processing - Requested Layout: '${slideModel.layout}', Normalized: '${layoutType}', RawType: '${rawType}'`);
        let slideLayout;
        const layouts = pres.getLayouts();
        const targetLayoutId = layoutIdMap.get(layoutType);
        if (targetLayoutId) {
          slideLayout = layouts.find((l) => l.getObjectId() === targetLayoutId);
        }
        if (!slideLayout) {
          slideLayout = layouts.find((l) => l.getLayoutName().toUpperCase() === layoutType);
        }
        if (!slideLayout && layoutType !== "TITLE" && layoutType !== "SECTION") {
          const fallbackId = layoutIdMap.get("CONTENT");
          if (fallbackId) {
            slideLayout = layouts.find((l) => l.getObjectId() === fallbackId);
            if (slideLayout) {
              Logger.log(`[Slide ${index + 1}] Layout '${layoutType}' not found. Falling back to 'CONTENT'.`);
            }
          }
          if (!slideLayout) {
            slideLayout = layouts.find((l) => l.getLayoutName().toUpperCase() === "CONTENT" || l.getLayoutName().toUpperCase() === "TITLE_AND_BODY");
          }
        }
        if (!slideLayout) {
          slideLayout = layouts.find((l) => l.getLayoutName().toUpperCase() === "BLANK") || layouts[layouts.length - 1];
          Logger.log(`[Slide ${index + 1}] Layout '${layoutType}' not found. Falling back to '${slideLayout.getLayoutName()}'`);
        } else {
          Logger.log(`[Slide ${index + 1}] Found Layout: '${slideLayout.getLayoutName()}' (ID: ${slideLayout.getObjectId()})`);
        }
        const slideId = `gen_slide_${index}`;
        allRequests.push(RequestFactory.createSlide(slideId, slideLayout.getObjectId(), false));
        const commonData = {
          title: slideModel.title.value,
          subtitle: slideModel.subtitle,
          date: (/* @__PURE__ */ new Date()).toLocaleDateString(),
          points: slideModel.content.items,
          content: slideModel.content.items,
          ...slideModel.rawData
        };
        let reqs = [];
        if (layoutType === "TITLE") {
          reqs = titleGenerator.generate(slideId, commonData, layoutManager, index + 1, settings);
        } else if (layoutType === "SECTION" || layoutType === "SECTION_HEADER") {
          reqs = sectionGenerator.generate(slideId, commonData, layoutManager, index + 1, settings);
        } else if (rawType.includes("timeline") || rawType.includes("process") || // ... other diagrams
        rawType.includes("diagram")) {
          reqs = diagramGenerator.generate(slideId, commonData, layoutManager, index + 1, settings);
        } else {
          reqs = contentGenerator.generate(slideId, commonData, layoutManager, index + 1, settings);
        }
        Logger.log(`[Slide ${index + 1}] Generator Used for layoutType '${layoutType}': ${reqs.length} requests generated.`);
        allRequests.push(...reqs);
      });
      const presentationUrl = pres.getUrl();
      if (allRequests.length > 0) {
        if (cleanupSlideId) {
          allRequests.push({ deleteObject: { objectId: cleanupSlideId } });
        }
        try {
          Slides.Presentations.batchUpdate({ requests: allRequests }, presId);
          Logger.log(`Executed ${allRequests.length} requests successfully.`);
        } catch (e) {
          Logger.log(`Batch Update Failed: ${e}`);
          throw e;
        }
      }
      return presentationUrl;
    }
  }
  function generateSlides(data) {
    try {
      Logger.log("Library Call: generateSlides with data: " + JSON.stringify(data));
      let slides = data.slides;
      let theme = data.theme;
      if (Array.isArray(data)) {
        slides = data;
      }
      if (!Array.isArray(slides)) {
        throw new Error("Invalid input: 'slides' must be an array.");
      }
      const request = {
        title: data.title || "無題のプレゼンテーション",
        templateId: data.templateId,
        // Optional ID for template
        destinationId: data.destinationId,
        // Optional ID for existing destination
        settings: {
          ...data.settings || {},
          theme: theme || data.settings && data.settings.theme
        },
        slides: slides.map((s) => ({
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
  class GasSlidesApiTester {
    /**
     * Executes a prototype batch update to demonstrate API speed and atomicity.
     * @param presentationId The ID of the target presentation
     */
    runPrototype(presentationId) {
      const pageId = "TEST_SLIDE_" + Math.floor(Math.random() * 1e5);
      const titleId = pageId + "_TITLE";
      const bodyId = pageId + "_BODY";
      const shapeId = pageId + "_SHAPE";
      const requests = [
        // 1. Create a generic Slide
        {
          createSlide: {
            objectId: pageId,
            slideLayoutReference: {
              predefinedLayout: "TITLE_AND_BODY"
            },
            placeholderIdMappings: [
              {
                layoutPlaceholder: { type: "TITLE", index: 0 },
                objectId: titleId
              },
              {
                layoutPlaceholder: { type: "BODY", index: 0 },
                objectId: bodyId
              }
            ]
          }
        },
        // 2. Populate Title
        {
          insertText: {
            objectId: titleId,
            text: "🚀 High-Speed API Generation"
          }
        },
        // 3. Populate Body with multiple bullets
        {
          insertText: {
            objectId: bodyId,
            text: "This entire slide was created in ONE round-trip request.\n\nBenefits:\n- No timeout risks for large decks\n- Atomic (all or nothing)\n- Precise ID control"
          }
        },
        // 4. Create a custom shape (Star)
        {
          createShape: {
            objectId: shapeId,
            shapeType: "STAR_5_POINT",
            elementProperties: {
              pageObjectId: pageId,
              size: {
                width: { magnitude: 120, unit: "PT" },
                height: { magnitude: 120, unit: "PT" }
              },
              transform: {
                scaleX: 1,
                scaleY: 1,
                translateX: 500,
                translateY: 100,
                unit: "PT"
              }
            }
          }
        },
        // 5. Add text to shape and style it
        {
          insertText: {
            objectId: shapeId,
            text: "FAST"
          }
        },
        {
          updateTextStyle: {
            objectId: shapeId,
            style: {
              bold: true,
              fontSize: { magnitude: 18, unit: "PT" },
              foregroundColor: {
                opaqueColor: { themeColor: "ACCENT1" }
              }
            },
            fields: "bold,fontSize,foregroundColor"
          }
        }
      ];
      Logger.log("Starting Batch Update...");
      const startTime = (/* @__PURE__ */ new Date()).getTime();
      try {
        const resource = { requests };
        Slides.Presentations.batchUpdate(resource, presentationId);
        const duration = (/* @__PURE__ */ new Date()).getTime() - startTime;
        Logger.log(`✅ Batch Update Complete in ${duration}ms!`);
        return `Generated 1 slide with ${requests.length} operations in ${duration}ms.`;
      } catch (e) {
        Logger.log("❌ Error in Batch Update: " + e.toString());
        throw new Error("API Prototype Failed: " + e.message);
      }
    }
  }
  global.generateSlides = generateSlides;
  global.doPost = doPost;
  global.doGet = doGet;
  global.testSlidesApi = () => {
    const pres = SlidesApp.create("API Prototype Test " + (/* @__PURE__ */ new Date()).toLocaleString());
    const id = pres.getId();
    Logger.log("Created Temp Presentation: " + pres.getUrl());
    const tester = new GasSlidesApiTester();
    const result = tester.runPrototype(id);
    Logger.log(result);
  };
})();
