var global = this;
(() => {
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __esm = (fn, res) => function __init() {
    return fn && (res = (0, fn[__getOwnPropNames(fn)[0]])(fn = 0)), res;
  };
  var __commonJS = (cb, mod) => function __require() {
    return mod || (0, cb[__getOwnPropNames(cb)[0]])((mod = { exports: {} }).exports, mod), mod.exports;
  };

  // src/domain/model/Presentation.ts
  var Presentation;
  var init_Presentation = __esm({
    "src/domain/model/Presentation.ts"() {
      Presentation = class {
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
      };
    }
  });

  // src/domain/model/Slide.ts
  var Slide;
  var init_Slide = __esm({
    "src/domain/model/Slide.ts"() {
      Slide = class {
        constructor(title, content, layout = "CONTENT", subtitle, notes) {
          this.title = title;
          this.content = content;
          this.layout = layout;
          this.subtitle = subtitle;
          this.notes = notes;
        }
      };
    }
  });

  // src/domain/model/SlideElement.ts
  var SlideTitle, SlideContent;
  var init_SlideElement = __esm({
    "src/domain/model/SlideElement.ts"() {
      SlideTitle = class {
        constructor(value) {
          this.value = value;
          if (!value) {
          }
        }
      };
      SlideContent = class {
        constructor(items) {
          this.items = items;
        }
      };
    }
  });

  // src/application/PresentationApplicationService.ts
  var PresentationApplicationService;
  var init_PresentationApplicationService = __esm({
    "src/application/PresentationApplicationService.ts"() {
      init_Presentation();
      init_Slide();
      init_SlideElement();
      PresentationApplicationService = class {
        constructor(slideRepository) {
          this.slideRepository = slideRepository;
        }
        createPresentation(request) {
          const presentation = new Presentation(request.title);
          for (const slideData of request.slides) {
            const title = new SlideTitle(slideData.title);
            const content = new SlideContent(slideData.content);
            const layout = slideData.layout || "CONTENT";
            const slide = new Slide(title, content, layout, slideData.subtitle, slideData.notes);
            presentation.addSlide(slide);
          }
          return this.slideRepository.createPresentation(presentation, request.templateId);
        }
      };
    }
  });

  // src/common/config/SlideConfig.ts
  var CONFIG;
  var init_SlideConfig = __esm({
    "src/common/config/SlideConfig.ts"() {
      CONFIG = {
        BASE_PX: {
          W: 960,
          H: 540
        },
        BACKGROUND_IMAGES: {
          title: "",
          closing: "",
          section: "",
          main: ""
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
          family: "Noto Sans JP",
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
          primary_color: "#4285F4",
          text_primary: "#333333",
          text_small_font: "#1F2937",
          background_white: "#FFFFFF",
          card_bg: "#f6e9f0",
          background_gray: "",
          faint_gray: "",
          ghost_gray: "",
          table_header_bg: "",
          lane_border: "",
          card_border: "",
          neutral_gray: "",
          process_arrow: ""
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
          header: "",
          closing: ""
        },
        FOOTER_TEXT: `\xA9 ${(/* @__PURE__ */ new Date()).getFullYear()} Your Company`
      };
    }
  });

  // src/common/utils/LayoutManager.ts
  var LayoutManager;
  var init_LayoutManager = __esm({
    "src/common/utils/LayoutManager.ts"() {
      init_SlideConfig();
      LayoutManager = class {
        constructor(pageW_pt, pageH_pt) {
          this.pxToPtFn = (px) => px * 0.75;
          const baseW_pt = this.pxToPt(CONFIG.BASE_PX.W);
          const baseH_pt = this.pxToPt(CONFIG.BASE_PX.H);
          this.scaleX = pageW_pt / baseW_pt;
          this.scaleY = pageH_pt / baseH_pt;
          this.pageW_pt = pageW_pt;
          this.pageH_pt = pageH_pt;
        }
        pxToPt(px) {
          return this.pxToPtFn(px);
        }
        getPositionFromPath(path) {
          return path.split(".").reduce((obj, key) => obj && obj[key], CONFIG.POS_PX);
        }
        getRect(spec) {
          const pos = typeof spec === "string" ? this.getPositionFromPath(spec) : spec;
          if (!pos) return { left: 0, top: 0, width: 0, height: 0 };
          let left_px = pos.left;
          if (pos.right !== void 0 && pos.left === void 0) {
            left_px = CONFIG.BASE_PX.W - pos.right - pos.width;
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
      };
    }
  });

  // src/common/utils/ColorUtils.ts
  function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
      r: parseInt(result[1], 16),
      g: parseInt(result[2], 16),
      b: parseInt(result[3], 16)
    } : null;
  }
  var init_ColorUtils = __esm({
    "src/common/utils/ColorUtils.ts"() {
    }
  });

  // src/common/utils/SlideUtils.ts
  function applyTextStyle(textRange, opt) {
    const style = textRange.getTextStyle();
    let defaultColor;
    if (opt.fontType === "large") {
      defaultColor = CONFIG.COLORS.text_primary;
    } else {
      defaultColor = CONFIG.COLORS.text_small_font;
    }
    style.setFontFamily(CONFIG.FONTS.family).setForegroundColor(opt.color || defaultColor).setFontSize(opt.size || CONFIG.FONTS.sizes.body).setBold(opt.bold || false);
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
  function setBulletsWithInlineStyles(shape, points) {
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
    tr.setText(combined || "\u2014");
    applyTextStyle(tr, {
      size: CONFIG.FONTS.sizes.body
    });
    try {
      tr.getParagraphs().forEach((p) => {
        p.getRange().getParagraphStyle().setLineSpacing(100).setSpaceBelow(6);
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
            color: CONFIG.COLORS.primary_color
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
            color: CONFIG.COLORS.primary_color
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
  function setBackgroundImageFromUrl(slide, layout, imageUrl, fallbackColor, imageUpdateOption = "update") {
    slide.getBackground().setSolidFill(fallbackColor);
    if (imageUpdateOption === "update") {
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
  function createGradientRectangle(slide, x, y, width, height, colors) {
    const numStrips = Math.max(20, Math.floor(width / 2));
    const stripWidth = width / numStrips;
    const startColor = hexToRgb(colors[0]), endColor = hexToRgb(colors[1]);
    if (!startColor || !endColor) return null;
    const shapes = [];
    for (let i = 0; i < numStrips; i++) {
      const ratio = i / (numStrips - 1);
      const r = Math.round(startColor.r + (endColor.r - startColor.r) * ratio);
      const g = Math.round(startColor.g + (endColor.g - startColor.g) * ratio);
      const b = Math.round(startColor.b + (endColor.b - startColor.b) * ratio);
      const strip = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x + i * stripWidth, y, stripWidth + 0.5, height);
      strip.getFill().setSolidFill(r, g, b);
      strip.getBorder().setTransparent();
      shapes.push(strip);
    }
    if (shapes.length > 1) {
      try {
        if (slide.group) return slide.group(shapes);
      } catch (e) {
      }
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
        if (slide.group) slide.group(shapes);
      } catch (e) {
      }
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
  function adjustShapeText_External(shape, preCalculatedWidthPt = null, widthOverride = null, heightOverride = null) {
    const PADDING_TOP_BOTTOM = 7.5;
    const PADDING_LEFT_RIGHT = 10;
    function getEffectiveCharCount(text) {
      let count = 0;
      if (!text) return 0;
      for (let i = 0; i < text.length; i++) {
        const char = text[i];
        if (char.match(/[^\x00-\x7F\uFF61-\uFF9F]/)) {
          count += 1;
        } else {
          count += 0.6;
        }
      }
      return count;
    }
    function _isShapeShortBox(shape2, baseFontSize, heightOverride2 = null) {
      if (!baseFontSize || baseFontSize === 0) {
        return false;
      }
      const boxHeight = heightOverride2 !== null ? heightOverride2 : shape2.getHeight();
      return boxHeight <= baseFontSize * 2;
    }
    return { isOverflow: false, details: "Simplified implementation in cleanup." };
  }
  function drawBottomBar(slide, layout, settings) {
    const barRect = layout.getRect("bottomBar");
    applyFill(slide, barRect.left, barRect.top, barRect.width, barRect.height, settings);
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
      const newLeft = layout.pageW_pt - rightMarginPt - newWidth;
      const topPt = layout.pxToPt(creditPosPx.top) * layout.scaleY;
      img.setLeft(newLeft).setTop(topPt).setWidth(newWidth).setHeight(newHeight).setLinkUrl(creditLink);
    } catch (e) {
    }
  }
  function addCucFooter(slide, layout, pageNum, settings, creditImageBlob) {
    const CREDIT_IMAGE_LINK = "https://note.com/majin_108";
    if (CONFIG.FOOTER_TEXT && CONFIG.FOOTER_TEXT.trim() !== "") {
      const leftRect = layout.getRect("footer.leftText");
      const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftRect.left, leftRect.top, leftRect.width, leftRect.height);
      const tr = leftShape.getText();
      tr.setText(CONFIG.FOOTER_TEXT);
      applyTextStyle(tr, {
        size: CONFIG.FONTS.sizes.footer,
        fontType: "large"
      });
      try {
        leftShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {
      }
    }
    if (creditImageBlob) {
      drawCreditImage(slide, layout, creditImageBlob, CREDIT_IMAGE_LINK);
    }
    if (pageNum > 0 && settings && settings.showPageNumber) {
      const rightRect = layout.getRect("footer.rightPage");
      const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightRect.left, rightRect.top, rightRect.width, rightRect.height);
      const tr = rightShape.getText();
      tr.setText(String(pageNum));
      applyTextStyle(tr, {
        size: CONFIG.FONTS.sizes.footer,
        color: CONFIG.COLORS.primary_color,
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
  function estimateTextWidthPt(text, fontSize) {
    let count = 0;
    if (!text) return 0;
    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      if (char.match(/[^\x00-\x7F\uFF61-\uFF9F]/)) {
        count += 1;
      } else {
        count += 0.6;
      }
    }
    return count * fontSize;
  }
  function drawStandardTitleHeader(slide, layout, key, title, settings, preCalculatedWidthPt = null, imageUpdateOption = "update") {
    if (imageUpdateOption === "update") {
      const logoRect = safeGetRect(layout, `${key}.headerLogo`);
      try {
        if (CONFIG.LOGOS.header && logoRect) {
          const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.header);
          if (imageData && typeof imageData !== "string") {
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
    const titleShape = slide.insertShape(
      SlidesApp.ShapeType.TEXT_BOX,
      titleRect.left,
      adjustedTop,
      titleRect.width,
      optimalHeight
    );
    setStyledText(titleShape, title || "", {
      size: initialFontSize,
      bold: true,
      fontType: "large"
    });
    try {
      titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {
    }
    if (settings.showTitleUnderline && title) {
      const uRect = safeGetRect(layout, `${key}.titleUnderline`);
      if (!uRect) {
        return;
      }
      let underlineWidthPt = 0;
      if (preCalculatedWidthPt !== null && preCalculatedWidthPt > 0) {
        underlineWidthPt = preCalculatedWidthPt;
      } else {
        underlineWidthPt = estimateTextWidthPt(title, initialFontSize);
      }
      const desiredWidthPt = underlineWidthPt + 10;
      const maxUnderlineWidth = layout.pageW_pt - uRect.left - layout.pxToPt(25);
      const finalWidth = Math.min(desiredWidthPt, maxUnderlineWidth);
      createPillShapeUnderline(slide, uRect.left, uRect.top, finalWidth, uRect.height, settings);
    }
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
      fontType: "large"
    });
    return layout.pxToPt(36);
  }
  function createContentCushion(slide, area, settings, layout) {
    const cushion = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, area.width, area.height);
    cushion.getFill().setSolidFill(CONFIG.COLORS.card_bg);
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
  function setBoldTextSize(shapeOrTextRange, size = 16) {
    let textRange;
    try {
      if (shapeOrTextRange && typeof shapeOrTextRange.getText === "function") {
        textRange = shapeOrTextRange.getText();
      } else if (shapeOrTextRange && typeof shapeOrTextRange.getRuns === "function") {
        textRange = shapeOrTextRange;
      } else {
        return;
      }
      if (!textRange || textRange.isEmpty()) {
        return;
      }
      const runs = textRange.getRuns();
      runs.forEach((run) => {
        const style = run.getTextStyle();
        if (style.isBold()) {
          style.setFontSize(size);
        }
      });
    } catch (e) {
    }
  }
  var init_SlideUtils = __esm({
    "src/common/utils/SlideUtils.ts"() {
      init_SlideConfig();
      init_ColorUtils();
    }
  });

  // src/infrastructure/gas/generators/GasTitleSlideGenerator.ts
  var GasTitleSlideGenerator;
  var init_GasTitleSlideGenerator = __esm({
    "src/infrastructure/gas/generators/GasTitleSlideGenerator.ts"() {
      init_SlideConfig();
      init_SlideUtils();
      GasTitleSlideGenerator = class {
        constructor(creditImageBlob) {
          this.creditImageBlob = creditImageBlob;
        }
        generate(slide, data, layout, pageNum, settings, imageUpdateOption = "update") {
          setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.title, CONFIG.COLORS.background_white, imageUpdateOption);
          if (imageUpdateOption === "update") {
            const logoRect = layout.getRect("titleSlide.logo");
            try {
              if (CONFIG.LOGOS.header) {
                const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.header);
                if (imageData && typeof imageData !== "string") {
                  const logo = slide.insertImage(imageData);
                  const aspect = logo.getHeight() / logo.getWidth();
                  logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * aspect);
                }
              }
            } catch (e) {
            }
          }
          const titleRect = layout.getRect("titleSlide.title");
          const newTop = (layout.pageH_pt - titleRect.height) / 2;
          const newWidth = titleRect.width + layout.pxToPt(60);
          const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, newTop, newWidth, titleRect.height);
          setStyledText(titleShape, data.title, {
            size: CONFIG.FONTS.sizes.title,
            bold: true,
            fontType: "large"
          });
          try {
            titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {
          }
          try {
            const titleText = data.title || "";
            if (titleText.indexOf("\n") === -1) {
              const preCalculatedWidth = data && typeof data._title_widthPt === "number" ? data._title_widthPt : null;
              if (preCalculatedWidth !== null && preCalculatedWidth < 900) {
                adjustShapeText_External(titleShape, preCalculatedWidth);
              } else {
                adjustShapeText_External(titleShape, null);
              }
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
            const dateRect = layout.getRect("titleSlide.date");
            const dateShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, dateRect.left, dateRect.top, dateRect.width, dateRect.height);
            dateShape.getText().setText(data.date || "");
            applyTextStyle(dateShape.getText(), {
              size: CONFIG.FONTS.sizes.date,
              fontType: "large"
            });
          }
          if (settings.showBottomBar) {
            drawBottomBar(slide, layout, settings);
          }
          if (this.creditImageBlob) {
            const CREDIT_IMAGE_LINK = "https://note.com/majin_108";
            drawCreditImage(slide, layout, this.creditImageBlob, CREDIT_IMAGE_LINK);
          }
        }
      };
    }
  });

  // src/infrastructure/gas/generators/GasSectionSlideGenerator.ts
  var GasSectionSlideGenerator;
  var init_GasSectionSlideGenerator = __esm({
    "src/infrastructure/gas/generators/GasSectionSlideGenerator.ts"() {
      init_SlideConfig();
      init_SlideUtils();
      GasSectionSlideGenerator = class {
        constructor(creditImageBlob) {
          this.creditImageBlob = creditImageBlob;
          this.sectionCounter = 0;
        }
        generate(slide, data, layout, pageNum, settings, imageUpdateOption = "update") {
          const imageUrl = CONFIG.BACKGROUND_IMAGES.section || CONFIG.BACKGROUND_IMAGES.main || "";
          let fallbackColor;
          if (imageUrl) {
            fallbackColor = CONFIG.COLORS.background_white;
          } else {
            if (CONFIG.COLORS.background_white.toUpperCase() !== "#FFFFFF") {
              fallbackColor = CONFIG.COLORS.background_white;
            } else {
              fallbackColor = CONFIG.COLORS.background_gray;
            }
          }
          setBackgroundImageFromUrl(slide, layout, imageUrl, fallbackColor, imageUpdateOption);
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
            ghostTextStyle.setFontFamily(CONFIG.FONTS.family).setFontSize(CONFIG.FONTS.sizes.ghostNum).setBold(true);
            try {
              ghostTextStyle.setForegroundColor(CONFIG.COLORS.ghost_gray);
            } catch (e) {
              ghostTextStyle.setForegroundColor(CONFIG.COLORS.ghost_gray);
            }
            try {
              ghost.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
          }
          const titleRect = layout.getRect("sectionSlide.title");
          const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, titleRect.top, titleRect.width, titleRect.height);
          titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          setStyledText(titleShape, data.title, {
            size: CONFIG.FONTS.sizes.sectionTitle,
            bold: true,
            align: SlidesApp.ParagraphAlignment.CENTER,
            fontType: "large"
          });
          try {
            const titleText = data.title || "";
            const textAreaWidthPt = titleRect.width;
            if (titleText.indexOf("\n") === -1) {
              const preCalculatedWidth = data && typeof data._title_widthPt === "number" ? data._title_widthPt : null;
              if (textAreaWidthPt > 0 && (preCalculatedWidth === null || preCalculatedWidth < textAreaWidthPt * 1.4)) {
                adjustShapeText_External(titleShape, preCalculatedWidth);
              }
            }
          } catch (e) {
          }
          addCucFooter(slide, layout, pageNum, settings, this.creditImageBlob);
        }
      };
    }
  });

  // src/infrastructure/gas/generators/GasContentSlideGenerator.ts
  var GasContentSlideGenerator;
  var init_GasContentSlideGenerator = __esm({
    "src/infrastructure/gas/generators/GasContentSlideGenerator.ts"() {
      init_SlideConfig();
      init_SlideUtils();
      GasContentSlideGenerator = class {
        constructor(creditImageBlob) {
          this.creditImageBlob = creditImageBlob;
        }
        generate(slide, data, layout, pageNum, settings, imageUpdateOption = "update") {
          setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.main, CONFIG.COLORS.background_white, imageUpdateOption);
          const titleWidthPt = data && typeof data._title_widthPt === "number" ? data._title_widthPt : null;
          drawStandardTitleHeader(slide, layout, "contentSlide", data.title, settings, titleWidthPt, imageUpdateOption);
          const subheadWidthPt = data && typeof data._subhead_widthPt === "number" ? data._subhead_widthPt : null;
          const dy = drawSubheadIfAny(slide, layout, "contentSlide", data.subhead, subheadWidthPt);
          let points = Array.isArray(data.points) ? data.points.slice(0) : [];
          const isAgenda = /(agenda|アジェンダ|目次|本日お伝えすること)/i.test(String(data.title || ""));
          if (isAgenda && points.length === 0) {
            points = ["\u672C\u65E5\u306E\u76EE\u7684", "\u9032\u3081\u65B9", "\u6B21\u306E\u30A2\u30AF\u30B7\u30E7\u30F3"];
          }
          const hasImages = Array.isArray(data.images) && data.images.length > 0;
          const isTwo = !!(data.twoColumn || data.columns);
          if (isTwo && (data.columns || points) || !isTwo && points && points.length > 0) {
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
              const baseLeftRect = layout.getRect("contentSlide.twoColLeft");
              const baseRightRect = layout.getRect("contentSlide.twoColRight");
              const adjustedLeftRect = adjustAreaForSubhead(baseLeftRect, data.subhead, layout);
              const adjustedRightRect = adjustAreaForSubhead(baseRightRect, data.subhead, layout);
              const leftRect = offsetRect(adjustedLeftRect, 0, dy);
              const rightRect = offsetRect(adjustedRightRect, 0, dy);
              createContentCushion(slide, leftRect, settings, layout);
              createContentCushion(slide, rightRect, settings, layout);
              const padding = layout.pxToPt(20);
              const leftTextRect = { left: leftRect.left + padding, top: leftRect.top + padding, width: leftRect.width - padding * 2, height: leftRect.height - padding * 2 };
              const rightTextRect = { left: rightRect.left + padding, top: rightRect.top + padding, width: rightRect.width - padding * 2, height: rightRect.height - padding * 2 };
              const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftTextRect.left, leftTextRect.top, leftTextRect.width, leftTextRect.height);
              const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightTextRect.left, rightTextRect.top, rightTextRect.width, rightTextRect.height);
              setBulletsWithInlineStyles(leftShape, L);
              setBulletsWithInlineStyles(rightShape, R);
              setBoldTextSize(leftShape, 16);
              setBoldTextSize(rightShape, 16);
              try {
                adjustShapeText_External(leftShape, null);
              } catch (e) {
              }
              try {
                adjustShapeText_External(rightShape, null);
              } catch (e) {
              }
            } else {
              const baseBodyRect = layout.getRect("contentSlide.body");
              const adjustedBodyRect = adjustAreaForSubhead(baseBodyRect, data.subhead, layout);
              const bodyRect = offsetRect(adjustedBodyRect, 0, dy);
              createContentCushion(slide, bodyRect, settings, layout);
              const padding = layout.pxToPt(20);
              const textRect = { left: bodyRect.left + padding, top: bodyRect.top + padding, width: bodyRect.width - padding * 2, height: bodyRect.height - padding * 2 };
              const bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
              setBulletsWithInlineStyles(bodyShape, points);
              setBoldTextSize(bodyShape, 16);
              try {
                bodyShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
              } catch (e) {
              }
              try {
                adjustShapeText_External(bodyShape, null);
              } catch (e) {
              }
            }
          }
          if (hasImages && !points.length && !isTwo) {
            const baseArea = layout.getRect("contentSlide.body");
            const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
            const area = offsetRect(adjustedArea, 0, dy);
            createContentCushion(slide, area, settings, layout);
            renderImagesInArea(slide, layout, area, normalizeImages(data.images), imageUpdateOption);
          }
          if (settings.showBottomBar) {
            drawBottomBar(slide, layout, settings);
          }
          addCucFooter(slide, layout, pageNum, settings, this.creditImageBlob);
        }
      };
    }
  });

  // src/infrastructure/gas/GasSlideRepository.ts
  var GasSlideRepository;
  var init_GasSlideRepository = __esm({
    "src/infrastructure/gas/GasSlideRepository.ts"() {
      init_LayoutManager();
      init_GasTitleSlideGenerator();
      init_GasSectionSlideGenerator();
      init_GasContentSlideGenerator();
      init_SlideConfig();
      GasSlideRepository = class {
        createPresentation(presentation, templateId) {
          const slidesApp = SlidesApp;
          const driveApp = DriveApp;
          let pres;
          if (templateId) {
            const templateFile = driveApp.getFileById(templateId);
            const newFile = templateFile.makeCopy(presentation.title);
            pres = slidesApp.openById(newFile.getId());
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
          const settings = {
            primaryColor: CONFIG.COLORS.primary_color,
            enableGradient: false,
            showTitleUnderline: true,
            showBottomBar: true,
            showDateColumn: true,
            showPageNumber: true,
            ...CONFIG.COLORS
          };
          presentation.slides.forEach((slideModel, index) => {
            const commonData = {
              title: slideModel.title.value,
              subtitle: slideModel.subtitle,
              date: (/* @__PURE__ */ new Date()).toLocaleDateString(),
              points: slideModel.content.items,
              // Map content items to points for Content/Agenda
              content: slideModel.content.items
              // Add images or other specific props if they exist in domain model (currently minimal)
            };
            const slide = pres.appendSlide(slidesApp.PredefinedLayout.BLANK);
            if (slideModel.layout === "TITLE") {
              titleGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else if (slideModel.layout === "SECTION") {
              sectionGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else if (slideModel.layout === "CONTENT" || slideModel.layout === "AGENDA") {
              contentGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            } else {
              contentGenerator.generate(slide, commonData, layoutManager, index + 1, settings);
            }
            if (slideModel.notes) {
              const notesPage = slide.getNotesPage();
              notesPage.getSpeakerNotesShape().getText().setText(slideModel.notes);
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
      };
    }
  });

  // src/api.ts
  var require_api = __commonJS({
    "src/api.ts"() {
      init_PresentationApplicationService();
      init_GasSlideRepository();
      function doPost(e) {
        try {
          const postData = JSON.parse(e.postData.contents);
          const data = postData.json || postData;
          const request = {
            title: data.title || "Untitled Presentation",
            templateId: data.templateId,
            // Optional ID for template
            slides: data.slides.map((s) => ({
              title: s.title,
              subtitle: s.subtitle,
              content: s.content || [],
              layout: s.layout,
              notes: s.notes
            }))
          };
          const repository = new GasSlideRepository();
          const service = new PresentationApplicationService(repository);
          const slideUrl = service.createPresentation(request);
          return createJsonResponse({
            success: true,
            url: slideUrl
          });
        } catch (error) {
          Logger.log("API Error: " + error.toString());
          return createJsonResponse({
            success: false,
            error: error.message || error.toString()
          });
        }
      }
      global.doPost = doPost;
      function createJsonResponse(data) {
        return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
      }
    }
  });
  require_api();
})();
