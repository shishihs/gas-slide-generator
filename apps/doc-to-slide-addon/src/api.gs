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
        constructor(title, content, layout = "CONTENT", subtitle, notes, rawData) {
          this.title = title;
          this.content = content;
          this.layout = layout;
          this.subtitle = subtitle;
          this.notes = notes;
          this.rawData = rawData;
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
          // [LEGACY] These coordinates are now used primarily as fallbacks when template placeholders are missing.
          // We encourage using Google Slides Templates instead of modifying these values.
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
          primary_color: "#8FB130",
          deep_primary: "#526717",
          text_primary: "#333333",
          text_small_font: "#1F2937",
          background_white: "#FFFFFF",
          card_bg: "#f6e9f0",
          background_gray: "#F1F3F4",
          faint_gray: "#FAFAFA",
          ghost_gray: "#E0E0E0",
          table_header_bg: "#E8EAED",
          lane_border: "#DADCE0",
          card_border: "#DADCE0",
          neutral_gray: "#9AA0A6",
          process_arrow: "#8FB130"
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
        FOOTER_TEXT: ``
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

  // src/infrastructure/gas/generators/GasTitleSlideGenerator.ts
  var GasTitleSlideGenerator;
  var init_GasTitleSlideGenerator = __esm({
    "src/infrastructure/gas/generators/GasTitleSlideGenerator.ts"() {
      GasTitleSlideGenerator = class {
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
  var init_SlideUtils = __esm({
    "src/common/utils/SlideUtils.ts"() {
      init_SlideConfig();
      init_ColorUtils();
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
              size: CONFIG.FONTS.sizes.sectionTitle,
              bold: true,
              align: SlidesApp.ParagraphAlignment.CENTER,
              fontType: "large"
            });
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
      init_SlideUtils();
      GasContentSlideGenerator = class {
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
              points = ["\u672C\u65E5\u306E\u76EE\u7684", "\u9032\u3081\u65B9", "\u6B21\u306E\u30A2\u30AF\u30B7\u30E7\u30F3"];
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
            if (Array.isArray(data.columns) && data.columns.length === 2) {
            }
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
              const adjustedBodyRect = adjustAreaForSubhead(baseBodyRect, data.subhead, layout);
              const bodyRect = offsetRect(adjustedBodyRect, 0, dy);
              const padding = layout.pxToPt(20);
              const textRect = { left: bodyRect.left + padding, top: bodyRect.top + padding, width: bodyRect.width - padding * 2, height: bodyRect.height - padding * 2 };
              const bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
              setBulletsWithInlineStyles(bodyShape, points);
            }
          }
          if (hasImages && !points.length && !isTwo) {
            const baseArea = layout.getRect("contentSlide.body");
            const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
            const area = offsetRect(adjustedArea, 0, dy);
            createContentCushion(slide, area, settings, layout);
            renderImagesInArea(slide, layout, area, normalizeImages(data.images), imageUpdateOption);
          }
          addCucFooter(slide, layout, pageNum, settings, this.creditImageBlob);
        }
      };
    }
  });

  // src/infrastructure/gas/generators/GasDiagramSlideGenerator.ts
  var GasDiagramSlideGenerator;
  var init_GasDiagramSlideGenerator = __esm({
    "src/infrastructure/gas/generators/GasDiagramSlideGenerator.ts"() {
      init_SlideConfig();
      init_SlideUtils();
      init_ColorUtils();
      GasDiagramSlideGenerator = class {
        constructor(creditImageBlob) {
          this.creditImageBlob = creditImageBlob;
        }
        generate(slide, data, layout, pageNum, settings, imageUpdateOption = "update") {
          Logger.log(`Generating Diagram Slide: ${data.layout || data.type}`);
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
          const bodyPlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
          const workArea = bodyPlaceholder ? { left: bodyPlaceholder.getLeft(), top: bodyPlaceholder.getTop(), width: bodyPlaceholder.getWidth(), height: bodyPlaceholder.getHeight() } : layout.getRect("contentSlide.body");
          if (bodyPlaceholder) {
            try {
              bodyPlaceholder.remove();
            } catch (e) {
              Logger.log("Warning: Could not remove body placeholder: " + e);
            }
          }
          const elementsBefore = slide.getPageElements().map((e) => e.getObjectId());
          try {
            if (type.includes("timeline")) {
              this.drawTimeline(slide, data, workArea, settings, layout);
            } else if (type.includes("process")) {
              this.drawProcess(slide, data, workArea, settings, layout);
            } else if (type.includes("cycle")) {
              this.drawCycle(slide, data, workArea, settings, layout);
            } else if (type.includes("pyramid")) {
              this.drawPyramid(slide, data, workArea, settings, layout);
            } else if (type.includes("triangle")) {
              this.drawTriangle(slide, data, workArea, settings, layout);
            } else if (type.includes("statscompare")) {
              this.drawStatsCompare(slide, data, workArea, settings, layout);
            } else if (type.includes("barcompare")) {
              this.drawBarCompare(slide, data, workArea, settings, layout);
            } else if (type.includes("compare") || type.includes("kaizen")) {
              this.drawComparison(slide, data, workArea, settings, layout);
            } else if (type.includes("stepup") || type.includes("stair")) {
              this.drawStepUp(slide, data, workArea, settings, layout);
            } else if (type.includes("flowchart")) {
              this.drawFlowChart(slide, data, workArea, settings, layout);
            } else if (type.includes("diagram")) {
              this.drawLanes(slide, data, workArea, settings, layout);
            } else if (type.includes("cards") || type.includes("headercards")) {
              this.drawCards(slide, data, workArea, settings, layout);
            } else if (type.includes("kpi")) {
              this.drawKPI(slide, data, workArea, settings, layout);
            } else if (type.includes("table")) {
              this.drawTable(slide, data, workArea, settings, layout);
            } else if (type.includes("faq")) {
              this.drawFAQ(slide, data, workArea, settings, layout);
            } else if (type.includes("progress")) {
              this.drawProgress(slide, data, workArea, settings, layout);
            } else if (type.includes("quote")) {
              this.drawQuote(slide, data, workArea, settings, layout);
            } else if (type.includes("imagetext")) {
              this.drawImageText(slide, data, workArea, settings, layout);
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
          if (newElements.length > 1) {
            try {
              slide.group(newElements);
              Logger.log(`Grouped ${newElements.length} content elements for ${type}`);
            } catch (e) {
              Logger.log(`Warning: Could not group elements: ${e}`);
            }
          }
          addCucFooter(slide, layout, pageNum, settings, this.creditImageBlob);
        }
        drawTimeline(slide, data, area, settings, layout) {
          const milestones = data.milestones || data.items || [];
          if (!milestones.length) return;
          const inner = layout.pxToPt(80), baseY = area.top + area.height * 0.5;
          const leftX = area.left + inner, rightX = area.left + area.width - inner;
          const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, leftX, baseY, rightX, baseY);
          line.getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
          line.setWeight(2);
          const dotR = layout.pxToPt(10);
          const gap = milestones.length > 1 ? (rightX - leftX) / (milestones.length - 1) : 0;
          const cardW_pt = layout.pxToPt(180);
          const vOffset = layout.pxToPt(40);
          const headerHeight = layout.pxToPt(28);
          const bodyHeight = layout.pxToPt(80);
          const timelineColors = generateTimelineCardColors(settings.primaryColor, milestones.length);
          milestones.forEach((m, i) => {
            const x = leftX + gap * i;
            const isAbove = i % 2 === 0;
            const dateText = String(m.date || "");
            const labelText = String(m.label || m.state || "");
            const cardH_pt = headerHeight + bodyHeight;
            const cardLeft = x - cardW_pt / 2;
            const cardTop = isAbove ? baseY - vOffset - cardH_pt : baseY + vOffset;
            const headerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cardLeft, cardTop, cardW_pt, headerHeight);
            headerShape.getFill().setSolidFill(timelineColors[i]);
            headerShape.getBorder().getLineFill().setSolidFill(timelineColors[i]);
            const bodyShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cardLeft, cardTop + headerHeight, cardW_pt, bodyHeight);
            bodyShape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
            bodyShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
            const connectorY_start = isAbove ? cardTop + cardH_pt : baseY;
            const connectorY_end = isAbove ? baseY : cardTop;
            const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, connectorY_start, x, connectorY_end);
            connector.getLineFill().setSolidFill(CONFIG.COLORS.neutral_gray);
            connector.setWeight(1);
            const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - dotR / 2, baseY - dotR / 2, dotR, dotR);
            dot.getFill().setSolidFill(timelineColors[i]);
            dot.getBorder().setTransparent();
            const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cardLeft, cardTop, cardW_pt, headerHeight);
            setStyledText(headerTextShape, dateText, {
              size: CONFIG.FONTS.sizes.body,
              bold: true,
              color: CONFIG.COLORS.background_gray,
              align: SlidesApp.ParagraphAlignment.CENTER
            });
            try {
              headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
            let bodyFontSize = CONFIG.FONTS.sizes.body;
            const textLength = labelText.length;
            if (textLength > 40) bodyFontSize = 10;
            else if (textLength > 30) bodyFontSize = 11;
            else if (textLength > 20) bodyFontSize = 12;
            const bodyTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cardLeft, cardTop + headerHeight, cardW_pt, bodyHeight);
            setStyledText(bodyTextShape, labelText, {
              size: bodyFontSize,
              align: SlidesApp.ParagraphAlignment.CENTER
            });
            try {
              bodyTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
          });
        }
        drawProcess(slide, data, area, settings, layout) {
          const steps = data.steps || data.items || [];
          if (!steps.length) return;
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
          const boxHPt = layout.pxToPt(boxHPx), arrowHPt = layout.pxToPt(arrowHPx);
          const headerWPt = layout.pxToPt(120);
          const bodyLeft = area.left + headerWPt;
          const bodyWPt = area.width - headerWPt;
          for (let i = 0; i < n; i++) {
            const cleanText = String(steps[i] || "").replace(/^\s*\d+[\.\s]*/, "");
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
            } catch (e) {
            }
            const body = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, bodyLeft, currentY, bodyWPt, boxHPt);
            body.getFill().setSolidFill(CONFIG.COLORS.background_gray);
            body.getBorder().setTransparent();
            const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, bodyLeft + layout.pxToPt(20), currentY, bodyWPt - layout.pxToPt(40), boxHPt);
            setStyledText(textShape, cleanText, { size: fontSize });
            try {
              textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
            currentY += boxHPt;
            if (i < n - 1) {
              const arrowLeft = area.left + headerWPt / 2 - layout.pxToPt(8);
              const arrow = slide.insertShape(SlidesApp.ShapeType.DOWN_ARROW, arrowLeft, currentY, layout.pxToPt(16), arrowHPt);
              arrow.getFill().setSolidFill(CONFIG.COLORS.process_arrow || CONFIG.COLORS.ghost_gray);
              arrow.getBorder().setTransparent();
              currentY += arrowHPt;
            }
          }
        }
        drawCycle(slide, data, area, settings, layout) {
          const items = data.items || [];
          if (!items.length) return;
          const textLengths = items.map((item) => {
            const labelLength = (item.label || "").length;
            const subLabelLength = (item.subLabel || "").length;
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
            try {
              centerTextBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
          }
          const positions = [
            { x: centerX + radiusX, y: centerY },
            { x: centerX, y: centerY + radiusY },
            { x: centerX - radiusX, y: centerY },
            { x: centerX, y: centerY - radiusY }
          ];
          const itemsToDraw = items.slice(0, 4);
          itemsToDraw.forEach((item, i) => {
            const pos = positions[i];
            const cardX = pos.x - cardW / 2;
            const cardY = pos.y - cardH / 2;
            const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardW, cardH);
            card.getFill().setSolidFill(settings.primaryColor);
            card.getBorder().setTransparent();
            const subLabelText = item.subLabel || `${i + 1}\u756A\u76EE`;
            const labelText = item.label || "";
            setStyledText(card, `${subLabelText}
${labelText}`, { size: fontSize, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
            try {
              card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
              const textRange = card.getText();
              const subLabelEnd = subLabelText.length;
              if (textRange.asString().length > subLabelEnd) {
                textRange.getRange(0, subLabelEnd).getTextStyle().setFontSize(Math.max(10, fontSize - 2));
              }
            } catch (e) {
            }
          });
          const arrowRadiusX = radiusX * 0.75;
          const arrowRadiusY = radiusY * 0.8;
          const arrowSize = layout.pxToPt(80);
          const arrowPositions = [
            { left: centerX + arrowRadiusX, top: centerY - arrowRadiusY, rotation: 90 },
            { left: centerX + arrowRadiusX, top: centerY + arrowRadiusY, rotation: 180 },
            { left: centerX - arrowRadiusX, top: centerY + arrowRadiusY, rotation: 270 },
            { left: centerX - arrowRadiusX, top: centerY - arrowRadiusY, rotation: 0 }
          ];
          arrowPositions.slice(0, itemsToDraw.length).forEach((pos) => {
            const arrow = slide.insertShape(SlidesApp.ShapeType.BENT_ARROW, pos.left - arrowSize / 2, pos.top - arrowSize / 2, arrowSize, arrowSize);
            arrow.getFill().setSolidFill(CONFIG.COLORS.ghost_gray);
            arrow.getBorder().setTransparent();
            arrow.setRotation(pos.rotation);
          });
        }
        drawPyramid(slide, data, area, settings, layout) {
          const levels = data.levels || data.items || [];
          if (!levels.length) return;
          const levelsToDraw = levels.slice(0, 4);
          const levelHeight = layout.pxToPt(70);
          const levelGap = layout.pxToPt(2);
          const totalHeight = levelHeight * levelsToDraw.length + levelGap * (levelsToDraw.length - 1);
          const startY = area.top + (area.height - totalHeight) / 2;
          const pyramidWidth = layout.pxToPt(480);
          const textColumnWidth = layout.pxToPt(400);
          const gap = layout.pxToPt(30);
          const pyramidLeft = area.left;
          const textColumnLeft = pyramidLeft + pyramidWidth + gap;
          const pyramidColors = generatePyramidColors(settings.primaryColor, levelsToDraw.length);
          const baseWidth = pyramidWidth;
          const widthIncrement = baseWidth / levelsToDraw.length;
          const centerX = pyramidLeft + pyramidWidth / 2;
          levelsToDraw.forEach((level, index) => {
            const levelWidth = baseWidth - widthIncrement * (levelsToDraw.length - 1 - index);
            const levelX = centerX - levelWidth / 2;
            const levelY = startY + index * (levelHeight + levelGap);
            const levelBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, levelX, levelY, levelWidth, levelHeight);
            levelBox.getFill().setSolidFill(pyramidColors[index]);
            levelBox.getBorder().setTransparent();
            const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, levelX, levelY, levelWidth, levelHeight);
            titleShape.getFill().setTransparent();
            titleShape.getBorder().setTransparent();
            const levelTitle = level.title || `\u30EC\u30D9\u30EB${index + 1}`;
            setStyledText(titleShape, levelTitle, {
              size: CONFIG.FONTS.sizes.body,
              bold: true,
              color: CONFIG.COLORS.background_gray,
              align: SlidesApp.ParagraphAlignment.CENTER
            });
            try {
              titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
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
              connectionLine.getLineFill().setSolidFill("#D0D7DE");
              connectionLine.setWeight(1.5);
            }
            const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textColumnLeft, levelY, textColumnWidth, levelHeight);
            textShape.getFill().setTransparent();
            textShape.getBorder().setTransparent();
            const levelDesc = level.description || "";
            let formattedText;
            if (levelDesc.includes("\u2022") || levelDesc.includes("\u30FB")) {
              formattedText = levelDesc;
            } else if (levelDesc.includes("\n")) {
              formattedText = levelDesc.split("\n").filter((l) => l.trim()).slice(0, 2).map((l) => `\u2022 ${l.trim()}`).join("\n");
            } else {
              formattedText = levelDesc;
            }
            setStyledText(textShape, formattedText, {
              size: CONFIG.FONTS.sizes.body - 1,
              align: SlidesApp.ParagraphAlignment.START,
              // Fixed from LEFT
              color: CONFIG.COLORS.text_primary,
              bold: true
            });
            try {
              textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
          });
        }
        drawComparison(slide, data, area, settings, layout) {
          const leftTitle = data.leftTitle || "\u30D7\u30E9\u30F3A";
          const rightTitle = data.rightTitle || "\u30D7\u30E9\u30F3B";
          const leftItems = data.leftItems || [];
          const rightItems = data.rightItems || [];
          const gap = 20;
          const colWidth = (area.width - gap) / 2;
          const compareColors = generateCompareColors(settings.primaryColor);
          const headerH = layout.pxToPt(40);
          const leftHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, colWidth, headerH);
          leftHeader.getFill().setSolidFill(compareColors.left);
          leftHeader.getBorder().setTransparent();
          setStyledText(leftHeader, leftTitle, { size: 14, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
          try {
            leftHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {
          }
          const leftBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top + headerH, colWidth, area.height - headerH);
          leftBox.getFill().setSolidFill(CONFIG.COLORS.background_gray);
          leftBox.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
          setStyledText(leftBox, leftItems.join("\n\n"), { size: CONFIG.FONTS.sizes.body });
          const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left + colWidth + gap, area.top, colWidth, headerH);
          rightHeader.getFill().setSolidFill(compareColors.right);
          rightHeader.getBorder().setTransparent();
          setStyledText(rightHeader, rightTitle, { size: 14, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
          try {
            rightHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {
          }
          const rightBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left + colWidth + gap, area.top + headerH, colWidth, area.height - headerH);
          rightBox.getFill().setSolidFill(CONFIG.COLORS.background_gray);
          rightBox.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
          setStyledText(rightBox, rightItems.join("\n\n"), { size: CONFIG.FONTS.sizes.body });
        }
        drawStatsCompare(slide, data, area, settings, layout) {
          const leftTitle = data.leftTitle || "\u5C0E\u5165\u524D";
          const rightTitle = data.rightTitle || "\u5C0E\u5165\u5F8C";
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
          setStyledText(leftHeader, leftTitle, { size: 14, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
          try {
            leftHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {
          }
          const rightHeaderX = area.left + labelColW + valueColW;
          const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightHeaderX, area.top, valueColW, headerH);
          rightHeader.getFill().setSolidFill(compareColors.right);
          rightHeader.getBorder().setTransparent();
          setStyledText(rightHeader, rightTitle, { size: 14, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
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
            const rowBg = index % 2 === 0 ? CONFIG.COLORS.background_gray : "#FFFFFF";
            const labelCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, currentY, labelColW, rowHeight);
            labelCell.getFill().setSolidFill(rowBg);
            labelCell.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
            setStyledText(labelCell, label, { size: CONFIG.FONTS.sizes.body, bold: true, align: SlidesApp.ParagraphAlignment.START });
            try {
              labelCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
            const leftCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftHeaderX, currentY, valueColW, rowHeight);
            leftCell.getFill().setSolidFill(rowBg);
            leftCell.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
            setStyledText(leftCell, leftValue, { size: CONFIG.FONTS.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER, color: compareColors.left });
            try {
              leftCell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
            const rightCell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightHeaderX, currentY, valueColW - (trend ? layout.pxToPt(40) : 0), rowHeight);
            rightCell.getFill().setSolidFill(rowBg);
            rightCell.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
            setStyledText(rightCell, rightValue, { size: CONFIG.FONTS.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER, color: compareColors.right });
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
              setStyledText(trendShape, isUp ? "\u2191" : "\u2193", { size: 12, color: "#FFFFFF", bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
              try {
                trendShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
              } catch (e) {
              }
            }
            currentY += rowHeight;
          });
        }
        drawBarCompare(slide, data, area, settings, layout) {
          const leftTitle = data.leftTitle || "\u5C0E\u5165\u524D";
          const rightTitle = data.rightTitle || "\u5C0E\u5165\u5F8C";
          const stats = data.stats || [];
          if (!stats.length) return;
          const compareColors = generateCompareColors(settings.primaryColor);
          let maxValue = 0;
          stats.forEach((stat) => {
            const leftNum = parseFloat(String(stat.leftValue || "0").replace(/[^0-9.]/g, "")) || 0;
            const rightNum = parseFloat(String(stat.rightValue || "0").replace(/[^0-9.]/g, "")) || 0;
            maxValue = Math.max(maxValue, leftNum, rightNum);
          });
          if (maxValue === 0) maxValue = 100;
          const labelColW = area.width * 0.2;
          const barAreaW = area.width * 0.6;
          const valueColW = area.width * 0.1;
          const trendColW = area.width * 0.1;
          const rowHeight = Math.min(layout.pxToPt(80), area.height / stats.length);
          const barHeight = layout.pxToPt(18);
          const barGap = layout.pxToPt(4);
          let currentY = area.top;
          stats.forEach((stat, index) => {
            const label = stat.label || "";
            const leftValue = stat.leftValue || "";
            const rightValue = stat.rightValue || "";
            const trend = stat.trend || null;
            const leftNum = parseFloat(String(leftValue).replace(/[^0-9.]/g, "")) || 0;
            const rightNum = parseFloat(String(rightValue).replace(/[^0-9.]/g, "")) || 0;
            const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, currentY, labelColW, rowHeight);
            setStyledText(labelShape, label, { size: CONFIG.FONTS.sizes.body, bold: true, align: SlidesApp.ParagraphAlignment.START });
            try {
              labelShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
            const barLeft = area.left + labelColW;
            const barTop = currentY + (rowHeight - (barHeight * 2 + barGap)) / 2;
            const leftBarWidth = leftNum / maxValue * barAreaW;
            if (leftBarWidth > 0) {
              const leftBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barTop, leftBarWidth, barHeight);
              leftBar.getFill().setSolidFill(compareColors.left);
              leftBar.getBorder().setTransparent();
            }
            const leftLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + leftBarWidth + layout.pxToPt(5), barTop, layout.pxToPt(60), barHeight);
            setStyledText(leftLabel, leftValue, { size: 10, color: compareColors.left });
            const rightBarWidth = rightNum / maxValue * barAreaW;
            if (rightBarWidth > 0) {
              const rightBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barTop + barHeight + barGap, rightBarWidth, barHeight);
              rightBar.getFill().setSolidFill(compareColors.right);
              rightBar.getBorder().setTransparent();
            }
            const rightLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + rightBarWidth + layout.pxToPt(5), barTop + barHeight + barGap, layout.pxToPt(60), barHeight);
            setStyledText(rightLabel, rightValue, { size: 10, color: compareColors.right });
            if (trend) {
              const trendX = area.left + area.width - trendColW;
              const trendShape = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, trendX, currentY + rowHeight / 2 - layout.pxToPt(12), layout.pxToPt(24), layout.pxToPt(24));
              const isUp = trend.toLowerCase() === "up";
              const trendColor = isUp ? "#28a745" : "#dc3545";
              trendShape.getFill().setSolidFill(trendColor);
              trendShape.getBorder().setTransparent();
              setStyledText(trendShape, isUp ? "\u2191" : "\u2193", { size: 12, color: "#FFFFFF", bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
              try {
                trendShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
              } catch (e) {
              }
            }
            if (index < stats.length - 1) {
              const lineY = currentY + rowHeight;
              const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, area.left, lineY, area.left + area.width, lineY);
              line.getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
              line.setWeight(1);
            }
            currentY += rowHeight;
          });
        }
        drawTriangle(slide, data, area, settings, layout) {
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
          trianglePath.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
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
              color: CONFIG.COLORS.background_gray,
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
        drawStepUp(slide, data, area, settings, layout) {
          const items = data.items || [];
          if (!items.length) return;
          const count = items.length;
          const stepWidth = area.width / count;
          const stepHeight = area.height / count;
          items.forEach((item, i) => {
            const h = (i + 1) * stepHeight;
            const x = area.left + i * stepWidth;
            const y = area.top + (area.height - h);
            const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, stepWidth - layout.pxToPt(5), h);
            shape.getFill().setSolidFill(settings.primaryColor, 0.5 + i * 0.1);
            shape.getBorder().setTransparent();
            const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y, stepWidth - layout.pxToPt(5), layout.pxToPt(50));
            setStyledText(textBox, item.title || "", { color: "#FFFFFF", bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
          });
        }
        drawLanes(slide, data, area, settings, layout) {
          const lanes = data.lanes || [];
          const n = Math.max(1, lanes.length);
          const { laneGap_px, lanePad_px, laneTitle_h_px, cardGap_px, cardMin_h_px, cardMax_h_px, arrow_h_px, arrowGap_px } = CONFIG.DIAGRAM;
          const px = (p) => layout.pxToPt(p);
          const laneW = (area.width - px(laneGap_px) * (n - 1)) / n;
          const cardBoxes = [];
          for (let j = 0; j < n; j++) {
            const lane = lanes[j] || { title: "", items: [] };
            const left = area.left + j * (laneW + px(laneGap_px));
            const lt = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, area.top, laneW, px(laneTitle_h_px));
            lt.getFill().setSolidFill(settings.primaryColor);
            lt.getBorder().getLineFill().setSolidFill(settings.primaryColor);
            setStyledText(lt, lane.title || "", { size: CONFIG.FONTS.sizes.laneTitle, bold: true, color: CONFIG.COLORS.background_gray, align: SlidesApp.ParagraphAlignment.CENTER });
            try {
              lt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
            const laneBodyTop = area.top + px(laneTitle_h_px);
            const laneBodyHeight = area.height - px(laneTitle_h_px);
            const laneBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, laneBodyTop, laneW, laneBodyHeight);
            laneBg.getFill().setSolidFill(CONFIG.COLORS.background_gray);
            laneBg.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);
            const items = Array.isArray(lane.items) ? lane.items : [];
            const rows = Math.max(1, items.length);
            const availH = laneBodyHeight - px(lanePad_px) * 2;
            const idealH = (availH - px(cardGap_px) * (rows - 1)) / rows;
            const cardH = Math.max(px(cardMin_h_px), Math.min(px(cardMax_h_px), idealH));
            const firstTop = laneBodyTop + px(lanePad_px) + Math.max(0, (availH - (cardH * rows + px(cardGap_px) * (rows - 1))) / 2);
            cardBoxes[j] = [];
            for (let i = 0; i < rows; i++) {
              const cardTop = firstTop + i * (cardH + px(cardGap_px));
              const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, left + px(lanePad_px), cardTop, laneW - px(lanePad_px) * 2, cardH);
              card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
              card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
              setStyledText(card, items[i] || "", { size: CONFIG.FONTS.sizes.body });
              try {
                card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
              } catch (e) {
              }
              cardBoxes[j][i] = {
                left: left + px(lanePad_px),
                top: cardTop,
                width: laneW - px(lanePad_px) * 2,
                height: cardH
              };
            }
          }
          const maxRows = Math.max(0, ...cardBoxes.map((a) => a ? a.length : 0));
          for (let j = 0; j < n - 1; j++) {
            for (let i = 0; i < maxRows; i++) {
              if (cardBoxes[j] && cardBoxes[j][i] && cardBoxes[j + 1] && cardBoxes[j + 1][i]) {
                drawArrowBetweenRects(slide, cardBoxes[j][i], cardBoxes[j + 1][i], px(arrow_h_px), px(arrowGap_px), settings);
              }
            }
          }
        }
        drawFlowChart(slide, data, area, settings, layout) {
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
            shape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
            shape.getBorder().getLineFill().setSolidFill(settings.primaryColor);
            shape.getBorder().setWeight(2);
            setStyledText(shape, typeof step === "string" ? step : step.label || "", { size: CONFIG.FONTS.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER });
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
              line.getLineFill().setSolidFill(CONFIG.COLORS.neutral_gray);
            }
          });
        }
        drawCards(slide, data, area, settings, layout) {
          const items = data.items || [];
          if (!items.length) return;
          const type = (data.type || "").toLowerCase();
          const hasHeader = type.includes("headercards");
          const cols = data.columns || 3;
          const rows = Math.ceil(items.length / cols);
          const gap = layout.pxToPt(20);
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
              const headerH = layout.pxToPt(36);
              const header = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, cardW, headerH);
              header.getFill().setSolidFill(settings.primaryColor);
              header.getBorder().setTransparent();
              setStyledText(header, title, { size: 14, bold: true, color: "#FFFFFF", align: SlidesApp.ParagraphAlignment.CENTER });
              try {
                header.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
              } catch (e) {
              }
              const body = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y + headerH, cardW, cardH - headerH);
              body.getFill().setSolidFill("#FFFFFF");
              body.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
              const textArea = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + layout.pxToPt(10), y + headerH + layout.pxToPt(10), cardW - layout.pxToPt(20), cardH - headerH - layout.pxToPt(20));
              setStyledText(textArea, desc, { size: 12, align: SlidesApp.ParagraphAlignment.START, color: CONFIG.COLORS.text_small_font });
            } else {
              const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, cardW, cardH);
              card.getFill().setSolidFill("#FFFFFF");
              card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
              const strip = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y + cardH * 0.1, layout.pxToPt(6), cardH * 0.8);
              strip.getFill().setSolidFill(settings.primaryColor);
              strip.getBorder().setTransparent();
              const contentX = x + layout.pxToPt(20);
              const contentW = cardW - layout.pxToPt(30);
              const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y + layout.pxToPt(10), contentW, layout.pxToPt(30));
              setStyledText(titleBox, title, { size: 16, bold: true, color: CONFIG.COLORS.text_primary });
              const descBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, y + layout.pxToPt(40), contentW, cardH - layout.pxToPt(50));
              setStyledText(descBox, desc, { size: 12, color: CONFIG.COLORS.text_small_font });
              try {
                descBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
              } catch (e) {
              }
            }
          });
        }
        drawKPI(slide, data, area, settings, layout) {
          const items = data.items || [];
          if (!items.length) return;
          const cols = items.length > 4 ? 4 : items.length || 1;
          const gap = layout.pxToPt(20);
          const cardW = (area.width - gap * (cols - 1)) / cols;
          const cardH = layout.pxToPt(160);
          const y = area.top + (area.height - cardH) / 2;
          items.forEach((item, i) => {
            const x = area.left + i * (cardW + gap);
            const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, cardW, cardH);
            card.getFill().setSolidFill("#FFFFFF");
            card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
            const padding = layout.pxToPt(10);
            const labelH = layout.pxToPt(30);
            const labelBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, y + padding, cardW - padding * 2, labelH);
            setStyledText(labelBox, item.label || "Metric", { size: 14, color: CONFIG.COLORS.neutral_gray, align: SlidesApp.ParagraphAlignment.CENTER });
            const valueH = layout.pxToPt(70);
            const valueBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, y + labelH + padding, cardW - padding * 2, valueH);
            const valStr = String(item.value || "0");
            let fontSize = 48;
            if (valStr.length > 6) fontSize = 36;
            if (valStr.length > 10) fontSize = 28;
            setStyledText(valueBox, valStr, { size: fontSize, bold: true, color: settings.primaryColor, align: SlidesApp.ParagraphAlignment.CENTER });
            if (item.change || item.status) {
              const statusH = layout.pxToPt(30);
              const statusBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, y + labelH + valueH + padding, cardW - padding * 2, statusH);
              let color = CONFIG.COLORS.neutral_gray;
              let prefix = "";
              if (item.status === "good") {
                color = "#28a745";
                prefix = "\u25B2 ";
              }
              if (item.status === "bad") {
                color = "#dc3545";
                prefix = "\u25BC ";
              }
              setStyledText(statusBox, prefix + (item.change || ""), { size: 14, bold: true, color, align: SlidesApp.ParagraphAlignment.CENTER });
            }
          });
        }
        drawTable(slide, data, area, settings, layout) {
          const headers = data.headers || [];
          const rows = data.rows || [];
          const numRows = rows.length + (headers.length ? 1 : 0);
          const numCols = headers.length || (rows[0] ? rows[0].length : 1);
          if (numRows === 0 || numCols === 0) return;
          const table = slide.insertTable(numRows, numCols);
          table.setLeft(area.left);
          table.setTop(area.top);
          table.setWidth(area.width);
          let rowIndex = 0;
          if (headers.length) {
            for (let c = 0; c < numCols; c++) {
              const cell = table.getCell(0, c);
              cell.getFill().setSolidFill(settings.primaryColor);
              cell.getText().setText(headers[c] || "");
              const style = cell.getText().getTextStyle();
              style.setBold(true);
              style.setFontSize(14);
              style.setForegroundColor("#FFFFFF");
              try {
                cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
              } catch (e) {
              }
            }
            rowIndex++;
          }
          rows.forEach((row, rIdx) => {
            const isAlt = rIdx % 2 !== 0;
            const rowColor = isAlt ? CONFIG.COLORS.faint_gray : "#FFFFFF";
            for (let c = 0; c < numCols; c++) {
              const cell = table.getCell(rowIndex, c);
              cell.getFill().setSolidFill(rowColor);
              cell.getText().setText(String(row[c] || ""));
              const rowStyle = cell.getText().getTextStyle();
              rowStyle.setFontSize(12);
              rowStyle.setForegroundColor(CONFIG.COLORS.text_primary);
              try {
                cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
              } catch (e) {
              }
            }
            rowIndex++;
          });
        }
        drawFAQ(slide, data, area, settings, layout) {
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
          const gap = layout.pxToPt(20);
          const itemH = (area.height - gap * (parsedItems.length - 1)) / parsedItems.length;
          parsedItems.forEach((item, i) => {
            const y = area.top + i * (itemH + gap);
            const iconSize = layout.pxToPt(40);
            const qCircle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, area.left, y + (itemH - iconSize) / 2, iconSize, iconSize);
            qCircle.getFill().setSolidFill(settings.primaryColor);
            qCircle.getBorder().setTransparent();
            setStyledText(qCircle, "Q", { size: 18, bold: true, color: "#FFFFFF", align: SlidesApp.ParagraphAlignment.CENTER });
            try {
              qCircle.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
            const boxLeft = area.left + iconSize + layout.pxToPt(15);
            const boxW = area.width - (iconSize + layout.pxToPt(15));
            const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, boxLeft, y, boxW, itemH);
            box.getFill().setSolidFill("#FFFFFF");
            box.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
            const qText = (item.q || "").replace(/^[QA][:. ]+/, "");
            const aText = (item.a || "").replace(/^[QA][:. ]+/, "");
            setStyledText(box, `Q. ${qText}

A. ${aText}`, { size: 12, color: CONFIG.COLORS.text_primary });
          });
        }
        drawQuote(slide, data, area, settings, layout) {
          const text = data.text || data.points && data.points[0] || "";
          const author = data.author || data.points && data.points[1] || "";
          const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, area.width, area.height);
          bg.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
          bg.getBorder().setTransparent();
          const quoteMark = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, area.top - layout.pxToPt(20), layout.pxToPt(100), layout.pxToPt(100));
          setStyledText(quoteMark, "\u201C", { size: 120, color: CONFIG.COLORS.ghost_gray, font: "Georgia" });
          const contentW = area.width * 0.8;
          const contentX = area.left + (area.width - contentW) / 2;
          const quoteBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, area.top, contentW, area.height - layout.pxToPt(60));
          setStyledText(quoteBox, text, { size: 28, bold: true, color: settings.primaryColor, align: SlidesApp.ParagraphAlignment.CENTER, font: "Serif" });
          try {
            quoteBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {
          }
          if (author) {
            const authorBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentX, area.top + area.height - layout.pxToPt(60), contentW, layout.pxToPt(40));
            setStyledText(authorBox, `\u2014 ${author}`, { size: 16, align: SlidesApp.ParagraphAlignment.END, color: CONFIG.COLORS.neutral_gray });
          }
        }
        drawProgress(slide, data, area, settings, layout) {
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
            barBg.getFill().setSolidFill(CONFIG.COLORS.ghost_gray);
            barBg.getBorder().setTransparent();
            const percent = Math.min(100, Math.max(0, parseInt(item.percent || 0)));
            if (percent > 0) {
              const barFg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left + labelW + layout.pxToPt(20), y + rowH / 3, barAreaW * (percent / 100), rowH / 3);
              barFg.getFill().setSolidFill(settings.primaryColor);
              barFg.getBorder().setTransparent();
            }
            const valBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + labelW + barAreaW + layout.pxToPt(30), y, layout.pxToPt(50), rowH);
            setStyledText(valBox, `${percent}%`, { size: 14, color: CONFIG.COLORS.neutral_gray });
            try {
              valBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
            } catch (e) {
            }
          });
        }
        drawImageText(slide, data, area, settings, layout) {
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
              } else if (imageUrl.startsWith("http")) {
                img = slide.insertImage(imageUrl);
              }
              if (img) {
                img.setLeft(imgX).setTop(area.top).setWidth(halfW).setHeight(area.height);
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
              ph.getFill().setSolidFill(CONFIG.COLORS.ghost_gray);
              setStyledText(ph, "Image Placeholder", { align: SlidesApp.ParagraphAlignment.CENTER });
            }
          }
          const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, txtX, area.top, halfW, area.height);
          const textContent = points.join("\n");
          setStyledText(textBox, textContent, { size: CONFIG.FONTS.sizes.body });
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
      init_GasDiagramSlideGenerator();
      init_SlideConfig();
      GasSlideRepository = class {
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
          const settings = {
            primaryColor: CONFIG.COLORS.primary_color,
            enableGradient: false,
            showTitleUnderline: true,
            showBottomBar: true,
            showDateColumn: true,
            showPageNumber: true,
            ...CONFIG.COLORS,
            ...settingsOverride && settingsOverride.colors ? {
              primaryColor: settingsOverride.colors.primary,
              primary_color: settingsOverride.colors.primary,
              text_primary: settingsOverride.colors.text
            } : {}
          };
          presentation.slides.forEach((slideModel, index) => {
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
            const rawType = (slideModel.rawData?.type || "").toLowerCase();
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
      };
    }
  });

  // src/api.ts
  var require_api = __commonJS({
    "src/api.ts"() {
      init_PresentationApplicationService();
      init_GasSlideRepository();
      global.doPost = doPost;
      global.doGet = doGet;
      global.generateSlides = generateSlides;
      function generateSlides(data) {
        try {
          Logger.log("Library Call: generateSlides with data: " + JSON.stringify(data));
          const request = {
            title: data.title || "\u7121\u984C\u306E\u30D7\u30EC\u30BC\u30F3\u30C6\u30FC\u30B7\u30E7\u30F3",
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
      global.doPost = doPost;
      function createJsonResponse(data) {
        return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
      }
    }
  });
  require_api();
})();
function generateSlides(data){ return global.generateSlides(data); } function doPost(e){ return global.doPost(e); } function doGet(e){ return global.doGet(e); }
