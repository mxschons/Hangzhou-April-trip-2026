/**
 * MxSchons Tours - Hangzhou Trip Document Builder
 * Using docx library with Jiangnan Style Guide
 *
 * Colors:
 * - Ink: #1C1917 (primary text, dark backgrounds)
 * - Rice Paper: #FEFDFB (backgrounds)
 * - Stone: #44403C (secondary text)
 * - Cinnabar: #C2675B (accent, highlights)
 * - Cinnabar Light: #E8ADA6 (subtle accents)
 * - Celadon: #8FABA0 (secondary accent)
 * - Celadon Dark: #6B8F82 (links)
 * - Bamboo: #D4CFC4 (borders)
 * - Mist: #F5F3EF (section backgrounds)
 */

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  ImageRun,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  WidthType,
  TableLayoutType,
  VerticalAlign,
  PageBreak,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
  ShadingType,
  convertInchesToTwip,
} = require("docx");
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

// ══════════════════════════════════════════════════════════════
// JIANGNAN COLOR PALETTE
// ══════════════════════════════════════════════════════════════

const COLORS = {
  ink: "1C1917",
  ricePaper: "FEFDFB",
  stone: "44403C",
  cinnabar: "C2675B",
  cinnabarDark: "A8524A",
  cinnabarLight: "E8ADA6",
  celadon: "8FABA0",
  celadonDark: "6B8F82",
  bamboo: "D4CFC4",
  mist: "F5F3EF",
  white: "FFFFFF",
};

// ══════════════════════════════════════════════════════════════
// TYPOGRAPHY SETTINGS
// Per style guide: Fraunces for display, Instrument Sans for body
// Fallbacks for systems without these fonts
// ══════════════════════════════════════════════════════════════

const FONTS = {
  display: "Georgia",           // Fraunces fallback (serif with warmth)
  body: "Avenir Next",          // Instrument Sans fallback (humanist sans)
  chinese: "PingFang SC",       // LXGW WenKai fallback
};

const SIZES = {
  display: 56,    // 28pt - Display headlines
  h1: 48,         // 24pt - Page titles
  h2: 36,         // 18pt - Section titles
  h3: 28,         // 14pt - Card titles
  h4: 24,         // 12pt - Small headers
  body: 22,       // 11pt - Default body
  small: 20,      // 10pt - Secondary text
  caption: 18,    // 9pt  - Labels
};

// ══════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ══════════════════════════════════════════════════════════════

function loadImage(filename) {
  const imagePath = path.join(__dirname, filename);
  if (fs.existsSync(imagePath)) {
    return fs.readFileSync(imagePath);
  }
  return null;
}

// Get image dimensions to maintain aspect ratio
function getImageDimensions(filename, maxWidth, maxHeight) {
  // For now, return square dimensions - in production you'd read actual image dimensions
  // This ensures images aren't distorted
  return { width: maxWidth, height: maxHeight };
}

function createSpacer(height = 200) {
  return new Paragraph({ spacing: { after: height } });
}

function createDivider() {
  return new Paragraph({
    children: [
      new TextRun({
        text: "───────────────────────────────────────",
        color: COLORS.bamboo,
        font: FONTS.body,
        size: SIZES.body,
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 200 },
  });
}

function createSectionLabel(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text: text.toUpperCase(),
        color: COLORS.cinnabar,
        font: FONTS.body,
        size: SIZES.caption,
        bold: true,
        characterSpacing: 40,
      }),
    ],
    spacing: { before: 400, after: 100 },
  });
}

function createHeading1(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text: text,
        color: COLORS.ink,
        font: FONTS.display,
        size: SIZES.h1,
        bold: true,
      }),
    ],
    spacing: { before: 400, after: 200 },
  });
}

function createHeading2(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text: text,
        color: COLORS.ink,
        font: FONTS.display,
        size: SIZES.h2,
        bold: true,
      }),
    ],
    spacing: { before: 300, after: 150 },
  });
}

function createHeading3(text, color = COLORS.stone) {
  return new Paragraph({
    children: [
      new TextRun({
        text: text,
        color: color,
        font: FONTS.body,
        size: SIZES.h3,
        bold: true,
      }),
    ],
    spacing: { before: 200, after: 100 },
  });
}

function createBodyText(text, options = {}) {
  const { italic = false, bold = false, color = COLORS.ink } = options;
  return new Paragraph({
    children: [
      new TextRun({
        text: text,
        color: color,
        font: FONTS.body,
        size: SIZES.body,
        italics: italic,
        bold: bold,
      }),
    ],
    spacing: { before: 100, after: 100 },
  });
}

function createItalicNote(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text: text,
        color: COLORS.stone,
        font: FONTS.body,
        size: SIZES.small,
        italics: true,
      }),
    ],
    spacing: { before: 100, after: 200 },
  });
}

function createBullet(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text: "◆ ",
        color: COLORS.cinnabar,
        font: FONTS.body,
        size: SIZES.body,
      }),
      new TextRun({
        text: text,
        color: COLORS.ink,
        font: FONTS.body,
        size: SIZES.body,
      }),
    ],
    spacing: { before: 50, after: 50 },
    indent: { left: convertInchesToTwip(0.25) },
  });
}

// ══════════════════════════════════════════════════════════════
// TABLE BUILDERS
// ══════════════════════════════════════════════════════════════

function createStyledTable(headers, rows, options = {}) {
  const { headerBg = COLORS.ink, headerText = COLORS.white, columnWidths = null } = options;

  // Calculate column widths - distribute evenly if not specified
  const numCols = headers.length;
  const defaultColWidth = Math.floor(9000 / numCols); // Total width ~9000 twips
  const colWidths = columnWidths || Array(numCols).fill(defaultColWidth);

  const headerRow = new TableRow({
    children: headers.map((header, idx) =>
      new TableCell({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: header,
                color: headerText,
                font: FONTS.body,
                size: SIZES.small,
                bold: true,
              }),
            ],
            alignment: AlignmentType.LEFT,
          }),
        ],
        width: { size: colWidths[idx], type: WidthType.DXA },
        shading: { fill: headerBg, type: ShadingType.CLEAR },
        margins: {
          top: convertInchesToTwip(0.1),
          bottom: convertInchesToTwip(0.1),
          left: convertInchesToTwip(0.15),
          right: convertInchesToTwip(0.15),
        },
        verticalAlign: VerticalAlign.CENTER,
      })
    ),
    tableHeader: true,
  });

  const dataRows = rows.map((row, rowIndex) =>
    new TableRow({
      children: row.map((cell, idx) =>
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: cell,
                  color: COLORS.ink,
                  font: FONTS.body,
                  size: SIZES.small,
                }),
              ],
            }),
          ],
          width: { size: colWidths[idx], type: WidthType.DXA },
          shading: {
            fill: rowIndex % 2 === 0 ? COLORS.ricePaper : COLORS.mist,
            type: ShadingType.CLEAR,
          },
          margins: {
            top: convertInchesToTwip(0.08),
            bottom: convertInchesToTwip(0.08),
            left: convertInchesToTwip(0.15),
            right: convertInchesToTwip(0.15),
          },
          verticalAlign: VerticalAlign.TOP,
        })
      ),
    })
  );

  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: colWidths,
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
      left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
      right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
    },
  });
}

function createScheduleTable(rows) {
  const timeColWidth = 1800;  // ~1.25 inches for time column
  const activityColWidth = 7200; // Rest for activity

  return new Table({
    rows: rows.map((row, index) =>
      new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: row.time,
                    color: COLORS.cinnabar,
                    font: FONTS.body,
                    size: SIZES.small,
                    bold: true,
                  }),
                ],
              }),
            ],
            width: { size: timeColWidth, type: WidthType.DXA },
            shading: { fill: index % 2 === 0 ? COLORS.ricePaper : COLORS.mist, type: ShadingType.CLEAR },
            margins: {
              top: convertInchesToTwip(0.08),
              bottom: convertInchesToTwip(0.08),
              left: convertInchesToTwip(0.15),
              right: convertInchesToTwip(0.1),
            },
            verticalAlign: VerticalAlign.TOP,
          }),
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: row.activity,
                    color: COLORS.ink,
                    font: FONTS.body,
                    size: SIZES.small,
                  }),
                ],
              }),
            ],
            width: { size: activityColWidth, type: WidthType.DXA },
            shading: { fill: index % 2 === 0 ? COLORS.ricePaper : COLORS.mist, type: ShadingType.CLEAR },
            margins: {
              top: convertInchesToTwip(0.08),
              bottom: convertInchesToTwip(0.08),
              left: convertInchesToTwip(0.15),
              right: convertInchesToTwip(0.15),
            },
            verticalAlign: VerticalAlign.TOP,
          }),
        ],
      })
    ),
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [timeColWidth, activityColWidth],
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.SINGLE, size: 12, color: COLORS.cinnabarLight },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
      insideVertical: { style: BorderStyle.NONE },
    },
  });
}

// ══════════════════════════════════════════════════════════════
// PAGE BUILDERS
// ══════════════════════════════════════════════════════════════

function buildCoverPage() {
  const children = [];

  // Logo
  const logoData = loadImage("logo-mxschons.png");
  if (logoData) {
    children.push(
      new Paragraph({
        children: [
          new ImageRun({
            data: logoData,
            transformation: { width: 120, height: 120 },
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 },
      })
    );
  }

  // Brand name
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "MxSchons Tours",
          color: COLORS.ink,
          font: FONTS.display,
          size: SIZES.h1,
          bold: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
    })
  );

  // Tagline
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "The invite-only travel agency for discerning adventurers",
          color: COLORS.stone,
          font: FONTS.body,
          size: SIZES.body,
          italics: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    })
  );

  children.push(createDivider());

  return children;
}

function buildWelcomePage() {
  const children = [];

  // Host photo
  const hostPhoto = loadImage("max-professional.JPG");
  if (hostPhoto) {
    children.push(
      new Paragraph({
        children: [
          new ImageRun({
            data: hostPhoto,
            transformation: { width: 200, height: 200 },
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 200 },
      })
    );
  }

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "A Personal Welcome from Your Host",
          color: COLORS.ink,
          font: FONTS.display,
          size: SIZES.h2,
          bold: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 300 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "Dear Esteemed Travelers,",
          color: COLORS.ink,
          font: FONTS.body,
          size: SIZES.body,
        }),
      ],
      spacing: { after: 200 },
    })
  );

  const welcomeText = [
    "On behalf of MxSchons Tours — the world's most exclusive invite-only travel agency (membership: 5 people, waiting list: also 5 people) — I am absolutely thrilled to present to you this meticulously crafted journey to Hangzhou!",
    "",
    "We have combined the best of the best to ensure you have an absolutely legendary experience. From soaring over West Lake in a helicopter (yes, really!) to picking tea leaves like ancient emperors, from watching 700,000 LEDs light up the night sky to getting lost in magical water towns — this trip has it all.",
    "",
    "Our team of experts (me, with help from Claude) has spent countless hours researching the finest restaurants, the most photogenic spots, and the perfect balance of adventure and relaxation. We've thought of everything: vegetarian options for certain picky eaters, massage breaks for tired feet, and strategic cafe stops for Chinese study sessions.",
    "",
    "This isn't just a trip. This is a MxSchons Tours Experience™.",
    "",
    "Pack your bags. Charge your cameras. Prepare your appetites.",
    "",
    "Hangzhou awaits!",
  ];

  welcomeText.forEach((line) => {
    if (line === "") {
      children.push(createSpacer(100));
    } else {
      children.push(createBodyText(line));
    }
  });

  children.push(createSpacer(200));

  children.push(
    new Paragraph({
      children: [
        new TextRun({ text: "With excitement and anticipation,", color: COLORS.ink, font: FONTS.body, size: SIZES.body }),
      ],
      spacing: { after: 200 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({ text: "Max", color: COLORS.ink, font: FONTS.display, size: SIZES.h3, bold: true }),
      ],
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "Founder, CEO, and Chief Adventure Officer",
          color: COLORS.stone,
          font: FONTS.body,
          size: SIZES.small,
          italics: true,
        }),
      ],
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "MxSchons Tours",
          color: COLORS.stone,
          font: FONTS.body,
          size: SIZES.small,
          italics: true,
        }),
      ],
      spacing: { after: 200 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "P.S. — No refunds. You're family.",
          color: COLORS.cinnabar,
          font: FONTS.body,
          size: SIZES.small,
          italics: true,
        }),
      ],
      spacing: { after: 400 },
    })
  );

  return children;
}

function buildTitlePage() {
  const children = [];

  // Chinese title
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "杭州家庭之旅",
          color: COLORS.ink,
          font: FONTS.chinese,
          size: SIZES.display,
          bold: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 100 },
    })
  );

  // English title
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "HANGZHOU FAMILY TRIP",
          color: COLORS.stone,
          font: FONTS.body,
          size: SIZES.h3,
          characterSpacing: 60,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    })
  );

  // Diamond
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "◆",
          color: COLORS.cinnabar,
          font: FONTS.body,
          size: SIZES.h2,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    })
  );

  // Date
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "April 13 – 27, 2026",
          color: COLORS.ink,
          font: FONTS.display,
          size: SIZES.h2,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
    })
  );

  // Subtitle
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "Ten Days Exploring West Lake, Tea Villages & Beyond",
          color: COLORS.stone,
          font: FONTS.body,
          size: SIZES.body,
          italics: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 300 },
    })
  );

  // Travelers
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "MAX • DION • MARGOT • ALEX & LYNN",
          color: COLORS.cinnabar,
          font: FONTS.body,
          size: SIZES.h4,
          bold: true,
          characterSpacing: 20,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    })
  );

  children.push(createDivider());

  return children;
}

function buildTravelersSection() {
  const children = [];

  children.push(createSectionLabel("Meet the Travelers"));

  // Create traveler photos table
  const travelers = [
    { name: "Max", city: "Frankfurt", photo: "max.png" },
    { name: "Dion", city: "Singapore", photo: "dion.jpg" },
    { name: "Margot", city: "Frankfurt", photo: "margot.png" },
    { name: "Alex", city: "Singapore", photo: "alex.png" },
    { name: "Lynn", city: "Singapore", photo: "lynn.png" },
  ];

  const cellWidth = { size: 20, type: WidthType.PERCENTAGE }; // Equal 20% each

  // Photo row
  const photoRow = new TableRow({
    children: travelers.map((t) => {
      const photoData = loadImage(t.photo);
      return new TableCell({
        children: [
          new Paragraph({
            children: photoData
              ? [new ImageRun({
                  data: photoData,
                  transformation: { width: 100, height: 100 },
                  type: "png",
                })]
              : [new TextRun({ text: "◆", size: SIZES.h1, color: COLORS.cinnabarLight })],
            alignment: AlignmentType.CENTER,
          }),
        ],
        width: cellWidth,
        borders: {
          top: { style: BorderStyle.NONE },
          bottom: { style: BorderStyle.NONE },
          left: { style: BorderStyle.NONE },
          right: { style: BorderStyle.NONE },
        },
        margins: {
          top: convertInchesToTwip(0.1),
          bottom: convertInchesToTwip(0.05),
          left: convertInchesToTwip(0.1),
          right: convertInchesToTwip(0.1)
        },
        verticalAlign: VerticalAlign.CENTER,
      });
    }),
  });

  // Name row
  const nameRow = new TableRow({
    children: travelers.map((t) =>
      new TableCell({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: t.name,
                color: COLORS.ink,
                font: FONTS.body,
                size: SIZES.body,
                bold: true,
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ],
        width: cellWidth,
        borders: {
          top: { style: BorderStyle.NONE },
          bottom: { style: BorderStyle.NONE },
          left: { style: BorderStyle.NONE },
          right: { style: BorderStyle.NONE },
        },
        margins: {
          top: convertInchesToTwip(0.02),
          bottom: convertInchesToTwip(0.02),
          left: convertInchesToTwip(0.1),
          right: convertInchesToTwip(0.1)
        },
      })
    ),
  });

  // City row
  const cityRow = new TableRow({
    children: travelers.map((t) =>
      new TableCell({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: t.city,
                color: COLORS.stone,
                font: FONTS.body,
                size: SIZES.small,
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ],
        width: cellWidth,
        borders: {
          top: { style: BorderStyle.NONE },
          bottom: { style: BorderStyle.NONE },
          left: { style: BorderStyle.NONE },
          right: { style: BorderStyle.NONE },
        },
        margins: {
          top: convertInchesToTwip(0.02),
          bottom: convertInchesToTwip(0.1),
          left: convertInchesToTwip(0.1),
          right: convertInchesToTwip(0.1)
        },
      })
    ),
  });

  children.push(
    new Table({
      rows: [photoRow, nameRow, cityRow],
      width: { size: 100, type: WidthType.PERCENTAGE },
      columnWidths: [1800, 1800, 1800, 1800, 1800], // Equal widths in twips
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.NONE },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
        insideHorizontal: { style: BorderStyle.NONE },
        insideVertical: { style: BorderStyle.NONE },
      },
      layout: TableLayoutType.FIXED,
    })
  );

  children.push(createSpacer(300));
  children.push(createDivider());

  return children;
}

function buildJourneyOverview() {
  const children = [];

  children.push(createSectionLabel("The Journey"));

  children.push(
    new Paragraph({
      children: [
        new TextRun({ text: "Frankfurt → Hangzhou → Wuzhen → Shanghai → Frankfurt", color: COLORS.ink, font: FONTS.body, size: SIZES.body }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 50 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({ text: "Singapore → Hangzhou → Wuzhen → Shanghai → Singapore", color: COLORS.stone, font: FONTS.body, size: SIZES.body }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 300 },
    })
  );

  children.push(createDivider());

  children.push(createHeading2("Complete Journey Overview"));

  children.push(
    createBodyText(
      "This trip brings the family together in Hangzhou from two directions: Max and Margot fly directly from Frankfurt, while Dion, Alex, and Lynn travel from Singapore. Everyone reunites in Hangzhou for the adventure, then travels together to Shanghai before parting ways."
    )
  );

  children.push(createSpacer(200));

  children.push(
    createStyledTable(
      ["Date", "Location", "Activity"],
      [
        ["Apr 13 (Mon)", "Frankfurt → Hangzhou", "Max + Margot fly direct to Hangzhou; arrive evening"],
        ["Apr 13 (Mon)", "Singapore → Hangzhou", "Dion + Alex & Lynn fly from Singapore; arrive evening"],
        ["Apr 13 (Mon)", "Hangzhou", "Family reunion at hotel; late dinner together"],
        ["Apr 14–18", "Hangzhou", "West Lake, tea picking, helicopter, G20 light show, local gems"],
        ["Apr 19 (Sun)", "Hangzhou → Wuzhen → Shanghai", "Scenic route via water town; arrive Shanghai evening"],
        ["Apr 20–21", "Shanghai", "French Concession, The Bund, Huangpu cruise, Yu Garden"],
        ["Apr 22 (Wed)", "Shanghai → Singapore", "Dion + Alex & Lynn fly to Singapore"],
        ["Apr 22 (Wed)", "Shanghai → Frankfurt", "Max + Margot fly back to Frankfurt"],
        ["Apr 23 (Thu)", "Frankfurt", "Max + Margot arrive home"],
      ],
      { columnWidths: [1500, 2700, 4800] } // Date, Location, Activity
    )
  );

  children.push(createSpacer(200));

  return children;
}

function buildAtAGlance() {
  const children = [];

  children.push(createSectionLabel("Trip Overview"));
  children.push(createHeading2("At a Glance"));

  children.push(
    createBodyText(
      "This ten-day journey brings together five family members across two generations to experience Hangzhou and Shanghai during peak spring season. The itinerary balances cultural depth, natural beauty, and modern China while accommodating different energy levels and interests."
    )
  );

  children.push(createSpacer(200));

  children.push(
    createStyledTable(
      ["", ""],
      [
        ["Dates", "Monday, April 13 – Wednesday, April 22, 2026 (China)"],
        ["Full Journey", "April 13–27 including Singapore time + return to Germany"],
        ["Travelers", "Max (DE), Dion (SG), Margot (DE), Alex & Lynn (SG)"],
        ["Accommodation", "Wulin Jingyu Tingyuan Hotel — Panorama Canal View Rooms"],
        ["Weather", "21–22°C highs, 12–13°C lows, occasional spring rain"],
        ["Dietary", "Options for vegan (Max); others omnivore"],
        ["Est. Budget", "~50,000 RMB total (~$6,800 USD) for all 5 people"],
      ],
      { headerBg: COLORS.celadonDark, columnWidths: [2200, 6800] }
    )
  );

  children.push(createSpacer(300));

  return children;
}

function buildAccommodation() {
  const children = [];

  children.push(createSectionLabel("Accommodation"));
  children.push(createHeading2("Wulin Jingyu Tingyuan Hotel"));

  children.push(
    createBodyText(
      "A refined boutique hotel blending traditional Jiangnan architecture with modern comfort. Located in the historic Wulin district, the hotel features elegant courtyard gardens, tea lounges, and rooms designed with classical Chinese aesthetics. Prime location for exploring both West Lake and the Grand Canal area."
    )
  );

  children.push(createSpacer(300));

  return children;
}

function buildHighlights() {
  const children = [];

  children.push(createSectionLabel("Highlights"));

  const highlights = [
    "Tea Leaf Picking at Meijiawu Village — Hands-on Longjing harvest during peak season",
    "Helicopter Flight Over Hangzhou — Route D covering West Lake, pagodas, and Qiantang River",
    "Qiantang River Night Cruise — G20 Summit light show with 700,000 LEDs",
    "Wuzhen Water Town — Illuminated canals en route to Shanghai",
    "Shanghai Exploration — French Concession, The Bund, Yu Garden, Huangpu River",
    "Gongyan Oriental Art Dinner — Immersive Han Dynasty dinner theater",
  ];

  highlights.forEach((h) => children.push(createBullet(h)));

  children.push(createSpacer(300));

  return children;
}

function buildDayItinerary(dayNum, date, title, description, schedule, note, addPageBreak = false) {
  const children = [];

  // Day header with colored background - using celadon for a softer look
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `Day ${dayNum} — ${date}`,
          color: COLORS.white,
          font: FONTS.body,
          size: SIZES.h3,
          bold: true,
        }),
      ],
      shading: { fill: COLORS.celadonDark, type: ShadingType.CLEAR },
      spacing: { before: 400 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: title.toUpperCase(),
          color: COLORS.cinnabar,
          font: FONTS.body,
          size: SIZES.caption,
          bold: true,
          characterSpacing: 40,
        }),
      ],
      spacing: { before: 150, after: 100 },
    })
  );

  children.push(createItalicNote(description));

  children.push(createScheduleTable(schedule));

  if (note) {
    children.push(createItalicNote(note));
  }

  children.push(createSpacer(200));

  // Add page break after each day if requested
  if (addPageBreak) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
  }

  return children;
}

// ══════════════════════════════════════════════════════════════
// MAIN DOCUMENT BUILDER
// ══════════════════════════════════════════════════════════════

async function buildDocument(options = {}) {
  const { includeBudget = true, outputSuffix = "" } = options;
  console.log(`Building Hangzhou Trip document${includeBudget ? "" : " (no budget)"}...`);

  const sections = [];

  // Cover page
  sections.push(...buildCoverPage());
  sections.push(new Paragraph({ children: [new PageBreak()] }));

  // Welcome page
  sections.push(...buildWelcomePage());
  sections.push(new Paragraph({ children: [new PageBreak()] }));

  // Title page
  sections.push(...buildTitlePage());

  // Travelers section
  sections.push(...buildTravelersSection());

  // Journey overview
  sections.push(...buildJourneyOverview());

  // At a glance
  sections.push(...buildAtAGlance());

  // Accommodation
  sections.push(...buildAccommodation());

  // Highlights
  sections.push(...buildHighlights());

  sections.push(new Paragraph({ children: [new PageBreak()] }));

  // Day-by-day itinerary
  sections.push(createHeading1("Day-by-Day Itinerary"));

  // Day 1
  sections.push(
    ...buildDayItinerary(
      1,
      "Monday, April 13",
      "Arrival",
      "A day of travel and reunion. The family converges on Hangzhou from two directions: Max and Margot fly direct from Frankfurt, while Dion, Alex, and Lynn travel from Singapore. Both groups arrive in the evening for an exciting reunion at the hotel and a late dinner together.",
      [
        { time: "Daytime", activity: "Max + Margot: Flight Frankfurt → Hangzhou (direct)" },
        { time: "Daytime", activity: "Dion + Alex & Lynn: Flight Singapore → Hangzhou" },
        { time: "Evening", activity: "Both groups arrive Hangzhou Xiaoshan Airport" },
        { time: "~8:00 PM", activity: "Transfer to Wulin Jingyu Tingyuan Hotel; family reunion!" },
        { time: "9:30 PM", activity: "Late dinner at 新白鹿 (Xin Bailu) — popular local chain, great intro to Hangzhou flavors" },
      ],
      "Pre-book separate airport transfers for each group (~280–400 RMB each). Coordinate arrival times for smooth reunion. Xin Bailu has vegetarian options for Max.",
      true // page break after
    )
  );

  // Day 2
  sections.push(
    ...buildDayItinerary(
      2,
      "Tuesday, April 14",
      "Gentle Start",
      "A recovery day after the travel. Max gets his first Chinese class in while others relax. The afternoon brings everyone together for the classic West Lake experience and an easy introduction to Hangzhou street food at the night market.",
      [
        { time: "9:00–10:30", activity: "Max: Chinese class at cafe. Margot relaxes; Alex & Lynn explore canal." },
        { time: "11:00", activity: "Group reunites. West Lake lakeside promenade stroll." },
        { time: "12:30", activity: "Lunch at 楼外楼 (Lou Wai Lou) — Hangzhou's most famous restaurant since 1848" },
        { time: "14:30", activity: "City God Pagoda for panoramic views" },
        { time: "17:30", activity: "Wulin Night Market — street food intro" },
      ],
      "Lou Wai Lou: Try the legendary West Lake Vinegar Fish and Dongpo Pork. Tofu dishes available for Max. Cafe: 河下咖啡 (Hexia Coffee) — canal-side, peaceful for study.",
      true // page break after
    )
  );

  // Day 3
  sections.push(
    ...buildDayItinerary(
      3,
      "Wednesday, April 15",
      "Nature & Tea",
      "The signature Hangzhou nature day. Morning starts with the beloved Nine Creeks walk through bamboo groves and shallow streams, followed by the highlight: hands-on tea picking at Meijiawu during peak Longjing harvest season. The day culminates with an immersive Han Dynasty dinner theater.",
      [
        { time: "9:00–10:30", activity: "Max: Chinese class at Longjing village cafe" },
        { time: "11:00–13:00", activity: "Nine Creeks Meandering Walk — shaded, scenic streams" },
        { time: "13:00", activity: "Lunch at 龙井草堂 (Longjing Caotang) — garden restaurant in the tea village" },
        { time: "14:30–17:00", activity: "Tea Picking at Meijiawu — picking, roasting, tasting (158 RMB/person)" },
        { time: "19:00", activity: "Dinner at Gongyan Oriental Art — Han Dynasty dinner theater (498 RMB/person)" },
      ],
      "April is peak Longjing harvest (Yuqian grade). The Meijiawu experience includes drone aerial photos. Longjing Caotang specializes in tea-infused local dishes.",
      true // page break after
    )
  );

  // Day 4
  sections.push(
    ...buildDayItinerary(
      4,
      "Thursday, April 16",
      "Big Experiences",
      "The most spectacular day of the trip. After a relaxed morning, the group takes a helicopter flight over West Lake for breathtaking aerial views. Evening brings the famous G20 Summit light show viewed from a river cruise — 700,000 LEDs illuminating 35 buildings.",
      [
        { time: "9:00–10:30", activity: "Max: Chinese class" },
        { time: "11:00–12:30", activity: "Group brunch at hotel or nearby cafe" },
        { time: "13:30", activity: "Depart for Xinlian Heliport (Xiaoshan District)" },
        { time: "14:30–16:00", activity: "Helicopter Flight — Route D (~25 min): West Lake, pagodas, river" },
        { time: "18:00", activity: "Early dinner at 知味观 (Zhiweiguan) — famous for dim sum and Hangzhou classics" },
        { time: "19:30", activity: "Qiantang River Night Cruise — Qianyin boat (168 RMB/person)" },
        { time: "20:30", activity: "G20 Light Show — 700,000 LEDs on 35 buildings" },
      ],
      "Helicopter: 2,580 RMB/person. Zhiweiguan (est. 1913): Try xiaolongbao, cat ear noodles, and shrimp-stuffed lotus root. Vegetable dumplings for Max.",
      true // page break after
    )
  );

  // Day 5
  sections.push(
    ...buildDayItinerary(
      5,
      "Friday, April 17",
      "Rest & Local Flavors",
      "A well-deserved rest day after the intensive helicopter and cruise day. Sleep in, enjoy a traditional breakfast, then indulge in affordable professional massages. Evening brings a special Hangzhou dining experience.",
      [
        { time: "Morning", activity: "Sleep in, relax at hotel" },
        { time: "11:00", activity: "Late brunch at 游埠豆浆 — savory soy milk, fried dough, scallion pancakes" },
        { time: "Afternoon", activity: "Group Massage — Yaoshi Blind Massage (~50–100 RMB/person)" },
        { time: "Evening", activity: "Dinner at 外婆家 (Grandma's Home) — beloved local chain, excellent value" },
      ],
      "Grandma's Home: Hugely popular for authentic home-style Hangzhou cooking. Try tea-smoked duck, braised pork belly, and stir-fried greens. Always has vegetable dishes. Expect a short wait — worth it.",
      true // page break after
    )
  );

  // Day 6
  sections.push(
    ...buildDayItinerary(
      6,
      "Saturday, April 18",
      "Morning Market & Local Gems",
      "An early start for the authentic morning market experience at Dama Long — 240 meters of old Hangzhou charm that disappears by noon. The rest of the day is free for revisiting favorite discoveries or exploring new corners of the city.",
      [
        { time: "7:00", activity: "Morning Market — Dama Long (大马弄) — 240m of authentic local life" },
        { time: "9:00–10:30", activity: "Max: Chinese class at cafe" },
        { time: "11:00", activity: "Leisurely walk around Grand Canal scenic area" },
        { time: "12:30", activity: "Lunch at discovered favorite or hotel restaurant" },
        { time: "Afternoon", activity: "Flexible: Shopping, cafe hopping, or see Optional Activities" },
        { time: "19:00", activity: "Farewell Hangzhou dinner at 朴竹 Pu Zhu (Michelin Green Star) or favorite spot" },
      ],
      "Dama Long opens ~6:30 AM, busiest 7–9 AM. Try 葱包桧 at 胡阿姨葱包桧. Pu Zhu: Hangzhou's first Michelin Green Star — elegant vegetarian fine dining if Max wants a special meal.",
      true // page break after
    )
  );

  // Day 7
  sections.push(
    ...buildDayItinerary(
      7,
      "Sunday, April 19",
      "Wuzhen & Shanghai",
      "A scenic travel day that doubles as an experience. The journey to Shanghai passes through the magical water town of Wuzhen — perfectly positioned to catch the famous illuminated night scene before continuing to Shanghai for the final leg of the China adventure.",
      [
        { time: "9:00", activity: "Check out of Wulin Jingyu Tingyuan Hotel; luggage in car" },
        { time: "9:30", activity: "Depart Hangzhou for Wuzhen (private car, ~1.5 hours)" },
        { time: "11:00", activity: "Arrive Wuzhen; store luggage at visitor center (free)" },
        { time: "11:30", activity: "Lunch in Wuzhen — local noodles, rice cakes" },
        { time: "13:00–18:00", activity: "Explore Xizha (West Gate) — canals, bridges, traditional architecture" },
        { time: "18:00", activity: "Lights come on at dusk — Wuzhen's legendary illuminated scene" },
        { time: "20:00", activity: "Depart Wuzhen for Shanghai (~1.5 hours)" },
        { time: "21:30", activity: "Arrive Shanghai; check into hotel; light supper nearby" },
      ],
      "Wuzhen entry: 150 RMB. Vegan options limited — try 定胜糕 (rice cake + red bean) and 青团 (green rice cakes). Max should bring backup snacks.",
      true // page break after
    )
  );

  // Day 8
  sections.push(
    ...buildDayItinerary(
      8,
      "Monday, April 20",
      "Shanghai Day 1",
      "A full day exploring China's most cosmopolitan city. The morning wanders through the tree-lined boulevards and Art Deco architecture of the French Concession. Evening brings the magical experience of The Bund at sunset followed by a Huangpu River cruise.",
      [
        { time: "9:00", activity: "Leisurely breakfast at hotel" },
        { time: "10:00–13:00", activity: "French Concession walk: Wukang Road → Fuxing Park → Tianzifang" },
        { time: "13:00", activity: "Lunch at Lost Heaven — stunning Yunnan cuisine in a beautiful heritage building" },
        { time: "15:00–17:00", activity: "Shanghai Museum (free, air-conditioned) or continued exploration" },
        { time: "17:30", activity: "The Bund promenade — sunset views, Pudong skyline" },
        { time: "19:00", activity: "Huangpu River Night Cruise (~135 RMB, 45 min)" },
        { time: "21:00", activity: "Dinner at 上海老饭店 (Shanghai Lao Fandian) — classic Shanghai cuisine since 1875" },
      ],
      "Lost Heaven: Beautiful setting, excellent mushroom and vegetable dishes for Max. Shanghai Lao Fandian: Try hongshao rou (red-braised pork), drunken chicken, and crab dishes in season.",
      true // page break after
    )
  );

  // Day 9
  sections.push(
    ...buildDayItinerary(
      9,
      "Tuesday, April 21",
      "Shanghai Day 2",
      "The final full day of the China adventure. Morning brings the historic Yu Garden at opening time before the crowds. The afternoon is flexible for last-minute shopping, exploring, or simply savoring the final hours in China.",
      [
        { time: "8:30", activity: "Early start — head to Yu Garden area" },
        { time: "9:00", activity: "Yu Garden (豫园) — arrive at opening for quietest experience (40 RMB)" },
        { time: "10:30", activity: "Brunch at 南翔馒头店 — legendary xiaolongbao since 1900" },
        { time: "12:00", activity: "Explore Yu Garden Bazaar — traditional shops, souvenirs" },
        { time: "Afternoon", activity: "Flexible: Nanjing Road shopping, more French Concession, or rest" },
        { time: "18:00", activity: "Final dinner at 功德林 (Godly Vegetarian) — Shanghai's oldest vegetarian since 1922" },
        { time: "Evening", activity: "Pack, early night before morning flight" },
      ],
      "Nanxiang: The original xiaolongbao restaurant — expect queues but worth it. Godly Vegetarian: A fitting final meal with both history and excellent food everyone can enjoy. Half-price Yu Garden entry for seniors 60+.",
      true // page break after
    )
  );

  // Day 10
  sections.push(
    ...buildDayItinerary(
      10,
      "Wednesday, April 22",
      "Farewell & Departure",
      "The final morning in China. After breakfast together, the family shares a bittersweet farewell at Pudong Airport before going separate ways: Dion, Alex, and Lynn head back to Singapore, while Max and Margot begin their long journey back to Frankfurt.",
      [
        { time: "6:30 AM", activity: "Wake up, final packing" },
        { time: "7:30 AM", activity: "Depart hotel for Pudong International Airport" },
        { time: "9:00 AM", activity: "Family farewell at airport" },
        { time: "Morning", activity: "Dion + Alex & Lynn: Flight Shanghai → Singapore" },
        { time: "Morning", activity: "Max + Margot: Flight Shanghai → Frankfurt" },
        { time: "Afternoon", activity: "Dion + Alex & Lynn arrive Singapore" },
        { time: "Next Day", activity: "Max + Margot arrive Frankfurt" },
      ],
      "Allow 2+ hours at Pudong Airport for international departure. Hotel near Pudong recommended for easiest morning."
    )
  );

  // Restaurant Guide
  sections.push(new Paragraph({ children: [new PageBreak()] }));
  sections.push(createSectionLabel("Restaurant Guide"));
  sections.push(createHeading2("Hangzhou Classics"));

  sections.push(
    createBodyText(
      "A mix of legendary establishments, local favorites, and vegetarian options. Reservations recommended for dinner at popular spots."
    )
  );

  sections.push(createSpacer(200));

  sections.push(
    createStyledTable(
      ["Restaurant", "Specialty", "Price"],
      [
        ["楼外楼 Lou Wai Lou", "West Lake Vinegar Fish, Dongpo Pork — est. 1848", "150–250/person"],
        ["知味观 Zhiweiguan", "Dim sum, xiaolongbao, cat ear noodles — est. 1913", "80–150/person"],
        ["外婆家 Grandma's Home", "Home-style Hangzhou — beloved local chain", "60–100/person"],
        ["新白鹿 Xin Bailu", "Popular local chain — great value, diverse menu", "50–80/person"],
        ["龙井草堂 Longjing Caotang", "Tea-infused dishes — garden setting in tea village", "100–180/person"],
        ["宫宴 Gongyan Oriental Art", "Han Dynasty dinner theater — immersive experience", "498/person"],
        ["朴竹 Pu Zhu", "Michelin Green Star — elegant vegetarian fine dining", "300–500/person"],
      ],
      { columnWidths: [2500, 4700, 1800] }
    )
  );

  sections.push(createSpacer(300));
  sections.push(createHeading2("Shanghai Highlights"));

  sections.push(
    createStyledTable(
      ["Restaurant", "Specialty", "Price"],
      [
        ["Lost Heaven", "Yunnan cuisine — stunning heritage building", "150–250/person"],
        ["上海老饭店 Lao Fandian", "Classic Shanghai — red-braised pork, crab — est. 1875", "120–200/person"],
        ["南翔馒头店 Nanxiang", "Original xiaolongbao — legendary since 1900", "50–100/person"],
        ["功德林 Godly Vegetarian", "Shanghai's oldest vegetarian — est. 1922", "70–120/person"],
      ],
      { columnWidths: [2500, 4700, 1800] }
    )
  );

  sections.push(createItalicNote("All restaurants have some vegetarian options. Lou Wai Lou, Zhiweiguan, and Grandma's Home have good tofu and vegetable dishes for Max."));

  // Budget section (optional)
  if (includeBudget) {
    sections.push(new Paragraph({ children: [new PageBreak()] }));
    sections.push(createSectionLabel("Budget Estimate"));
  sections.push(createHeading2("Trip Costs"));
  sections.push(createBodyText("All prices in RMB. Exchange rate: ~$1 USD = 7.3 RMB"));
  sections.push(createSpacer(200));

  sections.push(
    createStyledTable(
      ["Category", "Per Person", "×5", "Total"],
      [
        ["Wulin Jingyu Tingyuan Hotel (6 nights × 3 rooms)", "~700/night", "×3", "12,600"],
        ["Shanghai hotel (3 nights × 3 rooms)", "~800/night", "×3", "7,200"],
        ["Helicopter flight (Route D)", "2,580", "×5", "12,900"],
        ["Qiantang River night cruise", "168", "×5", "840"],
        ["Tea picking (Meijiawu)", "158", "×5", "790"],
        ["Gongyan dinner theater", "498", "×5", "2,490"],
        ["Wuzhen entry (full day)", "150", "×5", "750"],
        ["Huangpu River cruise", "135", "×5", "675"],
        ["Yu Garden entry", "40", "×5", "200"],
        ["Massage sessions (2×)", "150", "×5", "750"],
        ["Airport transfer Hangzhou", "~350", "group", "350"],
        ["Private car HZ→Wuzhen→Shanghai", "~1,200", "group", "1,200"],
        ["Local taxis/Didi (~10 days)", "~1,500", "group", "1,500"],
        ["Daily meals (~180/day × 9)", "1,620", "×5", "8,100"],
      ],
      { headerBg: COLORS.celadonDark, columnWidths: [4500, 1500, 1000, 2000] }
    )
  );

  sections.push(createSpacer(300));

  // Budget totals in highlighted box
  const budgetTotalColWidths = [5000, 4000];
  sections.push(
    new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "Base Trip Total", color: COLORS.white, font: FONTS.body, size: SIZES.body, bold: true })],
                }),
              ],
              width: { size: budgetTotalColWidths[0], type: WidthType.DXA },
              shading: { fill: COLORS.ink, type: ShadingType.CLEAR },
              margins: { top: 150, bottom: 150, left: 200, right: 200 },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "~50,000 RMB (~$6,800 USD)", color: COLORS.white, font: FONTS.body, size: SIZES.body, bold: true })],
                  alignment: AlignmentType.RIGHT,
                }),
              ],
              width: { size: budgetTotalColWidths[1], type: WidthType.DXA },
              shading: { fill: COLORS.ink, type: ShadingType.CLEAR },
              margins: { top: 150, bottom: 150, left: 200, right: 200 },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "With Optionals (photoshoot + Xixi)", color: COLORS.ink, font: FONTS.body, size: SIZES.body })],
                }),
              ],
              width: { size: budgetTotalColWidths[0], type: WidthType.DXA },
              shading: { fill: COLORS.mist, type: ShadingType.CLEAR },
              margins: { top: 150, bottom: 150, left: 200, right: 200 },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "~54,000 RMB (~$7,400 USD)", color: COLORS.ink, font: FONTS.body, size: SIZES.body })],
                  alignment: AlignmentType.RIGHT,
                }),
              ],
              width: { size: budgetTotalColWidths[1], type: WidthType.DXA },
              shading: { fill: COLORS.mist, type: ShadingType.CLEAR },
              margins: { top: 150, bottom: 150, left: 200, right: 200 },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "Per Person Average (base)", color: COLORS.cinnabar, font: FONTS.body, size: SIZES.body, bold: true })],
                }),
              ],
              width: { size: budgetTotalColWidths[0], type: WidthType.DXA },
              shading: { fill: COLORS.ricePaper, type: ShadingType.CLEAR },
              margins: { top: 150, bottom: 150, left: 200, right: 200 },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "~$1,360 USD / ~10,000 RMB", color: COLORS.cinnabar, font: FONTS.body, size: SIZES.body, bold: true })],
                  alignment: AlignmentType.RIGHT,
                }),
              ],
              width: { size: budgetTotalColWidths[1], type: WidthType.DXA },
              shading: { fill: COLORS.ricePaper, type: ShadingType.CLEAR },
              margins: { top: 150, bottom: 150, left: 200, right: 200 },
            }),
          ],
        }),
      ],
      width: { size: 100, type: WidthType.PERCENTAGE },
      columnWidths: budgetTotalColWidths,
      layout: TableLayoutType.FIXED,
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
        left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
        right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.bamboo },
        insideVertical: { style: BorderStyle.NONE },
      },
    })
  );

    sections.push(createItalicNote("Excludes international flights. Budget can be reduced by ~12,000 RMB by skipping helicopter."));
  } // end includeBudget

  // Practical Information
  sections.push(new Paragraph({ children: [new PageBreak()] }));
  sections.push(createSectionLabel("Practical Information"));

  sections.push(createHeading2("Key Contacts"));
  sections.push(
    createStyledTable(
      ["Service", "Contact"],
      [
        ["Tour Operator", "John Wu: WhatsApp +86 135 6716 1784"],
        ["Tour Guide", "Dannie (from previous trip)"],
        ["Wulin Jingyu Tingyuan Hotel", "To be confirmed"],
        ["Binjiang Pier (cruise)", "+86-571-85178197"],
      ],
      { headerBg: COLORS.celadonDark, columnWidths: [3500, 5500] }
    )
  );

  sections.push(createSpacer(300));
  sections.push(createHeading2("Booking Timeline"));

  const bookingItems = [
    "30 days before: Gongyan Oriental Art dinner reservation",
    "15 days before: Photoshoot booking (if desired)",
    "7 days before: Helicopter, cruise, Meijiawu tea experience",
    "2–3 days before: Confirm all reservations, check weather for helicopter",
  ];
  bookingItems.forEach((item) => sections.push(createBullet(item)));

  sections.push(createSpacer(300));
  sections.push(createHeading2("Useful Phrases"));
  sections.push(
    createStyledTable(
      ["English", "Pinyin", "Chinese"],
      [
        ["Lighter pressure", "qīng yī diǎn", "轻一点"],
        ["Harder pressure", "zhòng yī diǎn", "重一点"],
        ["Pure vegan", "chún sù", "纯素"],
        ["No meat, fish, eggs", "bù yào ròu, yú, dàn", "不要肉、鱼、蛋"],
        ["This is delicious!", "zhè ge hěn hǎo chī", "这个很好吃"],
      ],
      { headerBg: COLORS.celadonDark, columnWidths: [3000, 3000, 3000] }
    )
  );

  sections.push(createSpacer(300));
  sections.push(createHeading2("Packing Reminders"));

  const packingItems = [
    "Layers (15–23°C range, cooler evenings)",
    "Rain jacket or umbrella (7–14 rainy days typical in April)",
    "Comfortable walking shoes",
    "Power adapter (China: Type A/I, 220V)",
    "VPN installed (for Google, WhatsApp)",
    "WeChat and Alipay set up before arrival",
    "Backup vegan snacks for Max (Wuzhen especially)",
  ];
  packingItems.forEach((item) => sections.push(createBullet(item)));

  // Final page
  sections.push(new Paragraph({ children: [new PageBreak()] }));

  sections.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "◆",
          color: COLORS.cinnabar,
          font: FONTS.body,
          size: SIZES.display,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 800 },
    })
  );

  sections.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "一路平安",
          color: COLORS.ink,
          font: FONTS.chinese,
          size: SIZES.h1,
          bold: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 200 },
    })
  );

  sections.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "Safe Travels",
          color: COLORS.stone,
          font: FONTS.body,
          size: SIZES.h3,
          italics: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    })
  );

  sections.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "◆",
          color: COLORS.cinnabar,
          font: FONTS.body,
          size: SIZES.display,
        }),
      ],
      alignment: AlignmentType.CENTER,
    })
  );

  // Create document
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(1),
              right: convertInchesToTwip(1),
              bottom: convertInchesToTwip(1),
              left: convertInchesToTwip(1),
            },
          },
        },
        children: sections,
      },
    ],
  });

  // Write DOCX file
  const buffer = await Packer.toBuffer(doc);
  const baseName = `Hangzhou_Trip_Jiangnan_Style${outputSuffix}`;
  const docxPath = path.join(__dirname, `${baseName}.docx`);
  fs.writeFileSync(docxPath, buffer);

  console.log(`✓ Created: ${docxPath}`);
  console.log(`  File size: ${(buffer.length / 1024 / 1024).toFixed(2)} MB`);

  // Convert to PDF using pandoc (with LibreOffice backend if available, or direct)
  const pdfPath = path.join(__dirname, `${baseName}.pdf`);
  console.log("→ Converting to PDF...");

  try {
    // Try using soffice (LibreOffice) first - best quality for docx
    try {
      execSync(`soffice --headless --convert-to pdf --outdir "${__dirname}" "${docxPath}"`, {
        stdio: 'pipe',
        timeout: 60000
      });
      console.log(`✓ Created: ${pdfPath}`);
    } catch {
      // Fall back to pandoc with xelatex
      console.log("  (Using pandoc for PDF conversion...)");

      // Set up PATH to include TeX
      const env = { ...process.env };
      env.PATH = `/Library/TeX/texbin:${env.PATH}`;

      execSync(`pandoc "${docxPath}" -o "${pdfPath}" --pdf-engine=xelatex -V mainfont="PingFang SC"`, {
        stdio: 'pipe',
        timeout: 120000,
        env
      });
      console.log(`✓ Created: ${pdfPath}`);
    }

    // Get PDF file size
    if (fs.existsSync(pdfPath)) {
      const pdfStats = fs.statSync(pdfPath);
      console.log(`  File size: ${(pdfStats.size / 1024 / 1024).toFixed(2)} MB`);
    }
  } catch (err) {
    console.log(`⚠ PDF conversion failed: ${err.message}`);
    console.log("  You can manually convert the .docx file to PDF");
  }
}

// Run - build both versions
async function main() {
  console.log("");
  console.log("╔══════════════════════════════════════════════╗");
  console.log("║  MxSchons Tours - Document Builder           ║");
  console.log("║  Jiangnan Style Guide                        ║");
  console.log("╚══════════════════════════════════════════════╝");
  console.log("");

  // Build full version (with budget)
  await buildDocument({ includeBudget: true, outputSuffix: "" });

  console.log("");

  // Build version without budget
  await buildDocument({ includeBudget: false, outputSuffix: "_NoBudget" });

  console.log("");
  console.log("══════════════════════════════════════════════");
  console.log("  Build complete! 一路平安");
  console.log("══════════════════════════════════════════════");
  console.log("");
}

main().catch(console.error);
