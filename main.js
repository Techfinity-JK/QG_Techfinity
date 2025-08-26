const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const { Document, Section, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Header, Table, TableRow, TableCell, WidthType, PageMargin, PageOrientation, HeightRule} = require("docx");


function createWindow() {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false, // allow ipcRenderer in index.html
    },
  });

  win.loadFile("index.html");
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

// Listen for generate-docx event
ipcMain.on("generate-docx", async (event, details) => {
  const {
    client = "-",
    address = "-",
    person = "-",
    number = "-",
    email = "-",
    product,
    quantity, // to be used later
    cost,      // 👈 user override
  } = details;

   // Map product values
  const productValues = {
    LX50:     {image: path.join(__dirname, "assets/img/device/lx50.png"),
               undiscounted: "10,900.00",
               price: 5700, 
               description:`~ 500 User capacity
                            ~ 500 Fingerprint capacity
                            ~ 50,000 transaction logs capacity
                            ~ USB flash disk download
                            ~ Dimension: 106x60x42mm
                            ~ 18 MONTHS WARRANTY`,
              },
    TX628:    {image: path.join(__dirname, "assets/img/device/tx628.png"), 
               undiscounted: "14,900.00",
               price: 8900,
               description:`~ 3,200 User capacity
                            ~ 3,200 Fingerprint capacity
                            ~120,000 transaction logs capacity
                            ~ network connectivity
                            ~ USB flash disk download
                            ~ built-in scheduler bell
                            ~ Audio Visual Indication and rejection of Fingers
                            ~ 180 x 135 x 37mm dimension (W*H*T)
                            ~ 36 MONTHS WARRANTY`,
              },
    SC700:    {image: path.join(__dirname, "assets/img/device/sc700.png"),
               undiscounted: "10,800.00",
               price: 8800,
               description:`~ 30,000 Card Capacity
                            ~ 100,000 Logs Capacity
                            ~ Network Connectivity/ USB Host
                            ~ Support Magnetic Contact
                            ~ 125khz RFID
                            ~ Dimension: 106x104x36mm
                            ~ 18 MONTHS WARRANTY`
              },
    T8:       {image: path.join(__dirname, "assets/img/device/t8.png"),
               undiscounted: "15,900.00",
               price: 11200,
               description:`~ 3000 Fingerprint capacity
                            ~ 3000 card capacity
                            ~ 100,000 transaction logs capacity
                            ~ network connectivity
                            ~ USB flash disk download
                            ~ built-in scheduler bell
                            ~ Audio Visual Indication and rejection of Fingers
                            ~ Dimension: 108x140x30mm
                            ~ with ADMS
                            36 MONTHS WARRANTY`
              },
    FA1000:   {image: path.join(__dirname, "assets/img/device/fa1000.png"),
               undiscounted: "12,500.00",
               price: 9200,
               description:`~  4.3 inch Touch Screen  
                            ~ 1,000 User Capacity
                            ~ 500 Face Capacity
                            ~ 1,000 Card Capacity
                            ~ 150,000 Transaction Logs Capacity
                            ~ TCP/IP, WIFI
                            ~Standard Function: ADMS, Work Code, Self    Service Query, Automatic Status, T9 Input, Camera, 9 Digit User
                            ~ 18 MONTHS WARRANTY`
              },
    BK100:    {image: path.join(__dirname, "assets/img/device/bk100.png"), 
               undiscounted: "16,500.00",
               price: 9000,
               description:`~ 1000 User Capacity
                            ~ 800 Face Capacity
                            ~ 3,000 fingerprint templates capacity
                            ~ 250,000 transaction logs capacity
                            ~ network//USB flash disk   download
                            ~ Standard Function DSLT, Scheduled bell, Self Service Query, Automatic status switch, Photo Id
                            ~ Dimension: 161x93x152
                            ~ 36 MONTHS WARRANTY`
              },
    FA110:    {image: path.join(__dirname, "assets/img/device/fa110.png"),
               undiscounted: "17,900.00",
               price: 9700,
               description:`~ 2.8 Inch TFT Screen
                            ~ 500 Face Capacity
                            ~ 500 fingerprint templates capacity
                            ~ 50,000 transaction logs capacity
                            ~ network//USB Host/ WIFI
                            ~ Standard Function DSLT, Scheduled bell, Self Service Query, Automatic status switch, Photo Id 
                            Adms
                            ~ Dimension: 161x93x152
                            ~ 36 MONTHS WARRANTY`
              },
    F22:      {image: path.join(__dirname, "assets/img/device/f22.png"),
               undiscounted: "19,800.00",
               price: 13900,
               description:`~ 3000 fingerprint templates
                            ~ 5,000 card capacity
                            ~ 50,000 transaction logs capacity
                            ~ built-in EM card reader
                            ~ network/USB flash disk download/ WI-FI
                            ~ support multiple Time Zone
                            ~ support magnetic contact (
                            ~ Standard Function: 9 digital ID, Automatic Status Switch, Anti-Passback, Scheduler Bell
                            ~125khz RFID
                            ~ Dimension 78x158.5x41m
                            36 MONTHS WARRANTY `
              },
    SF200:    {image: path.join(__dirname, "assets/img/device/sf200.png"), 
               undiscounted: "17,500.00",
               price: 15700,
               description:`~ 2000 Fingerprint Templates Capacity
                            ~ 5,000 Card Capacity
                            ~ 100,000 Transaction logs Capacity
                            ~ Network Connectivity/ Wi-Fi 
                            ~ Support Multiple Time Zone
                            ~ support magnetic contact
                            ~ With Power Adapter
                            36 Months Warranty`
              },
    IFACE3:   {image: path.join(__dirname, "assets/img/device/iface3.png"),
               undiscounted: "19,400.00",
               price: 14200,
               description:`~ 1,500 Face Capacity
                            ~ 4,000 fingerprint templates capacity
                            ~ 5,000 Card Capacity
                            ~ 100,000 transaction logs capacity
                            ~ network//USB flash disk   download
                            ~ Automatic Switch Status
                            ~ power adapter
                            ~ 36 MONTHS WARRANTY`
              },
    MB460:    {image: path.join(__dirname, "assets/img/device/mb460.png"), 
               undiscounted: "18,500.00",
               price: 14800,
               description:`~ 1,500 Face Capacity
                            ~ 2,000 fingerprint templates capacity
                            ~ 5000 Card Capacity
                            ~ 100,000 transaction logs capacity
                            ~ network//USB flash disk   download
                            ~ Automatic Switch Status
                            ~ Dimension: 167x148x32mm
                            ~ with ADMS
                            ~ 36 MONTHS WARRANTY`
              },
    FA210:    {image: path.join(__dirname, "assets/img/device/fa210.png"),
               undiscounted: "22,500.00",
               price: 14800,
               description:`~ 1,500 Face Capacity
                            ~ 2,000 fingerprint templates capacity
                            ~ 100,000 transaction logs capacity
                            ~ network//USB flash disk   download
                            ~ Automatic Switch Status
                            ~ optional wifi. Special order (17,000.00)
                            ~ 36 MONTHS WARRANTY`
              },
    XFACE100: {image: path.join(__dirname, "assets/img/device/xface100.png"),
               undiscounted: "22,500.00",
               price: 18900,
               description:`~ 1,500 Face Capacity
                            ~ 2,000 fingerprint templates capacity
                            ~ 100,000 transaction logs capacity
                            ~ network//USB flash disk   download
                            ~ WIFI
                            ~ ADMS
                            ~ Automatic Switch Status
                            ~ 36 MONTHS WARRANTY`
              },
    MB560VL:  {image: path.join(__dirname, "assets/img/device/mb560vl.png"), 
               undiscounted: "27,000.00",
               price: 21800,
               description:`~ 2.8” TFT Screen
                            ~ 1,500 face templates capacity
                            ~ 2,000 fingerprint templates capacity
                            ~ 100,000 transactions logs capacity
                            ~ network connectivity/ wifi connection
                            ~ USB flash disk download
                            ~ Can identify with Face Mask
                            ~ 36 MONTHS WARRANTY`
              },
    UFACE800: {image: path.join(__dirname, "assets/img/device/uface800.png"),
               undiscounted: "32,900.00",
               price: 22800,
               description:`~ Touchscreen display with heat-sensitive function keys
                            ~ 3000 face templates capacity
                            ~ 4,000 fingerprint templates capacity
                            ~ 100,000 transactions logs capacity
                            ~ network connectivity
                            ~ wifi(optional)
                            ~ USB flash disk download
                            ~ built-in scheduler bell
                            ~ Dimension: 194x165x86mm
                            ~ 36 MONTHS WARRANTY`
              },
  };

  // user override
  const userCost = typeof cost === "number" && Number.isFinite(cost) ? cost : null;
  const catalogItem = productValues[product] ?? productValues["LX50"]; // fallback
  const finalCost = userCost !== null
    ? quantity * userCost
    : quantity * catalogItem.price;
  const vat = finalCost * 0.12;
  const total = finalCost + vat;

  // Get image path for selected product
  const selectedImagePath = catalogItem.image;
  // Get description from selected product
  const deviceDescription = productValues[product].description
    // Get undiscounted price from selected product
  const undiscountedPrice = productValues[product].undiscounted

  // Function to format number as Philippine Peso
  const formatPeso = (value) =>
    new Intl.NumberFormat("en-PH", {
      style: "currency",
      currency: "PHP",
      minimumFractionDigits: 2,
    }).format(value);

const doc = new Document({
  sections: [
    {
      properties: {
          page: {
              margin: {
                  top: 0, // Example: 0.5 inch (720 twips)
                  right: 720,
                  bottom: 720,
                  left: 720,
                  header: 288,
              },
              orientation: PageOrientation.PORTRAIT,
          },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new ImageRun({
                  data: fs.readFileSync(path.join(__dirname, "assets/img/header-jk-final.png")),
                  transformation: {
                    width: 698,   // adjust as needed
                    height: 178,  // adjust as needed
                  },
                }),
              ],
            }),
          ],
        }),
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: fs.readFileSync(path.join(__dirname, "assets/img/flex-row.png")),
              transformation: {
                width: 698, 
                height: 98,  
              },
            }),
          ],
        }),

        /* ----- Company Details Table ----- */
        new Paragraph({ text: "" }), // spacer
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },  // ~6.5" usable width
          layout: "fixed",                             // MUST be fixed
          rows: [
            new TableRow({
              children: [
                styledCell("Date", 10, true),                      
                styledCell(new Date().toLocaleDateString(), 45, true), 
                styledCell("Contact Person", 15, true),            
                styledCell(person || "-", 30, true),                 
              ],
            }),
            new TableRow({
              children: [
                styledCell("Client", 10, true),
                styledCell(client || "-", 45, true),
                styledCell("Contact Number", 15, true),
                styledCell(number || "-", 30, true),
              ],
            }),
            new TableRow({
              children: [
                styledCell("Address", 10, true),
                styledCell(address || "-", 45, true),
                styledCell("Email Address", 15, true),
                styledCell(email || "-", 30, true),
              ],
            }),
          ],
        }),

        /* ----- Thank you for your interest text ----- */
        new Paragraph({ text: "" }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Thank you for your interest in our products; we will assist you with selecting the best systems & solutions that would fit your requirements. ",
              font: "Century Gothic",
              size: 18,
            }),
          ],
         }),
        new Paragraph({ text: "" }),
        /* ----- Biometrics Table Here ----- */
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: "fixed",
          rows: [
            new TableRow({
              height: { value: 500, rule: HeightRule.EXACT },
              children: [
                styledCellWhite("Model", 20, false),
                styledCellWhite("Item Description", 40, false),
                styledCellWhite("Qty", 5, false),
                styledCellWhite("Unit", 9, false),
                styledCellWhite("Unit price", 13, false),
                styledCellWhite("Promo Amount", 13, true),
              ],
            }),
            new TableRow({
              children: [
                styledCellImage(selectedImagePath, 20, product, true),
                styledCellDescription(deviceDescription, 40, false, AlignmentType.LEFT),
                styledCell(String(quantity), 5, false, AlignmentType.CENTER),
                styledCell("pc", 9, false, AlignmentType.CENTER),
                styledCell(String(undiscountedPrice), 13, false, AlignmentType.CENTER),
                styledCell(String(formatPeso(finalCost)), 13, true, AlignmentType.CENTER),           
              ],
            }),
            new TableRow({
              children: [
                styledCell("Software", 20, false, AlignmentType.CENTER),
                styledCell("ZKTeco Attendance Management", 40, false, AlignmentType.LEFT),
                styledCell(String(quantity), 5, false, AlignmentType.CENTER),
                styledCell("License", 9, false, AlignmentType.CENTER),
                styledCell("Free", 13, false, AlignmentType.CENTER),
                styledCell("Free", 13, true, AlignmentType.CENTER),           
              ],
            }),
            new TableRow({
              children: [
                styledCell("", 20, false),
                styledCell("16GB USB FLASH DISK DRIVE", 40, false, AlignmentType.LEFT),
                styledCell(String(quantity), 5, false, AlignmentType.CENTER),
                styledCell("pc", 9, false, AlignmentType.CENTER),
                styledCell("Free", 13, false, AlignmentType.CENTER),
                styledCell("Free", 13, true, AlignmentType.CENTER),           
              ],
            }),
          ],
        }),

        /* ----- Prices ----- */

        new Paragraph({ text: "" }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({
              text: `EQUIPMENT COST = ${formatPeso(finalCost)}`,
              font: "Century Gothic",
              size: 18, //9pt
              bold: true,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({
              text: `PLUS 12% VAT = ${formatPeso(vat)}`,
              font: "Century Gothic",
              size: 18,
              bold: true,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({
              text: `TOTAL EQUIPMENT PRICE = ${formatPeso(total)}`,
              font: "Century Gothic",
              size: 18,
              bold: true,
              underline: true,
              highlight: "yellow",
            }),
          ],
        }),
      ],
    },
  ],
});

  // Helper to create styled table cell
  function styledCell(text, widthPercent, boldBool, alignmentType) {
    return new TableCell({
      width: { size: widthPercent, type: WidthType.PERCENTAGE },
      children: [
        new Paragraph({
          alignment: alignmentType,
          children: [
            new TextRun({
              text,
              font: "Century Gothic",
              size: 18,
              bold: boldBool,
            }),
          ],
        }),
      ],
    });
  }

  // Helper to create styled table cell White text
  function styledCellWhite(text, widthPercent) {
    return new TableCell({
      width: { size: widthPercent, type: WidthType.PERCENTAGE },
      shading: {
        type: "clear",
        fill: "0070C0", // Blue background
        color: "auto",
      },
      verticalAlign: "center",
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text,
              font: "Century Gothic",
              color: "FFFFFF",
              size: 18,
              bold: false,
            }),
          ],
        }),
      ],
    });
  }

function styledCellImage(imagePath, widthPercent, caption = "", boldBool) {
  return new TableCell({
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    verticalAlign: "center",
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: fs.readFileSync(imagePath),
            transformation: {
              width: 100,   // adjust thumbnail size
              height: 100,
            },
          }),
        ],
      }),
      ...(caption
        ? [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: caption,
                  font: "Century Gothic",
                  size: 18,
                  bold: boldBool,
                  color: "000000", // normal black text
                }),
              ],
            }),
          ]
        : []),
    ],
  });
}

function styledCellDescription(text, widthPercent, boldBool, alignment) {
  // Split text into lines (handles both \n and template literal breaks)
  const lines = String(text ?? "").split(/\r?\n/);

  return new TableCell({
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    verticalAlign: "top", // descriptions usually align to the top
    children: lines.map(line =>
      new Paragraph({
        alignment,
        children: [
          new TextRun({
            text: line.trim(), // trim spaces
            font: "Century Gothic",
            size: 18,
            bold: boldBool,
          }),
        ],
      })
    ),
  });
}

  const buffer = await Packer.toBuffer(doc);

  const filePath = path.join(app.getPath("desktop"), "Quotation.docx");
  fs.writeFileSync(filePath, buffer);

  event.sender.send("docx-done", filePath);

});