const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const {
  Document,
  Section,
  Packer,
  Paragraph,
  TextRun,
  Tab,
  AlignmentType,
  VerticalAlign,
  ImageRun,
  Header,
  Table,
  TableRow,
  TableCell,
  WidthType,
  PageMargin,
  PageOrientation,
  HeightRule,
} = require("docx");


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
    agent,          // 👈 Sales Agent
    client = "-",
    address = "-",
    person = "-",
    number = "-",
    email = "-",
    product,
    quantity,       // 👈 No. of Biometrics
    vatType,        // 👈 "VAT Inclusive/Exclusive"
    cost,           // 👈 "Custom Amount for i.e. Dealers Price"
  } = details;

   // Map product values
  const productValues = {
    LX50:     {image: path.join(__dirname, "assets/img/device/lx50.png"),
               undiscounted: "₱10,900.00",
               price: 5700,
               warranty: 18,
               description:`~ 500 User capacity
                            ~ 500 Fingerprint capacity
                            ~ 50,000 transaction logs capacity
                            ~ USB flash disk download
                            ~ Dimension: 106x60x42mm
                            ~ 18 MONTHS WARRANTY`,
              },
    TX628:    {image: path.join(__dirname, "assets/img/device/tx628.png"), 
               undiscounted: "₱14,900.00",
               price: 8900,
               warranty: 36,
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
               undiscounted: "₱10,800.00",
               price: 8800,
               warranty: 18,
               description:`~ 30,000 Card Capacity
                            ~ 100,000 Logs Capacity
                            ~ Network Connectivity/ USB Host
                            ~ Support Magnetic Contact
                            ~ 125khz RFID
                            ~ Dimension: 106x104x36mm
                            ~ 18 MONTHS WARRANTY`
              },
    T8:       {image: path.join(__dirname, "assets/img/device/t8.png"),
               undiscounted: "₱15,900.00",
               price: 11200,
               warranty: 36,
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
               undiscounted: "₱12,500.00",
               price: 9200,
               warranty: 18,
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
               undiscounted: "₱16,500.00",
               price: 9000,
               warranty: 36,
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
               undiscounted: "₱17,900.00",
               price: 9700,
               warranty: 36,
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
               undiscounted: "₱19,800.00",
               price: 13900,
               warranty: 36,
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
               undiscounted: "₱17,500.00",
               price: 15700,
               warranty: 36,
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
               undiscounted: "₱19,400.00",
               price: 14200,
               warranty: 36,
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
               undiscounted: "₱18,500.00",
               price: 14800,
               warranty: 36,
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
               undiscounted: "₱22,500.00",
               price: 14800,
               warranty: 36,
               description:`~ 1,500 Face Capacity
                            ~ 2,000 fingerprint templates capacity
                            ~ 100,000 transaction logs capacity
                            ~ network//USB flash disk   download
                            ~ Automatic Switch Status
                            ~ optional wifi. Special order (17,000.00)
                            ~ 36 MONTHS WARRANTY`
              },
    XFACE100: {image: path.join(__dirname, "assets/img/device/xface100.png"),
               undiscounted: "₱22,500.00",
               price: 18900,
               warranty: 36,
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
               undiscounted: "₱27,000.00",
               price: 21800,
               warranty: 36,
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
               undiscounted: "₱32,900.00",
               price: 22800,
               warranty: 36,
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
    TMD95E:   {image: path.join(__dirname, "assets/img/device/tmd95e.png"),
               undiscounted: "₱9,500.00",
               price: 9500,
               warranty: 12,
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
  // Get warranty from selected product
  const warranty = productValues[product].warranty || 12; // fallback 12 months

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
          titlePage: true,
      },
      headers: {
        first: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new ImageRun({
                  data: fs.readFileSync(path.join(__dirname, getAgentCard(agent))),
                  transformation: {
                    width: 698,   // adjust as needed
                    height: 168,  // adjust as needed
                  },
                }),
              ],
            }),
          ],
        }),
        default: new Header({children: []}),
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

        /* ----- Quote Reference Number ----- */
        new Paragraph({ text: "" }), // spacer
        new Paragraph({          
          children: [
            new TextRun({
              text: generateQuoteRef(agent) ,
              font: "Century Gothic",
              size: 18,
              bold: true,
            }),
          ],
        }),
        /* ----- Company Details Table ----- */
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
                styledCellPromoAmount(`Promo
                                      Amount`, 13, true),
              ],
            }),
            new TableRow({
              children: [
                styledCellImage(selectedImagePath, 20, product, true),
                styledCellDescription(deviceDescription, 40, false, AlignmentType.LEFT),
                styledCellCenter(String(quantity), 5, false, AlignmentType.CENTER),
                styledCellCenter("pc", 9, false, AlignmentType.CENTER),
                styledCellCenter(String(undiscountedPrice), 13, false, AlignmentType.CENTER),
                styledCellCenter(String(formatPeso(finalCost)), 13, true, AlignmentType.CENTER),           
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
        new Paragraph({ text: "" }),
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
        new Paragraph({ text: "" }),

        /* ----- Conforme Signature ----- */




        /* ----- Optional Accessories ----- */

        new Paragraph({          
          children: [
            new TextRun({
              text: "OPTIONAL ACCESSORIES",
              font: "Century Gothic",
              size: 18,
              bold: true,
            }),
          ],
        }),

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
                styledCellPromoAmount(`Promo
                                      Amount`, 13, true),
              ],
            }),
            new TableRow({
              children: [
                styledCellImage(productValues.TMD95E.image, 20, "TMD95E", true),
                styledCellDescription(`Model: TDM95E
                                      Communication: USB 1.1
                                      USB: Type-C
                                      Temperature Detection: Distance 3 - 5cm
                                      Temperature: Unit °C or °F
                                      Temperature: Measurement Range 32.0°C - 42.9°C / 89.6°F - 109.22°F
                                      Temperature Measurement Deviation: ±0.3°C / ±0.54°F
                                      Digital Display: Tube 4
                                      Operating Voltage: 5V
                                      Operating Temperature: 15°C - 35°C / 59°F - 95°F
                                      Operating Humidity: 10% - 85%
                                      Dimensions: 88 * 88 * 54.63 (mm)
                                      Weight of the Device: 0.17kg
                                      Weight of the Device with Packaging: 0.29kg
                                      `, 40, false, AlignmentType.LEFT),
                styledCellCenter(String(quantity), 5, false, AlignmentType.CENTER),
                styledCellCenter("pc", 9, false, AlignmentType.CENTER),
                styledCellCenter(String(formatPeso(9500)), 13, false, AlignmentType.CENTER),
                styledCellCenter(String(formatPeso(9500)), 13, true, AlignmentType.CENTER),           
              ],
            }),
          ],
        }),
        new Paragraph({ text: "" }),


        /* ----- Terms & Conditions ----- */

        new Table({
          width: { size: 96, type: WidthType.PERCENTAGE },
          layout: "fixed",
          rows: [
            new TableRow({
              height: { value: 400, rule: HeightRule.EXACT },
              children: [
                new TableCell({
                  children: [new Paragraph("")], // empty cell
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: "TERMS & CONDITIONS:",
                          font: "Century Gothic",
                          size: 18,
                          bold: true,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: "1.) Prices quoted above are VAT Inclusive. Email or fax certification if your company is vat exempt and zero rated for billing preparation.",
                          font: "Century Gothic",
                          size: 18,
                          bold: false,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: "2.) Prices are subject to change without prior notice. Validity for this quotation is 15 days from the date stated above.",
                          font: "Century Gothic",
                          size: 18,
                          bold: false,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: "3.) Payment terms is Fifty Percent (50%) upon P.O. or signing of this CONFORME.  Remaining balance shall be paid upon receive of items or after the installation. ",
                          font: "Century Gothic",
                          size: 18,
                          bold: false,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: "4. Payment will be accepted in CASH, COD, and Dated Check or thru Bank Transfer payable to ",
                          font: "Century Gothic",
                          size: 18,
                          bold: false,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: "TECHFINITY SECURITY DEVICE TRADING.",
                          font: "Century Gothic",
                          size: 18,
                          bold: true,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: "5. FREE DELIVERY for purchases above Php10, 000 within Metro Manila, otherwise additional Php 500 delivery fee depending on location.",
                          font: "Century Gothic",
                          size: 18,
                          bold: false,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: "6. Cancelled orders are subject to a cancellation charge of Fifty Percent (50%).",
                          font: "Century Gothic",
                          size: 18,
                          bold: false,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: `7. Up to ${productValues[product].warranty} months limited warranty in service and parts will be given for main equipment from date of purchase/delivery/installation. Accessories such as power supply, adaptor, magnetic lock, exit button have six (6) months warranty. The warranty covers the parts cause of factory defect not including upgrades and relocation. Unauthorized repair will void its warranty. Warranty claims is strictly carried in basis, client must send the item to our office for repair. For those with installation, we will do the onsite checking and troubleshooting for free within metro manila, for outside metro manila client will pay for the mobilization/demobilization cost.`,
                          font: "Century Gothic",
                          size: 18,
                          bold: false,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                      },
                      children: [
                        new TextRun({
                          text: "8. Should client will require service unit while defective device is under repair; client must pay a service unit fee but depends on the availability of the service unit.",
                          font: "Century Gothic",
                          size: 18,
                          bold: false,
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: {
                        line: 238,   // 1.08 * 216 twips
                        lineRule: "exactly",
                        after: 238,
                        before: 0,
                      },
                      children: [
                        new TextRun({
                          text: "9. After sales support is from Monday – Friday 8:30 – 5:30 pm",
                          font: "Century Gothic",
                          size: 18,
                          bold: false,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({
              text: "Best Regards,",
              font: "Century Gothic",
              size: 18,
              bold: true,
            }),
          ],
        }),
        new Paragraph({ text: "" }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({
              text: getAgentName(agent),
              font: "Century Gothic",
              size: 18,
              bold: true,
              underline: true,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({
              text: "Sales Account Officer",
              font: "Century Gothic",
              size: 18,
              bold: false,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({
              text: getAgentNumber(agent),
              font: "Century Gothic",
              size: 18,
              bold: false,
            }),
          ],
        }),
      ],
    },
  ],
});







// Helper to generate Quote Reference based on agent code
function getAgentCard(agent) {
  if (!agent) return 'assets/img/header-jk-final.png';
  
  if (agent == "CLEO") {
    return `assets/img/header-jk-final.png`;
  } 
  else if (agent == "JHEL") {
    return `assets/img/header-jhel-final.png`;
  }
  else if (agent == "JK") {
    return `assets/img/header-jk-final.png`;
  }
  else if (agent == "SHAE") {
    return `assets/img/header-jk-final.png`;
  }
  else {
    return "assets/img/header-jk-final.png";
  }
}

// Helper to generate Quote Reference based on agent code
function generateQuoteRef(agent) {
  const quotePrefix = "Quote Ref No:";
  const quoteYear = new Date().getFullYear();
  let agentCode = "debug";
  if (agent) {
    const cleanAgent = agent.toUpperCase();
    agentCode = cleanAgent.length <= 2 ? cleanAgent : cleanAgent[0];
  }
  const quoteNum = "X" //placeholder for now
  return `${quotePrefix} ${quoteYear}-${agentCode}-${quoteNum}`;
}

// Helper to get FullName of agent
function getAgentName(agent) {
  if (!agent) return "TECHFINITY";

  if (agent == "CLEO") {
    return `LEO DURA`;
  } 
  else if (agent == "JHEL") {
    return `JHEL VILLAVECENCIO`;
  }
  else if (agent == "JK") {
    return `JOHN KARL NOLASCO`;
  }
  else if (agent == "SHAE") {
    return `SHAENA FALLE`;
  }
  else {
    return "TECHFINITY";
  }
}

// Helper to get ContactNumber of agent
function getAgentNumber(agent) {
  if (!agent) return "09463360774";
  
  if (agent == "CLEO") {
    return `09100255412`;
  } 
  else if (agent == "JHEL") {
    return `09460378085`;
  }
  else if (agent == "JK") {
    return `09484263778`;
  }
  else if (agent == "SHAE") {
    return `09070456737`;
  }
  else {
    return "09463360774";
  }
}

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

// Helper to create styled table cell Centered Vertically
function styledCellCenter(text, widthPercent, boldBool, alignmentType) {
  return new TableCell({
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
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

function styledCellPromoAmount(text, widthPercent, boldBool) {
  // Split text into lines (handles both \n and template literal breaks)
  const lines = String(text ?? "").split(/\r?\n/);

  return new TableCell({
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    shading: {
      type: "clear",
      fill: "0070C0", // Blue background
      color: "auto",
    },
    verticalAlign: "top", // descriptions usually align to the top
    children: lines.map(line =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: line.trim(), // trim spaces
            font: "Century Gothic",
            color: "FFFFFF",
            size: 18,
            bold: boldBool,
          }),
        ],
      })
    ),
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