const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const { Document, Section, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Header, Table, TableRow, TableCell, WidthType, PageMargin, PageOrientation} = require("docx");


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
    cost,
    product = "ProductA", // Default product fallback
  } = details;

  const vat = cost * 0.12;
  const total = cost + vat;

  // Map product keys to image file paths
  const productImages = {
    LX50: path.join(__dirname, "assets/img/device/lx50.png"),
    TX628: path.join(__dirname, "assets/img/device/tx628.png"),
    SC700: path.join(__dirname, "assets/img/device/sc700.png"),
    T8: path.join(__dirname, "assets/img/device/t8.png"),
    FA1000: path.join(__dirname, "assets/img/device/fa1000.png"),
    BK100: path.join(__dirname, "assets/img/device/bk100.png"),
    FA110: path.join(__dirname, "assets/img/device/fa110.png"),
    F22: path.join(__dirname, "assets/img/device/f22.png"),
    SF200: path.join(__dirname, "assets/img/device/sf200.png"),
    IFACE3: path.join(__dirname, "assets/img/device/iface3.png"),
    MB460: path.join(__dirname, "assets/img/device/mb460.png"),
    FA210: path.join(__dirname, "assets/img/device/fa210.png"),
    XFACE100: path.join(__dirname, "assets/img/device/xface100.png"),
    MB560VL: path.join(__dirname, "assets/img/device/mb560vl.png"),
    UFACE800: path.join(__dirname, "assets/img/device/uface800.png"),
  };

  // Pick the correct image (fallback to ProductA if missing)
  const selectedImagePath = productImages[product] || productImages["ProductA"];

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
                width: 698,   // adjust as needed
                height: 98,  // adjust as needed
              },
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
                styledCell("Date", 10),                      
                styledCell(new Date().toLocaleDateString(), 45), 
                styledCell("Contact Person", 15),            
                styledCell(person || "-", 30),                 
              ],
            }),
            new TableRow({
              children: [
                styledCell("Client", 10),
                styledCell(client || "-", 45),
                styledCell("Contact Number", 15),
                styledCell(number || "-", 30),
              ],
            }),
            new TableRow({
              children: [
                styledCell("Address", 10),
                styledCell(address || "-", 45),
                styledCell("Email Address", 15),
                styledCell(email || "-", 30),
              ],
            }),
          ],
        }),

        /* ----- Product Image Here ----- */

        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: fs.readFileSync(selectedImagePath),
              transformation: {
                width: 400,   // adjust as needed
                height: 300,  // adjust as needed
              },
            }),
          ],
        }),

        /* ----- Prices ----- */

        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({
              text: `EQUIPMENT COST = ${cost.toFixed(2)}`,
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
              text: `PLUS 12% VAT = ${vat.toFixed(2)}`,
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
              text: `TOTAL EQUIPMENT PRICE = ${total.toFixed(2)}`,
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

  const buffer = await Packer.toBuffer(doc);

  const filePath = path.join(app.getPath("desktop"), "Quotation.docx");
  fs.writeFileSync(filePath, buffer);

  event.sender.send("docx-done", filePath);

  function styledCell(text, widthPercent) {
    return new TableCell({
      width: { size: widthPercent, type: WidthType.PERCENTAGE },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text,
              font: "Century Gothic",
              size: 18,
              bold: true,
            }),
          ],
        }),
      ],
    });
  }

});