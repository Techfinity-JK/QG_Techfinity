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
ipcMain.on("generate-docx", async (event, cost) => {
  const vat = cost * 0.12;
  const total = cost + vat;

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
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: "fixed",
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 10, type: WidthType.PERCENTAGE },
                  children: [new Paragraph("Date")],
                }),
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  children: [new Paragraph(new Date().toLocaleDateString())],
                }),
                new TableCell({
                  width: { size: 15, type: WidthType.PERCENTAGE },
                  children: [new Paragraph("Contact Person")],
                }),
                new TableCell({
                  width: { size: 35, type: WidthType.PERCENTAGE },
                  children: [new Paragraph("John Doe")],
                }),
              ],
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

});