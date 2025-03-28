require("dotenv").config();
const express = require("express");
const bodyParser = require("body-parser");
const { google } = require("googleapis");
const mongoose = require("mongoose");
const multer = require("multer");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const axios = require("axios");
const app = express();
const PORT = process.env.PORT || 3001;

app.use(
  cors({
    origin: "http://localhost:3000",
    methods: ["GET", "POST", "PUT", "DELETE"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: true,
    optionsSuccessStatus: 200,
  })
);

app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, "public")));

const storage = multer.memoryStorage();
const upload = multer({ storage });

const headerValues = [
  "Voucher No.",
  "Date",
  "Filter",
  "Pay to",
  "Account Head",
  "Towards",
  "The Sum",
  "Amount Rs.",
  "Checked By",
  "Approved By",
  "Receiver Signature",
  "PDF Link",
];

// MongoDB Connection
const mongoURI = process.env.MONGO_URI || `mongodb://${mongoUsername}:${mongoPassword}@localhost:27017/voucherDB?authSource=admin`;

mongoose.connect(mongoURI)
  .then(() => console.log("Connected to MongoDB with authentication"))
  .catch(err => console.error("MongoDB connection error:", err));

// Voucher Schema
const voucherSchema = new mongoose.Schema({
  email: { type: String, required: true },
  company: { type: String, required: true },
  voucherNo: { type: Number, required: true },
  date: String,
  payTo: String,
  accountHead: String,
  account: String,
  amount: String,
  amountRs: String,
  checkedBy: String,
  approvedBy: String,
  receiverSignature: String,
  pdfLink: String,
  spreadsheetId: String,
  folderId: String,
  pdfFileId: String,
});

const Voucher = mongoose.model("Voucher", voucherSchema);

// Middleware to authenticate Google token
const authenticateGoogle = (req) => {
  if (!req.headers.authorization) throw new Error("No authorization token provided");
  const token = req.headers.authorization.split("Bearer ")[1];
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: token });
  return auth;
};

// Function to create a new spreadsheet
async function createSpreadsheet(companyName, auth) {
  const sheets = google.sheets({ version: "v4", auth });
  try {
    console.log(`Creating new spreadsheet for ${companyName}`);
    const spreadsheet = await sheets.spreadsheets.create({
      resource: {
        properties: {
          title: `${companyName} Vouchers`,
        },
        sheets: [
          {
            properties: {
              title: companyName,
              gridProperties: {
                rowCount: 1000,
                columnCount: 14,
              },
            },
          },
        ],
      },
      fields: "spreadsheetId",
    });
    const spreadsheetId = spreadsheet.data.spreadsheetId;
    console.log(`Created spreadsheet ID: ${spreadsheetId}`);

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${companyName}!A1:O1`,
      valueInputOption: "RAW",
      requestBody: {
        values: [headerValues],
      },
    });
    console.log(`Headers set for ${spreadsheetId}`);
    return spreadsheetId;
  } catch (error) {
    console.error(`Error creating spreadsheet for ${companyName}:`, error.message);
    throw error;
  }
}

// Function to create a new Drive folder
async function createDriveFolder(companyName, auth) {
  const drive = google.drive({ version: "v3", auth });
  try {
    console.log(`Creating new Drive folder for ${companyName}`);
    const folder = await drive.files.create({
      resource: {
        name: `${companyName} Vouchers`,
        mimeType: "application/vnd.google-apps.folder",
      },
      fields: "id",
    });
    const folderId = folder.data.id;
    console.log(`Created Drive folder ID: ${folderId}`);
    return folderId;
  } catch (error) {
    console.error(`Error creating Drive folder for ${companyName}:`, error.message);
    throw error;
  }
}

// Get or create resources per user and generate unique voucher number
async function getUserResources(email, companyName, auth) {
  let spreadsheetId, folderId;

  const existingVoucher = await Voucher.findOne({ email, company: companyName });
  if (existingVoucher) {
    spreadsheetId = existingVoucher.spreadsheetId;
    folderId = existingVoucher.folderId;
  } else {
    spreadsheetId = await createSpreadsheet(companyName, auth);
    folderId = await createDriveFolder(companyName, auth);
  }

  const highestVoucher = await Voucher.findOne({ email, company: companyName })
    .sort({ voucherNo: -1 })
    .select("voucherNo");
  const voucherNumber = highestVoucher ? highestVoucher.voucherNo + 1 : 1;

  return { spreadsheetId, folderId, voucherNumber };
}

setInterval(() => {
  axios
    .get(`http://localhost:${PORT}/ping`)
    .then((response) => {
      console.log("Pinged server to keep it warm.");
    })
    .catch((error) => {
      console.error("Error pinging the server:", error.message);
    });
}, 30000);

app.get("/ping", (req, res) => {
  res.status(200).send({ message: "Server is active" });
});

app.get("/get-voucher-no", async (req, res) => {
  const filter = req.query.filter;
  if (!filter || !["Contentstack", "Surfboard", "RawEngineering"].includes(filter)) {
    return res.status(400).send({ error: "Invalid filter option" });
  }

  try {
    const auth = authenticateGoogle(req);
    const userInfo = await google.oauth2({ version: "v2", auth }).userinfo.get();
    const email = userInfo.data.email;
    const { voucherNumber } = await getUserResources(email, filter, auth);
    res.send({ voucherNo: voucherNumber });
  } catch (error) {
    console.error("Error in get-voucher-no:", error.message);
    res.status(500).send({ error: "Failed to generate voucher number: " + error.message });
  }
});

// Get all vouchers for a user
app.get("/vouchers", async (req, res) => {
  try {
    const auth = authenticateGoogle(req);
    const userInfo = await google.oauth2({ version: "v2", auth }).userinfo.get();
    const email = userInfo.data.email;

    const { company, date, sort } = req.query;
    let query = { email };
    if (company) query.company = company;
    if (date) query.date = date;

    let sortOption = {};
    if (sort === "lowToHigh") sortOption.amount = 1;
    else if (sort === "highToLow") sortOption.amount = -1;
    else sortOption.voucherNo = 1;

    const vouchers = await Voucher.find(query).sort(sortOption);
    res.status(200).send(vouchers);
  } catch (error) {
    console.error("Error retrieving vouchers:", error.message);
    res.status(500).send({ error: "Failed to retrieve vouchers: " + error.message });
  }
});

// Edit voucher endpoint
app.put("/edit-voucher/:id", upload.none(), async (req, res) => {
  try {
    const voucherId = req.params.id;
    const voucherData = req.body;
    const filterOption = voucherData.filter;

    if (!["Contentstack", "Surfboard", "RawEngineering"].includes(filterOption)) {
      return res.status(400).send({ error: "Invalid filter option" });
    }

    const auth = authenticateGoogle(req);
    const userInfo = await google.oauth2({ version: "v2", auth }).userinfo.get();
    const email = userInfo.data.email;

    const existingVoucher = await Voucher.findOne({ _id: voucherId, email });
    if (!existingVoucher) {
      return res.status(404).send({ error: "Voucher not found" });
    }

    const { spreadsheetId, folderId } = existingVoucher;
    const voucherNo = existingVoucher.voucherNo;
    const sheetTitle = filterOption;
    const sheetURL = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    const pdfFileName = `${filterOption}_${voucherNo}.pdf`;
    const pdfFilePath = path.join(__dirname, pdfFileName);
    const doc = new PDFDocument({ margin: 30 });
    const pdfStream = fs.createWriteStream(pdfFilePath);
    doc.pipe(pdfStream);

    const underlineYPosition = 35;

    doc.fontSize(12).text("Date:", 400, 20);
    doc.fontSize(12).text(voucherData.date, 440, 20);
    doc.moveTo(440, underlineYPosition).lineTo(550, underlineYPosition).stroke();

    doc.fontSize(12).text("Voucher No:", 400, 40);
    doc.fontSize(12).text(voucherNo, 470, 40);
    doc.moveTo(440, underlineYPosition + 20).lineTo(550, underlineYPosition + 20).stroke();

    const filterLogoMap = {
      Contentstack: path.join(__dirname, "public", "contentstack.png"),
      Surfboard: path.join(__dirname, "public", "surfboard.png"),
      RawEngineering: path.join(__dirname, "public", "raw.png"),
    };
    const filterLogo = filterLogoMap[voucherData.filter];
    if (fs.existsSync(filterLogo)) {
      doc.image(filterLogo, 30, 30, { width: 100 });
    }

    doc.moveDown(3);

    const drawLineAndText = (label, value, yPosition) => {
      doc.fontSize(12).text(label, 30, yPosition);
      doc.moveTo(120, yPosition + 12).lineTo(550, yPosition + 12).stroke();
      doc.fontSize(12).text(value, 130, yPosition);
    };

    drawLineAndText("Pay to:", voucherData.payTo, 160);
    drawLineAndText("Account Head:", voucherData.accountHead, 200);
    drawLineAndText("Towards:", voucherData.account, 240);

    doc.fontSize(12).text("Amount Rs.", 30, 280);
    doc.moveTo(120, 292).lineTo(550, 292).stroke();
    doc.fontSize(12).text(voucherData.amount, 130, 280);

    doc.fontSize(12).text("The Sum.", 30, 320);
    doc.moveTo(120, 332).lineTo(550, 332).stroke();
    doc.fontSize(12).text(voucherData.amountRs, 130, 320);

    const amountSectionY = 320;
    const gap = 65;
    const signatureSectionY = amountSectionY + gap;

    const drawSignatureLine = (label, xPosition, yPosition) => {
      doc.moveTo(xPosition, yPosition).lineTo(xPosition + 100, yPosition).stroke();
      doc.fontSize(12).text(label, xPosition, yPosition + 5);
    };

    drawSignatureLine("Checked By", voucherData.checkedBy, 50, signatureSectionY);
    drawSignatureLine("Approved By", voucherData.approvedBy, 250, signatureSectionY);
    drawSignatureLine("Receiver Signature", voucherData.receiverSignature, 450, signatureSectionY);

    doc.end();

    pdfStream.on("finish", async () => {
      try {
        const drive = google.drive({ version: "v3", auth });
        console.log(`Uploading updated PDF ${pdfFileName} to Drive folder ${folderId}`);

        if (existingVoucher.pdfFileId) {
          await drive.files.delete({ fileId: existingVoucher.pdfFileId });
          console.log(`Deleted old PDF: ${existingVoucher.pdfFileId}`);
        }

        const pdfFileMetadata = {
          name: pdfFileName,
          parents: [folderId],
        };
        const pdfMedia = {
          mimeType: "application/pdf",
          body: fs.createReadStream(pdfFilePath),
        };
        const pdfUploadResponse = await drive.files.create({
          resource: pdfFileMetadata,
          media: pdfMedia,
          fields: "id, webViewLink",
        });

        const pdfFileId = pdfUploadResponse.data.id;
        const pdfLink = pdfUploadResponse.data.webViewLink;
        console.log(`PDF uploaded: ${pdfFileId}, Link: ${pdfLink}`);

        const sheets = google.sheets({ version: "v4", auth });
        const values = [
          [
            voucherNo,
            voucherData.date,
            voucherData.filter,
            voucherData.payTo,
            voucherData.accountHead,
            voucherData.account,
            voucherData.amount,
            voucherData.amountRs,
            voucherData.checkedBy,
            voucherData.approvedBy,
            voucherData.receiverSignature,
            pdfLink,
          ],
        ];

        const sheetData = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${sheetTitle}!A2:O`,
        });
        const rows = sheetData.data.values || [];
        const rowIndex = rows.findIndex(row => row[0] == voucherNo);
        if (rowIndex === -1) {
          throw new Error("Voucher not found in sheet");
        }
        const rowRange = `${sheetTitle}!A${rowIndex + 2}:O${rowIndex + 2}`;

        console.log(`Updating data in sheet ${spreadsheetId} at ${rowRange}`);
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: rowRange,
          valueInputOption: "RAW",
          requestBody: {
            values: [values[0]],
          },
        });
        console.log(`Data updated in sheet ${spreadsheetId}`);

        await Voucher.updateOne(
          { _id: voucherId, email },
          {
            date: voucherData.date,
            payTo: voucherData.payTo,
            accountHead: voucherData.accountHead,
            account: voucherData.account,
            amount: voucherData.amount,
            amountRs: voucherData.amountRs,
            checkedBy: voucherData.checkedBy,
            approvedBy: voucherData.approvedBy,
            receiverSignature: voucherData.receiverSignature,
            pdfLink,
            pdfFileId,
          }
        );
        console.log(`Voucher updated in MongoDB: ${voucherNo}`);

        fs.unlinkSync(pdfFilePath);

        res.status(200).send({
          message: "Voucher updated successfully!",
          sheetURL: sheetURL,
          pdfFileId: pdfFileId,
        });
      } catch (error) {
        console.error("Error updating voucher:", error.message);
        res.status(500).send({ error: "Failed to update voucher: " + error.message });
      }
    });
  } catch (error) {
    console.error("Error in /edit-voucher endpoint:", error.message);
    res.status(500).send({ error: "Failed to edit voucher: " + error.message });
  }
});

// Delete voucher endpoint
app.delete("/vouchers/:voucherNo", async (req, res) => {
  try {
    const voucherNo = req.params.voucherNo;

    const auth = authenticateGoogle(req);
    const userInfo = await google.oauth2({ version: "v2", auth }).userinfo.get();
    const email = userInfo.data.email;

    const existingVoucher = await Voucher.findOne({ voucherNo: Number(voucherNo), email });
    if (!existingVoucher) {
      return res.status(404).send({ error: "Voucher not found" });
    }

    const { spreadsheetId, folderId, pdfFileId } = existingVoucher;
    const sheetTitle = existingVoucher.company;

    const drive = google.drive({ version: "v3", auth });
    if (pdfFileId) {
      await drive.files.delete({ fileId: pdfFileId });
      console.log(`Deleted PDF: ${pdfFileId}`);
    }

    const sheets = google.sheets({ version: "v4", auth });
    const sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetTitle}!A2:O`,
    });
    const rows = sheetData.data.values || [];
    const rowIndex = rows.findIndex(row => row[0] == voucherNo);
    if (rowIndex !== -1) {
      const rowRange = `${sheetTitle}!A${rowIndex + 2}:O${rowIndex + 2}`;
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: rowRange,
      });
      console.log(`Deleted row from sheet ${spreadsheetId} at ${rowRange}`);
    }

    await Voucher.deleteOne({ voucherNo: Number(voucherNo), email });
    console.log(`Deleted voucher from MongoDB: ${voucherNo}`);

    res.status(200).send({ message: "Voucher deleted successfully!" });
  } catch (error) {
    console.error("Error in /vouchers endpoint:", error.message);
    res.status(500).send({ error: "Failed to delete voucher: " + error.message });
  }
});

app.post("/submit", upload.none(), async (req, res) => {
  try {
    const voucherData = req.body;
    const filterOption = voucherData.filter;

    if (!["Contentstack", "Surfboard", "RawEngineering"].includes(filterOption)) {
      return res.status(400).send({ error: "Invalid filter option" });
    }

    const auth = authenticateGoogle(req);
    const userInfo = await google.oauth2({ version: "v2", auth }).userinfo.get();
    const email = userInfo.data.email;
    const { spreadsheetId, folderId, voucherNumber } = await getUserResources(email, filterOption, auth);
    const voucherNo = voucherData.voucherNo || voucherNumber;
    const sheetTitle = filterOption;
    const sheetURL = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    if (!voucherNo) {
      throw new Error("Voucher number is required");
    }

    const pdfFileName = `${filterOption}_${voucherNo}.pdf`;
    const pdfFilePath = path.join(__dirname, pdfFileName);
    const doc = new PDFDocument({ margin: 30 });
    const pdfStream = fs.createWriteStream(pdfFilePath);
    doc.pipe(pdfStream);

    const underlineYPosition = 35;

    doc.fontSize(12).text("Date:", 400, 20);
    doc.fontSize(12).text(voucherData.date, 440, 20);
    doc.moveTo(440, underlineYPosition).lineTo(550, underlineYPosition).stroke();

    doc.fontSize(12).text("Voucher No:", 400, 40);
    doc.fontSize(12).text(voucherNo, 470, 40);
    doc.moveTo(440, underlineYPosition + 20).lineTo(550, underlineYPosition + 20).stroke();

    const filterLogoMap = {
      Contentstack: path.join(__dirname, "public", "contentstack.png"),
      Surfboard: path.join(__dirname, "public", "surfboard.png"),
      RawEngineering: path.join(__dirname, "public", "raw.png"),
    };
    const filterLogo = filterLogoMap[voucherData.filter];
    if (fs.existsSync(filterLogo)) {
      doc.image(filterLogo, 30, 30, { width: 100 });
    }

    doc.moveDown(3);

    const drawLineAndText = (label, value, yPosition) => {
      doc.fontSize(12).text(label, 30, yPosition);
      doc.moveTo(120, yPosition + 12).lineTo(550, yPosition + 12).stroke();
      doc.fontSize(12).text(value, 130, yPosition);
    };

    drawLineAndText("Pay to:", voucherData.payTo, 160);
    drawLineAndText("Account Head:", voucherData.accountHead, 200);
    drawLineAndText("Towards:", voucherData.account, 240);

    doc.fontSize(12).text("Amount Rs.", 30, 280);
    doc.moveTo(120, 292).lineTo(550, 292).stroke();
    doc.fontSize(12).text(voucherData.amount, 130, 280);

    doc.fontSize(12).text("The Sum.", 30, 320);
    doc.moveTo(120, 332).lineTo(550, 332).stroke();
    doc.fontSize(12).text(voucherData.amountRs, 130, 320);

    const amountSectionY = 320;
    const gap = 65;
    const signatureSectionY = amountSectionY + gap;

    const drawSignatureLine = (label, xPosition, yPosition) => {
      doc.moveTo(xPosition, yPosition).lineTo(xPosition + 100, yPosition).stroke();
      doc.fontSize(12).text(label, xPosition, yPosition + 5);
    };

    drawSignatureLine("Checked By", 50, signatureSectionY);
    drawSignatureLine("Approved By", 250, signatureSectionY);
    drawSignatureLine("Receiver Signature", 450, signatureSectionY);

    doc.end();

    pdfStream.on("finish", async () => {
      try {
        const drive = google.drive({ version: "v3", auth });
        console.log(`Uploading PDF ${pdfFileName} to Drive folder ${folderId}`);
        const pdfFileMetadata = {
          name: pdfFileName,
          parents: [folderId],
        };
        const pdfMedia = {
          mimeType: "application/pdf",
          body: fs.createReadStream(pdfFilePath),
        };
        const pdfUploadResponse = await drive.files.create({
          resource: pdfFileMetadata,
          media: pdfMedia,
          fields: "id, webViewLink",
        });

        const pdfFileId = pdfUploadResponse.data.id;
        const pdfLink = pdfUploadResponse.data.webViewLink;
        console.log(`PDF uploaded: ${pdfFileId}, Link: ${pdfLink}`);

        const sheets = google.sheets({ version: "v4", auth });
        const values = [
          [
            voucherNo,
            voucherData.date,
            voucherData.filter,
            voucherData.payTo,
            voucherData.accountHead,
            voucherData.account,
            voucherData.amount,
            voucherData.amountRs,
            voucherData.checkedBy,
            voucherData.approvedBy,
            voucherData.receiverSignature,
            pdfLink,
          ],
        ];

        console.log(`Appending data to sheet ${spreadsheetId}`);
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${sheetTitle}!A:O`,
          valueInputOption: "RAW",
          requestBody: {
            values,
          },
        });
        console.log(`Data appended to ${spreadsheetId}`);

        const voucher = new Voucher({
          email,
          company: filterOption,
          voucherNo,
          date: voucherData.date,
          payTo: voucherData.payTo,
          accountHead: voucherData.accountHead,
          account: voucherData.account,
          amount: voucherData.amount,
          amountRs: voucherData.amountRs,
          checkedBy: voucherData.checkedBy,
          approvedBy: voucherData.approvedBy,
          receiverSignature: voucherData.receiverSignature,
          pdfLink,
          spreadsheetId,
          folderId,
          pdfFileId,
        });
        await voucher.save();
        console.log(`Voucher data saved to MongoDB: ${voucherNo}`);

        fs.unlinkSync(pdfFilePath);

        res.status(200).send({
          message: "Data submitted successfully and PDF uploaded!",
          sheetURL: sheetURL,
          pdfFileId: pdfFileId,
        });
      } catch (error) {
        console.error("Error uploading PDF or appending to sheet:", error.message);
        res.status(500).send({ error: "Failed to upload PDF or append to sheet: " + error.message });
      }
    });

    pdfStream.on("error", (error) => {
      console.error("Error creating PDF:", error.message);
      res.status(500).send({ error: "Failed to create PDF: " + error.message });
    });
  } catch (error) {
    console.error("Error in /submit endpoint:", error.message);
    res.status(500).send({ error: "Failed to submit data: " + error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});