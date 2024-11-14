var express = require("express");
var bodyParser = require("body-parser");
var mongoose = require("mongoose");
var multer = require("multer");
var csv = require("csv-parser");
var fs = require("fs");
var ExcelJS = require('exceljs');

const app = express();

app.use(bodyParser.json());
app.use(express.static('public'));
app.use(bodyParser.urlencoded({ extended: true }));

// Set up multer for file uploads
const upload = multer({ dest: 'uploads/' });

// MongoDB connection
mongoose.connect('mongodb://localhost:27017/data');
var db = mongoose.connection;
db.on('error', () => console.log("Error in Connecting to Database"));
db.once('open', () => console.log("Connected to Database"));

// Define a schema and model
var siteSchema = new mongoose.Schema({
    Date: String,
    Dist: String,
    Site_ID: String,
    Toco_ID: String,
    Toco_Owner: String,
    Cir: String,
    SOB_hrs: String,
    SOB_count: String,
    DG_hrs: String,
    Derived_DG_hrs: String,
    DG_count: String,
    MF_hrs: String,
    MF_count: String,
    LB_hrs: String,
    LB_count: String,
    SO_hrs: String,
    SO_count: String,
    Source: String,
    EB_On: String,
    EB_Fail: String,
    EB_DG_SOB_24hrs: String,
    SLA_Remarks: String,
    Alarm_Compliance: String,
    Alarm_Non_compliance: String,
    AQI_EB_DG_BB: String,
    createdAt: { type: Date, default: Date.now }
});
var Site = mongoose.model('EB_01_2024', siteSchema);

// Handle CSV upload
app.post("/upload_csv", upload.single('csvFile'), (req, res) => {
    const results = [];
    fs.createReadStream(req.file.path)
        .pipe(csv())
        .on('data', (data) => results.push(data))
        .on('end', () => {
            db.collection('EB_01_2024').insertMany(results, (err, result) => {
                if (err) {
                    console.error('Error inserting data:', err);
                    res.status(500).send('Error inserting data');
                } else {
                    console.log('Inserted:', result.insertedCount);
                    res.send('Data successfully imported');
                }
                fs.unlinkSync(req.file.path); // Remove the uploaded file
            });
        });
});

// Handle exporting data to Excel with optional date and dist filters
app.get("/export_excel", async (req, res) => {
    try {
        const { startDate, endDate, dist } = req.query;

        // Build query object based on the presence of startDate, endDate, and dist
        let query = {};
        if (startDate && endDate) {
            query.Date = {
                $gte: startDate,
                $lte: endDate
            };
        }
        if (dist) {
            query.Dist = dist;
        }

        console.log(`Filtering data with query: ${JSON.stringify(query)}`);

        const sites = await Site.find(query);

        if (!sites.length) {
            console.log('No data found for the given query.');
            return res.status(404).send('No data found for the given query.');
        }

        console.log(`Found ${sites.length} sites`);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Sites');

        worksheet.columns = [
            { header: 'Date', key: 'Date', width: 15 },
            { header: 'Dist', key: 'Dist', width: 20 },
            { header: 'Site ID', key: 'Site_ID', width: 20 },
            { header: 'Toco ID', key: 'Toco_ID', width: 20 },
            { header: 'Toco Owner', key: 'Toco_Owner', width: 20 },
            { header: 'Cir', key: 'Cir', width: 15 },
            { header: 'SOB hrs', key: 'SOB_hrs', width: 15 },
            { header: 'SOB count', key: 'SOB_count', width: 15 },
            { header: 'DG hrs', key: 'DG_hrs', width: 15 },
            { header: 'Derived DG hrs', key: 'Derived_DG_hrs', width: 20 },
            { header: 'DG count', key: 'DG_count', width: 15 },
            { header: 'MF hrs', key: 'MF_hrs', width: 15 },
            { header: 'MF count', key: 'MF_count', width: 15 },
            { header: 'LB hrs', key: 'LB_hrs', width: 15 },
            { header: 'LB count', key: 'LB_count', width: 15 },
            { header: 'SO hrs', key: 'SO_hrs', width: 15 },
            { header: 'SO count', key: 'SO_count', width: 15 },
            { header: 'Source', key: 'Source', width: 20 },
            { header: 'EB On', key: 'EB_On', width: 15 },
            { header: 'EB Fail', key: 'EB_Fail', width: 15 },
            { header: 'EB DG SOB 24hrs', key: 'EB_DG_SOB_24hrs', width: 20 },
            { header: 'SLA Remarks', key: 'SLA_Remarks', width: 20 },
            { header: 'Alarm Compliance', key: 'Alarm_Compliance', width: 20 },
            { header: 'Alarm Non-compliance', key: 'Alarm_Non_compliance', width: 20 },
            { header: 'AQI EB DG BB', key: 'AQI_EB_DG_BB', width: 20 }
        ];

        sites.forEach((site) => {
            worksheet.addRow(site.toObject());
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=sites.xlsx');

        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.error('Error exporting data to Excel:', err);
        res.status(500).send('Error exporting data to Excel.');
    }
});

app.get("/", (req, res) => {
    res.set({
        "Allow-access-Allow-Origin": '*'
    });
    return res.redirect('index.html');
}).listen(3000);

console.log("Listening on port 3000");
