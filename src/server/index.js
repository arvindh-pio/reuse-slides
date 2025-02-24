const express = require("express");
const dotenv = require("dotenv");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const { exec } = require("child_process");

dotenv.config();
const PORT = process.env.PORT || 5000;
const app = express();

app.use(cors());
app.use("/screenshots", express.static(path.join(__dirname, "screenshots")));

const uploadDir = path.join(__dirname, "uploads");
const outputDir = path.join(__dirname, "screenshots");
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

app.post("/upload", async (req, res) => {
    if (!req.files || !req.files.ppt) {
        return res.status(400).send("No file uploaded");
    }

    const pptFile = req.file.ppt;
    const pptName = path.parse(pptFile.name).name;
    const pptPath = path.join(uploadDir, pptFile.name);
    await pptFile.mv(pptPath);

    try {
        const slides = await converPptToImages(pptPath, pptName);
        res.json({ slides });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
})

async function converPptToImages(pptPath, pptName) {
    return new Promise((resolve, reject) => {
        const pdfPath = pptPath.replace(".pptx", ".pdf");

        const pptOutputDir = path.join(outputDir, pptName);
        if (!fs.existsSync(pptOutputDir)) fs.mkdirSync(pptOutputDir);

        // convert to pdf
        const convertToPdf = `"C:\\Program Files\\LibreOffice\\program\\soffice.exe" --headless --convert-to pdf --outdir ${path.dirname(pptPath)} ${pptPath}`;
        exec(convertToPdf, (error, stdout, stderr) => {
            if (error) return reject(new Error(`Error converting: ${stderr}`));

            // waiting for pdf to generate
            const checkFileExists = setInterval(() => {
                if (fs.existsSync(pdfPath)) {
                    clearInterval(checkFileExists);

                    // extract images
                    const extractImages = `magick -density 150 "${pdfPath}" "${pptOutputDir}/slide_%d.png`;
                    exec(extractImages, (err, out, errOut) => {
                        if (err) return reject(new Error(`Error extracting images: ${errOut}`));

                        fs.readdir(pptOutputDir, (err, files) => {
                            if (err) return reject(err);

                            const slideImages = files
                                .filter(file => file.startsWith("slide_") && file.endsWith(".png"))
                                .map(file => `screenshots/${pptName}/${file}`);
                            resolve(slideImages);
                        })
                    })
                }
            }, 500);
        })
    })
}

app.get("/", (_, res) => {
    return res.json({
        message: "Welcome to reuse slides apis"
    });
})

app.listen(PORT, () => {
    console.log(`Server is running on PORT ${PORT}`);
})