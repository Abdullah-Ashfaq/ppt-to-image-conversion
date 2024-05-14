const express = require("express");
const fs = require("fs");
const path = require("path");
const app = express();
const officegen = require("officegen");
const { fromPath } = require("pdf2pic");
const PDFImage = require("pdf-image").PDFImage;
const { convertPDFToImages } = require("pdf2image");
const { pdf } = require("pdf-to-img");
const shell = require("shelljs");
const { removeBackground } = require("@imgly/background-removal-node");
const { exec } = require("child_process");
require("dotenv").config({});
const pptConvert = require("./pptConvert.controller");

function convertPowerPointToImages(pptFilePath, outputFolder, res) {
  try {
    // const pptData = fs.readFileSync(pptFilePath);
    // const ppt = officegen("pptx");
    // const pptObj = ppt.read(pptData);
    // const slides = ppt.getSlides(pptObj);
    // if (!fs.existsSync(outputFolder)) {
    //   fs.mkdirSync(outputFolder, { recursive: true });
    // }
    // slides.forEach((slide, index) => {
    //   const outputFilePath = path.join(outputFolder, `slide_${index + 1}.jpg`);
    //   const outputData = slide.slide.data;
    //   fs.writeFileSync(outputFilePath, outputData, "binary");
    // });
    // console.log("PowerPoint conversion completed successfully.");
    // const pptx = officegen("pptx");
    const pptx = officegen({
      type: "pptx",
      title: "Presentation",
    });
    const pptxFilePath = path.join(__dirname, "./ppt/one.pptx");
    const imagesFolder = path.join(__dirname, "output");
    if (!fs.existsSync(imagesFolder)) {
      fs.mkdirSync(imagesFolder);
    }

    const convertOptions = {
      outputDir: imagesFolder,
      quality: 100, // Adjust quality if needed
      format: "png",
    };

    const convertedImages = [];
    for (let i = 0; i < pptx.slides.length; i++) {
      const imagePath = path.join(imagesFolder, `slide${i + 1}.png`);
      pptx.slides[i].generateImageSync(imagePath, convertOptions);
      console.log(`Slide ${i + 1} converted to image: ${imagePath}`);
      convertedImages.push(imagePath);
    }

    console.log("All slides converted to images:", convertedImages);
    res.json({ images: convertedImages });
  } catch (err) {
    console.error("Error converting PowerPoint:", err);
    return res.send(err.message);
  }
}

// app.get("/", async (req, res) => {
//   // convertPowerPointToImages("./ppt/one.pptx", "./output", res);
//   try {
//     const pptxFilePath = path.join(__dirname, "./ppt/one.pptx");
//     const imagesFolder = path.join(__dirname, "output");

//     if (!fs.existsSync(imagesFolder)) {
//       fs.mkdirSync(imagesFolder);
//     }

//     const command = `unoconv -f png -o ${imagesFolder} ${pptxFilePath}`;

//     exec(command, (error, stdout, stderr) => {
//       if (error) {
//         console.error(`Error converting PowerPoint: ${error.message}`);
//         res.status(500).send("Error converting PowerPoint");
//         return;
//       }

//       if (stderr) {
//         console.error(`unoconv stderr: ${stderr}`);
//       }

//       console.log(`PowerPoint converted to PNG: ${stdout}`);

//       // Check if files were created in the output folder
//       fs.readdir(imagesFolder, (err, files) => {
//         if (err) {
//           console.error(`Error reading output folder: ${err.message}`);
//           res.status(500).send("Error reading output folder");
//           return;
//         }

//         console.log("Files in output folder:", files);
//         res.send("PowerPoint converted to PNG successfully!");
//       });
//     });
//     // return res.send("Generating files..")
//   } catch (error) {
//     res.send(error.message);
//   }
// });

app.get("/convert", async (req, res) => {
  try {
    const pptFilePath = path.join(__dirname, "./ppt/one.pptx");
    // const outputFolder = path.join(__dirname, "./output");
    // const pptData = fs.readFileSync(pptFilePath);
    // const ppt = officegen("pptx");
    // const pptObj = ppt.read(pptData);
    // const slides = ppt.getSlides(pptObj);
    // if (!fs.existsSync(outputFolder)) {
    //   fs.mkdirSync(outputFolder, { recursive: true });
    // }
    // slides.forEach((slide, index) => {
    //   const outputFilePath = path.join(outputFolder, `slide_${index + 1}.jpg`);
    //   const outputData = slide.slide.data;
    //   fs.writeFileSync(outputFilePath, outputData, "binary");
    // });
    // console.log("PowerPoint conversion completed successfully.");
    const convert = new pptConvert(pptFilePath);

    convert.start().then((result) => {
      if (result.split(";") && result.split(";")[0] == "ok") {
        const res = result.split(";");
        let images = [];
        for (let i = 1; i <= res[2]; i++) {
          images.push(`http://localhost:8000/slideshows/${res[1]}_${i}.jpg`);
        }
        response.json({
          status: true,
          message: result,
          total_page: res[2],
          images: images,
        });
      }
    });
  } catch (error) {
    console.log("error........", error);
    res.send(error.message);
  }
});

// app.get("/pdf-to-images", async (req, res) => {
//   const options = {
//     density: 100,
//     saveFilename: "untitled",
//     savePath: "./images",
//     format: "png",
//     width: 600,
//     height: 600,
//   };
//   const convert = fromPath("./uploads/one.pdf", options);
//   const pageToConvertAsImage = 3;

//   convert(pageToConvertAsImage, { responseType: "image" }).then((resolve) => {
//     console.log("Page 1 is now converted as image");

//     return resolve;
//   });
// });

app.get("/image-pdf", async (req, res) => {
  // const pdfFilePath = "./uploads/one.pdf";
  const pdfFilePath = path.join(__dirname, "./uploads/one.pdf");
  const outputDir = "./output";
  const pdfImage = new PDFImage(pdfFilePath, {
    convertOptions: {
      "-quality": "100", // Image quality (0-100)
      "-density": "300", // Image density (DPI)
      "-flatten": null, // Flatten transparency
      "-trim": null, // Trim whitespace
    },
  });

  pdfImage
    .convertFile()
    .then((imagePaths) => {
      console.log("Images converted successfully:", imagePaths);
    })
    .catch((error) => {
      console.error("Error converting PDF to images:", error);
    });
});

app.get("/pdf2image", async (req, res) => {
  // Path to your PDF file
  const pdfFilePath = path.join(__dirname, "./uploads/one.pdf");

  // Options for conversion
  const options = {
    outputFormat: "jpeg", // Output image format
    quality: 100, // Image quality (0-100)
    dpi: 300, // DPI (dots per inch)
    width: 1024, // Image width
    height: 768, // Image height
    pages: "1-3", // Pages to convert (e.g., '1', '1,2,3', '1-5')
    singleFile: false, // Convert each page to a separate file
    outputDir: path.join(__dirname, "./output"), // Output directory
  };

  // Convert PDF to images
  convertPDFToImages(pdfFilePath, options)
    .then((images) => {
      console.log("Images converted:", images);
    })
    .catch((error) => {
      console.error("Error converting PDF to images:", error);
    });
});

app.get("/pdf2pic", async (req, res) => {
  const options = {
    density: 100,
    saveFilename: "Test",
    savePath: "./output",
    format: "png",
    width: "100%",
    height: "100%",
  };
  const convert = fromPath(path.join(__dirname, "./uploads/one.pdf"), options);
  const pageToConvertAsImage = 1;

  // convert(1, { responseType: "images" }).then((resolve) => {
  //   console.log("Page 1 is now converted as image");

  //   return resolve;
  // });
  try {
    const result = await convert.bulk(-1);
    console.log("result", result);
    // Move images to a subdirectory with PDF name
    const pdfName = path.basename("./output/test", ".pdf");
    const pdfOutputDir = path.join(__dirname, pdfName);
    // await fs.ensureDir(pdfOutputDir);
    // await fs.move(result, pdfOutputDir, { overwrite: true });

    console.log(`PDF converted to images: ${pdfOutputDir}`);

    res.send(result);
  } catch (error) {
    res.send(error.message);
  }
});

app.get("/ppt2pdf", async (req, res) => {
  try {
    const file = path.join(__dirname, "./ppt/one.pptx");
    const fileUrl = await convertToPDF(file, ".pptx");

    const images = await convertPdfToImages(fileUrl);

    const imagesWithoutBG = images.map(
      async (img) => await removeBg(img?.path, img.name)
    );
    const imagesResult = await Promise.all(imagesWithoutBG);
    res.send(imagesResult);
  } catch (error) {
    res.send(error.message);
  }
});
function convertToPDF(file, ext) {
  return new Promise((resolve, reject) => {
    const office = "soffice";
    const outdir = path.join(__dirname, "./uploads");
    if (!fs.existsSync(file)) reject(Error("File Not Found"));

    if (ext === ".ppt" || ext === ".pptx" || ext == ".pps" || ext == ".ppsx") {
      const commandOffice = `"${office}" --headless --convert-to pdf --outdir "${outdir}" "${file}"`;
      console.log("commandOffice", commandOffice);
      shell.exec(commandOffice, (err, stdout, stderr) => {
        const pdf = file.replace(ext, ".pdf");
        const fileName = path.basename(pdf);
        if (err) {
          reject(Error(err));
        }
        resolve(outdir + "/" + fileName);
      });
    } else {
      reject(Error("Only selected extensions are allowed"));
    }
  });
}
async function convertPdfToImages(file) {
  console.log("file........", file);
  if (!fs.existsSync(file)) {
    throw Error("File Not Found" + fs.existsSync(file));
    return;
  }
  const options = {
    density: 100,
    saveFilename: "Test",
    savePath: "./output",
    format: "png",
    width: "100%",
    height: "100%",
  };
  const convert = fromPath(file, options);
  const pageToConvertAsImage = 1;

  try {
    const result = await convert.bulk(-1);
    console.log("result", result);
    // Move images to a subdirectory with PDF name
    const pdfName = path.basename("./output/test", ".pdf");
    const pdfOutputDir = path.join(__dirname, pdfName);
    // await fs.ensureDir(pdfOutputDir);
    // await fs.move(result, pdfOutputDir, { overwrite: true });

    console.log(`PDF converted to images: ${pdfOutputDir}`);

    return result;
  } catch (error) {
    throw error;
  }
}

async function removeBg(img, name) {
  // const image_buffer = fs.readFileSync();
  return await removeBackground(path.join(__dirname, img))
    .then(async (blob) => {
      console.log("Image blog......", blob);
      // Convert Blob to Buffer
      const arrayBuffer = await blob.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);
      // The result is a blob encoded as PNG
      const imagePath = path.join(__dirname, "/imageWithoutBg/" + name);
      fs.writeFileSync(imagePath, buffer);
      console.log("Image saved to", imagePath);
      fs.unlinkSync(img);
      return imagePath;
    })
    .catch((err) => {
      console.error(err);
      return null;
    });
}
app.listen(8000, () => {
  console.log("Example app listening on port 8000!");
});
