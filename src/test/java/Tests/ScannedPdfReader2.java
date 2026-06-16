package Tests;

import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;

import java.awt.image.BufferedImage;
import java.io.File;
import java.nio.file.*;
import java.util.Comparator;
import java.util.Optional;
import java.util.stream.Stream;

public class ScannedPdfReader2 {

    public static void main(String[] args) {

        try {
            // 1. Get latest downloaded PDF from Downloads folder
            File latestPdf = getLatestPdfFromDownloads();
 
            if (latestPdf == null) {
                System.out.println("No PDF file found in Downloads folder.");
                return;
            }

            System.out.println("Reading file: " + latestPdf.getAbsolutePath());

            // 2. Extract text
            String extractedText = extractTextFromScannedPdf(latestPdf.getAbsolutePath());

            

            // 3. Delete PDF after reading
            boolean deleted = latestPdf.delete();

            if (deleted) {
                System.out.println("PDF deleted successfully.");
            } else {
                System.out.println("Failed to delete PDF.");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    
 // Method to get latest PDF file from Downloads folder
    public static File getLatestPdfFromDownloads() throws Exception {

        String downloadsPath = System.getProperty("user.home") + "\\Downloads";
        //////System.out.println("downloadsPath >>>" + downloadsPath);

        long fiveMinutesAgo = System.currentTimeMillis() - (5 * 60 * 1000);

        try (Stream<Path> files = Files.list(Paths.get(downloadsPath))) {

            Optional<Path> latestFile = files
                    .filter(path -> path.toString().toLowerCase().endsWith(".pdf"))
                    .filter(path -> path.toFile().lastModified() >= fiveMinutesAgo)
                    .max(Comparator.comparingLong(path -> path.toFile().lastModified()));

            return latestFile.map(Path::toFile).orElse(null);
        }
    }

//    // Method to get latest PDF file from Downloads folder
//    public static File getLatestPdfFromDownloads() throws Exception {
//
//        String downloadsPath = System.getProperty("user.home") + "\\Downloads";
//        //////System.out.println("downloadsPath >>>" + downloadsPath);
// 
//        try (Stream<Path> files = Files.list(Paths.get(downloadsPath))) {
//
//            Optional<Path> latestFile = files
//                    .filter(path -> path.toString().toLowerCase().endsWith(".pdf"))
//                    .max(Comparator.comparingLong(path -> path.toFile().lastModified()));
//
//            return latestFile.map(Path::toFile).orElse(null);
//        } 
//    }

    public static String extractTextFromScannedPdf(String pdfPath) throws Exception {

        File pdfFile = new File(pdfPath);

        // =====================================================
        // STEP 1: Try the embedded text layer first.
        // Digital PDFs (e.g. Zoho/system-generated invoices) carry
        // exact text — values like 306.50 / 75.00 come out
        // character-perfect, no OCR misreads.
        // =====================================================
        try (PDDocument document = Loader.loadPDF(pdfFile)) {

            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setSortByPosition(true);

            String embeddedText = stripper.getText(document);

            // A scanned PDF returns empty/near-empty text here
            if (embeddedText != null && embeddedText.trim().length() > 50) {
                System.out.println("Text layer found — extracted without OCR.");
                return embeddedText;
            }

            System.out.println("No usable text layer — falling back to OCR.");
        }

        // =====================================================
        // STEP 2: OCR fallback for scanned PDFs
        // =====================================================
        Tesseract tesseract = new Tesseract();
        tesseract.setDatapath("c:\\tessdata");
        tesseract.setLanguage("eng");
        // Keep column spacing so table rows stay readable
        tesseract.setPageSegMode(6);
        tesseract.setVariable("preserve_interword_spaces", "1");

        StringBuilder fullText = new StringBuilder();

        // Auto close document using try-with-resources
        try (PDDocument document = Loader.loadPDF(pdfFile)) {

            PDFRenderer renderer = new PDFRenderer(document);

            int pageCount = document.getNumberOfPages();

            // Render each page to image and run OCR
            for (int page = 0; page < pageCount; page++) {

                BufferedImage image = renderer.renderImageWithDPI(page, 300);

                try {
                    String pageText = tesseract.doOCR(image);

                    fullText.append("--- Page ")
                            .append(page + 1)
                            .append(" ---\n");

                    fullText.append(pageText).append("\n");

                } catch (TesseractException e) {

                    System.err.println("OCR failed on page "
                            + (page + 1)
                            + ": "
                            + e.getMessage());
                }
            } 
        }
        //System.out.println("Extracted text from Pdf "+fullText.toString());
        return fullText.toString();
    }
}