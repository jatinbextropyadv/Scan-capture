package Tests;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.LoadState;
import com.microsoft.playwright.options.WaitForSelectorState;

// ===== Fillo (replaces Apache POI for ALL Excel work) =====
import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

// ===== POI is kept ONLY for the Word (.docx) part — Fillo cannot read Word files =====
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
// POI used once at startup only to ADD missing result columns (Fillo cannot create columns)
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class existingChromeFillo {

    // =====================================================
    // SHEET / COLUMN NAMES USED IN FILLO QUERIES
    //
    // Fillo addresses data by SHEET NAME + HEADER NAME (row 1
    // of each sheet must be a header row). Adjust these to
    // match your actual files.
    //
    // NOTE: Fillo does not handle spaces in column names
    // reliably inside "Update ... Set" queries, so the result
    // columns use underscores ("Expected_Value" instead of
    // "Expected value"). ensureResultColumns() creates them
    // automatically if they don't exist yet.
    // =====================================================
    static final String CENTERS_SHEET   = "Sheet1";          // sheet holding the "id" column
    static final String COL_ID          = "id";
    static final String COL_COMMENTS    = "Comments";
    static final String COL_EXPECTED    = "Expected_Value";
    static final String COL_ACTUAL      = "Actual_Value";

    static final String CKG_SHEET       = "English";
    static final String CKG_COL_FIELD   = "Field";            // was column B
    static final String CKG_COL_KEY1    = "Keyword1";         // was column C
    static final String CKG_COL_KEY2    = "Keyword2";         // was column D
    static final String CKG_COL_KEY3    = "Keyword3";         // was column E

    static final String EXTRACTS_SHEET  = "Sheet1";
    static final String COL_SUPPLIER    = "Supplier Name";

    public static String newComment = "";
    public static String Suppliername_new = "";

    // =====================================================
    // HELPER METHODS
    // =====================================================

    private static String normalize(String text) {

        if (text == null) {
            return "";
        }

        return text
                .replace("\r", " ")
                .replace("\n", " ")
                .replaceAll("\\s+", " ")
                .trim()
                .toLowerCase();
    }

    // Fillo chokes on single quotes inside query values — strip them
    private static String q(String value) {
        if (value == null) return "";
        return value.replace("'", "");
    }

    private static String safeField(Recordset rs, String column) {
        try {
            String v = rs.getField(column);
            return v == null ? "" : v.trim();
        } catch (FilloException e) {
            return "";
        }
    }

    // =====================================================
    // ONE-TIME BOOTSTRAP (POI)
    // Fillo cannot add new columns to a sheet, so the three
    // result columns are created here once if missing.
    // =====================================================
    public static void ensureResultColumns(String excelPath) throws Exception {

        try (FileInputStream fis = new FileInputStream(excelPath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            boolean hasComments = false, hasExpected = false, hasActual = false, changed = false;

            for (Cell cell : headerRow) {
                String col = cell.getStringCellValue().trim();
                if (col.equalsIgnoreCase(COL_COMMENTS)) hasComments = true;
                if (col.equalsIgnoreCase(COL_EXPECTED)) hasExpected = true;
                if (col.equalsIgnoreCase(COL_ACTUAL))   hasActual = true;
            }

            if (!hasComments) {
                headerRow.createCell(headerRow.getLastCellNum()).setCellValue(COL_COMMENTS);
                changed = true;
            }
            if (!hasExpected) {
                headerRow.createCell(headerRow.getLastCellNum()).setCellValue(COL_EXPECTED);
                changed = true;
            }
            if (!hasActual) {
                headerRow.createCell(headerRow.getLastCellNum()).setCellValue(COL_ACTUAL);
                changed = true;
            }

            if (changed) {
                try (FileOutputStream fos = new FileOutputStream(excelPath)) {
                    workbook.write(fos);
                }
                System.out.println("Result columns created in centers Excel.");
            }
        }
    }

    // =====================================================
    // ADD COMMENT (Fillo version)
    //
    // POI version received Cell objects and called setCellValue();
    // here the row is addressed by its id via WHERE clause, existing
    // values are read with SELECT and written back with UPDATE.
    // Fillo persists each executeUpdate immediately — no
    // workbook.write() needed anywhere.
    // =====================================================
    public static void Add_comment(
            Connection centersConn,
            String id,
            String newComment,
            String Suppliername_new,
            String Suppliername_old) throws FilloException {

        // Read existing values for this row
        String existingComment = "";
        String existingExpected_value = "";
        String existingActual_value = "";

        Recordset rs = centersConn.executeQuery(
                "Select * from " + CENTERS_SHEET + " where " + COL_ID + "='" + q(id) + "'");
        if (rs.next()) {
            existingComment        = safeField(rs, COL_COMMENTS);
            existingExpected_value = safeField(rs, COL_EXPECTED);
            existingActual_value   = safeField(rs, COL_ACTUAL);
        }
        rs.close();

        // Empty old value is recorded as "Blank" so the cell is never empty
        String actualValue = (Suppliername_old == null || Suppliername_old.trim().isEmpty())
                ? "Blank"
                : Suppliername_old.trim();

        if (!existingComment.isEmpty() && !existingComment.equals(newComment)) {
            updateCell(centersConn, id, COL_COMMENTS, existingComment + " ," + newComment);
        } else if (existingComment.isEmpty()) {
            updateCell(centersConn, id, COL_COMMENTS, newComment);
        }

        if (!existingExpected_value.isEmpty() && !existingExpected_value.equals(Suppliername_new)) {
            updateCell(centersConn, id, COL_EXPECTED, existingExpected_value + " ," + Suppliername_new);
        } else if (existingExpected_value.isEmpty()) {
            updateCell(centersConn, id, COL_EXPECTED, Suppliername_new);
        }

        if (!existingActual_value.isEmpty() && !existingActual_value.equals(actualValue)) {
            updateCell(centersConn, id, COL_ACTUAL, existingActual_value + " ," + actualValue);
        } else if (existingActual_value.isEmpty()) {
            updateCell(centersConn, id, COL_ACTUAL, actualValue);
        }
    }

    private static void updateCell(Connection conn, String id, String column, String value)
            throws FilloException {

        conn.executeUpdate(
                "Update " + CENTERS_SHEET +
                " Set " + column + "='" + q(value) + "'" +
                " where " + COL_ID + "='" + q(id) + "'");
    }

    // =====================================================
    // WORD DOCUMENT CHECK — unchanged (POI XWPF).
    // Fillo is Excel-only and cannot parse .docx tables.
    // =====================================================
    public static boolean isFieldMandatory(
            String docPath,
            String fieldName) throws Exception {

        String searchField = normalize(fieldName);

        try (FileInputStream fis = new FileInputStream(docPath);
             XWPFDocument doc = new XWPFDocument(fis)) {

            int tableNo = 0;

            for (XWPFTable table : doc.getTables()) {

                tableNo++;
                int rowNo = 0;

                for (XWPFTableRow row : table.getRows()) {

                    rowNo++;

                    List<XWPFTableCell> cells = row.getTableCells();

                    boolean fieldFound = false;
                    boolean hasCheckbox = false;
                    Boolean checkboxValue = null;

                    for (XWPFTableCell cell : cells) {

                        String cellText = normalize(cell.getText());

                        if (cellText.equals(searchField) || cellText.contains(searchField)) {
                            fieldFound = true;
                        }

                        if (cell.getText().contains("☒")) {
                            hasCheckbox = true;
                            checkboxValue = true;
                        }

                        if (cell.getText().contains("☐")) {
                            hasCheckbox = true;
                            checkboxValue = false;
                        }
                    }

                    if (fieldFound) {

                        System.out.println("\n================================");
                        System.out.println("Field Found : " + fieldName);
                        System.out.println("Table : " + tableNo + ", Row : " + rowNo);

                        if (hasCheckbox == true && checkboxValue == true) {
                            return true;
                        }

                        System.out.println("Field matched but row has no checkbox. Continuing search...");
                    }
                }
            }
        }

        System.out.println("Field not found or checkbox not found.");
        return false;
    }

    // =====================================================
    // LOAD CKG SHEET ONCE (Fillo version)
    //
    // POI iterated cells by index (B,C,D,E); Fillo reads by
    // header name, so the whole sheet is loaded into a list of
    // String[4] rows {field, keyword1, keyword2, keyword3}
    // preserving sheet order, then the original row-walking
    // logic runs on that list.
    // =====================================================
    private static List<String[]> loadCkgRows() throws FilloException {

        Fillo fillo = new Fillo();
        Connection conn = fillo.getConnection(ConfigReader.get("ckg.excel.path"));

        List<String[]> rows = new ArrayList<>();

        Recordset rs = conn.executeQuery("Select * from " + CKG_SHEET);
        while (rs.next()) {
            rows.add(new String[] {
                    safeField(rs, CKG_COL_FIELD),
                    safeField(rs, CKG_COL_KEY1),
                    safeField(rs, CKG_COL_KEY2),
                    safeField(rs, CKG_COL_KEY3)
            });
        }
        rs.close();
        conn.close();

        return rows;
    }

    private static boolean Keyword_Found_in_CKG(
            String searchText,
            String pdfText,
            List<String[]> ckgRows) {

        int startRow = -1;

        // Find text in the Field column (was column B)
        for (int i = 0; i < ckgRows.size(); i++) {

            String colB = ckgRows.get(i)[0];

            if (!colB.isEmpty() && colB.equalsIgnoreCase(searchText)) {
                startRow = i + 1; // DELIBERATELY DO THIS plus ONE +1 EARLIER IT WAS I ONLY
                System.out.println("STARTROW " + startRow);
                System.out.println("colB " + colB);
                break;
            }
        }

        if (startRow == -1) {
            return false;
        }

        boolean keywordFound = false;

        // Iterate until blank row
        for (int i = startRow; i < ckgRows.size(); i++) {

            String c = ckgRows.get(i)[1];
            String d = ckgRows.get(i)[2];
            String e = ckgRows.get(i)[3];

            // Stop when C, D, E all blank
            if (c.isEmpty() && d.isEmpty() && e.isEmpty()) {
                break;
            }

            String[] keywords = {c, d, e};

            for (String keyword : keywords) {

                if (keyword.isEmpty()) {
                    continue;
                }

                String str = pdfText.toLowerCase();
                String keyword_l = keyword.toLowerCase();

                if (str.contains(keyword_l)) {
                    System.out.println("MATCH FOUND => " + keyword_l);
                    keywordFound = true;
                }
            }
        }

        System.out.println("ISssss Keyword_Found_in_CKG " + keywordFound);
        return keywordFound;
    }

    public static boolean compareFieldWithCKG(String popupField, String pdfText) throws Exception {

        List<String[]> ckgRows = loadCkgRows();

        boolean fieldMatched = false;
        boolean allKeywordsMatched = true;

        for (String[] row : ckgRows) {

            String fieldName = row[0];

            if (!fieldMatched) {

                if (!fieldName.isEmpty() && fieldName.equalsIgnoreCase(popupField.trim())) {

                    fieldMatched = true;
                    System.out.println("Matched Field >>> " + fieldName);
                    allKeywordsMatched = processKeywords(row, pdfText);
                }
            } else {

                if (!fieldName.isEmpty()) {
                    System.out.println("Next Field Encountered >>> " + fieldName);
                    break;
                }

                if (!processKeywords(row, pdfText)) {
                    allKeywordsMatched = false;
                }
            }
        }

        return fieldMatched && allKeywordsMatched;
    }

    private static boolean processKeywords(String[] row, String pdfText) {

        boolean allKeywordsFound = true;

        // Keyword columns (were C, D, E)
        for (int col = 1; col <= 3; col++) {

            String keyword = row[col];

            if (keyword.isEmpty()) {
                continue;
            }

            if (pdfText.toLowerCase().contains(keyword.toLowerCase())) {
                // keyword found in PDF
            } else {
                allKeywordsFound = false;
            }
        }

        return allKeywordsFound;
    }

    public static void checkKeywordInPdf(String keyword, String pdfText) {

        if (keyword == null || keyword.trim().isEmpty())
            return;

        if (pdfText.toLowerCase().contains(keyword.toLowerCase())) {
            // keyword present in PDF
        } else {
            // keyword NOT present in PDF
        }
    }

    // =====================================================
    // SUPPLIER COUNT IN EXTRACTS (Fillo version)
    //
    // The comparison stays in Java (not a WHERE clause) because
    // the original normalizes whitespace and compares
    // case-insensitively — Fillo WHERE is exact-match only.
    // =====================================================
    public static int getSupplierCountInExtracts(String supplierName) throws Exception {

        int count = 0;

        String searchSupplier = supplierName == null
                ? ""
                : supplierName.replace("\r", " ")
                              .replace("\n", " ")
                              .replaceAll("\\s+", " ")
                              .trim();

        Fillo fillo = new Fillo();
        Connection conn = fillo.getConnection(ConfigReader.get("extracts.excel.path"));

        Recordset rs = conn.executeQuery("Select * from " + EXTRACTS_SHEET);

        while (rs.next()) {

            String excelSupplier = safeField(rs, COL_SUPPLIER)
                    .replaceAll("\\s+", " ")
                    .trim();

            if (excelSupplier.equalsIgnoreCase(searchSupplier)) {
                count++;
            }
        }

        rs.close();
        conn.close();

        return count;
    }

    // =====================================================
    // MAIN FUNCTION
    // =====================================================

    public static void main(String[] args) throws Exception {

        String excelPath = ConfigReader.get("centers.excel.path");
        String baseUrl = ConfigReader.get("base.url");

        String Suppliername_old = "";
        String Suppliername_new = "";
        String New_value = "";
        String Old_value = "";

        String[] Field_values = {""};

        boolean FieldValue_Old_foundInInvoice = false;
        boolean Field_Name_FoundInInvoice = false;
        String pdfText = "";

        try {
            Playwright playwright = Playwright.create();

            // One-time: make sure result columns exist (Fillo can't create columns)
            ensureResultColumns(excelPath);

            // Fillo connection to the centers workbook — replaces
            // FileInputStream + XSSFWorkbook + all column-index hunting
            Fillo fillo = new Fillo();
            Connection centersConn = fillo.getConnection(excelPath);

            // Connect to already opened Chrome browser
            Browser browser = playwright.chromium()
                    .connectOverCDP("http://127.0.0.1:9224");

            System.out.println("Connected to existing Chrome browser");

            Thread.sleep(5000);

            // Get existing browser context
            BrowserContext context = browser.contexts().get(0);

            Page page;

            if (context.pages().size() > 0) {
                page = context.pages().get(0);
            } else {
                page = context.newPage();
            }

            // =====================================================
            // COLLECT ALL IDs FIRST
            // (collected up-front so the recordset isn't held open
            // while UPDATE queries run against the same file)
            // =====================================================
            List<String> ids = new ArrayList<>();

            Recordset idRs = centersConn.executeQuery("Select * from " + CENTERS_SHEET);
            while (idRs.next()) {
                String id = safeField(idRs, COL_ID);
                if (!id.isEmpty()) {
                    ids.add(id);
                }
            }
            idRs.close();

            if (ids.isEmpty()) {
                System.out.println("No ids found in Excel");
                return;
            }

            // =====================================================
            // ITERATE THROUGH EACH INVOICE ROW
            // =====================================================

            for (int i = 0; i < ids.size(); i++) {

                String id = ids.get(i);

                // Clear results from previous runs so Comments / Expected /
                // Actual stay aligned and don't accumulate stale entries
                updateCell(centersConn, id, COL_COMMENTS, "");
                updateCell(centersConn, id, COL_EXPECTED, "");
                updateCell(centersConn, id, COL_ACTUAL, "");

                String finalUrl = baseUrl + id;
                System.out.println("\nOpening URL: " + finalUrl);

                // Navigate to invoice page
                page.navigate(finalUrl);
                Thread.sleep(4000);

                // =====================================================
                // STEP 1: CLICK "Images" TAB TO LOAD PDF VIEWER
                // The PSPDFKit viewer + download button only appear
                // after the Images tab is active
                // =====================================================
                pdfText = "";
                Page pdfPage = null;
                Download download = null;

                try {
                    // Click the Images tab to activate the PDF viewer
                    Locator imagesTab = page.locator("//div[contains(@class,'tab') and contains(text(),'Images')]");

                    if (imagesTab.count() > 0) {
                        imagesTab.first().click();
                        System.out.println("Clicked Images tab.");
                    } else {
                        // Try alternate locator for Images tab
                        Locator imagesTab2 = page.locator("//button[contains(text(),'Images') or contains(@aria-label,'Images')]");
                        if (imagesTab2.count() > 0) {
                            imagesTab2.first().click();
                            System.out.println("Clicked Images tab (alternate locator).");
                        } else {
                            System.out.println("Images tab not found — PDF viewer may already be loaded.");
                        }
                    }

                    // Wait for PDF viewer toolbar to fully load
                    Thread.sleep(4000);

                    // =====================================================
                    // STEP 2: FIND DOWNLOAD BUTTON & CAPTURE NEW TAB
                    // PSPDFKit renders inside an iframe, so try frameLocator first
                    // =====================================================

                    Locator downloadBtn = null;

                    // Try inside common iframe selectors first (PSPDFKit uses iframes)
                    String[] iframeSelectors = {
                        "iframe[id*='pspdfkit']",
                        "iframe[class*='pspdfkit']",
                        "iframe[src*='pdf']",
                        "iframe"
                    };

                    for (String iframeSel : iframeSelectors) {
                        try {
                            Locator iframeEl = page.locator(iframeSel);
                            if (iframeEl.count() > 0) {
                                Locator candidate = page.frameLocator(iframeSel)
                                        .locator("button[data-block-id='download-button']");
                                candidate.waitFor(new Locator.WaitForOptions()
                                        .setState(WaitForSelectorState.VISIBLE)
                                        .setTimeout(8000));
                                downloadBtn = candidate;
                                System.out.println("Download button found inside iframe: " + iframeSel);
                                break;
                            }
                        } catch (Exception ignored) {}
                    }

                    // Fallback: try directly on the page (no iframe)
                    if (downloadBtn == null) {
                        Locator candidate = page.locator("button[data-block-id='download-button']");
                        candidate.waitFor(new Locator.WaitForOptions()
                                .setState(WaitForSelectorState.VISIBLE)
                                .setTimeout(15000));
                        downloadBtn = candidate;
                    }

                    System.out.println("Download button found. Clicking...");

                    final Locator finalDownloadBtn = downloadBtn;

                    // PSPDFKit's download button usually triggers a file download
                    // event rather than opening a new tab. Capture the Download
                    // first; if that doesn't fire, fall back to a new-tab check.
                    int pagesBefore = context.pages().size();

                    try {
                        download = page.waitForDownload(
                            new Page.WaitForDownloadOptions().setTimeout(20000),
                            () -> finalDownloadBtn.click()
                        );
                        System.out.println("Download captured: " + download.suggestedFilename());
                    } catch (Exception dlEx) {
                        System.out.println("No download event captured: " + dlEx.getMessage());

                        // The click already happened — a new tab may have opened instead
                        Thread.sleep(3000);
                        if (context.pages().size() > pagesBefore) {
                            pdfPage = context.pages().get(context.pages().size() - 1);
                            System.out.println("New tab detected instead of download event.");
                        }
                    }

                } catch (Exception ex) {
                    System.out.println("Could not click download button or capture new tab: " + ex.getMessage());
                    System.out.println("Skipping PDF extraction for ID: " + id);
                }

                // =====================================================
                // STEP 3: SAVE PDF (FROM DOWNLOAD EVENT OR NEW TAB) & OCR
                // =====================================================
                String downloadsPath = System.getProperty("user.home") + "\\Downloads";
                String fileName = "invoice_" + id + ".pdf";
                Path savedPath = null;

                if (download != null) {

                    try {
                        savedPath = Paths.get(downloadsPath, fileName);
                        download.saveAs(savedPath);
                        System.out.println("PDF saved to: " + savedPath);
                    } catch (Exception saveEx) {
                        System.out.println("Error saving download for ID " + id + ": " + saveEx.getMessage());
                        savedPath = null;
                    }

                } else if (pdfPage != null) {

                    try {
                        pdfPage.waitForLoadState(LoadState.LOAD);

                        String pdfUrl = pdfPage.url();
                        System.out.println("PDF opened in new tab URL: " + pdfUrl);

                        // Fetch PDF as bytes using JavaScript fetch()
                        Object result = pdfPage.evaluate(
                            "async () => {" +
                            "  const res = await fetch(window.location.href);" +
                            "  const buf = await res.arrayBuffer();" +
                            "  return Array.from(new Uint8Array(buf));" +
                            "}"
                        );

                        // Convert result to byte array
                        @SuppressWarnings("unchecked")
                        List<Number> byteList = (List<Number>) result;
                        byte[] pdfBytes = new byte[byteList.size()];
                        for (int b = 0; b < byteList.size(); b++) {
                            pdfBytes[b] = byteList.get(b).byteValue();
                        }

                        savedPath = Paths.get(downloadsPath, fileName);
                        Files.write(savedPath, pdfBytes);
                        System.out.println("PDF saved to: " + savedPath);

                        // Close new tab
                        pdfPage.close();

                    } catch (Exception pdfEx) {
                        System.out.println("Error reading PDF for ID " + id + ": " + pdfEx.getMessage());
                        savedPath = null;
                        if (pdfPage != null && !pdfPage.isClosed()) {
                            pdfPage.close();
                        }
                    }
                }

                // OCR the saved PDF
                if (savedPath != null) {

                    File latestPdf = savedPath.toFile();

                    if (!latestPdf.exists()) {
                        System.out.println("PDF file not found after save. Skipping ID: " + id);
                    } else {
                        try {
                            ScannedPdfReader2 sf = new ScannedPdfReader2();
                            pdfText = sf.extractTextFromScannedPdf(latestPdf.getAbsolutePath());
                            System.out.println("PDF text extracted successfully for ID: " + id);
                        } catch (Exception ocrEx) {
                            System.out.println("OCR failed for ID " + id + ": " + ocrEx.getMessage());
                        }

                        // Delete PDF after reading
                        boolean deleted = latestPdf.delete();
                        System.out.println(deleted ? "PDF deleted." : "Failed to delete PDF.");
                    }
                }

                // =====================================================
                // STEP 4: WAIT FOR PAGE TO LOAD & SELECT FILTER
                // =====================================================
                page.waitForLoadState(LoadState.DOMCONTENTLOADED);
                page.waitForSelector("select[data-t-id='history-filter']");
                page.selectOption("select[data-t-id='history-filter']", "0");
                page.waitForSelector("//button[@data-t-rel='history-dialog-link-button']");

                // =====================================================
                // STEP 5: LOAD CKG SHEET ONCE PER INVOICE (Fillo)
                // =====================================================
                List<String[]> ckgRows = loadCkgRows();

                // =====================================================
                // STEP 6: ITERATE "VIEW CHANGES" BUTTONS
                // =====================================================
                Locator viewChangesLinks = page.locator("//button[text()=' View changes ']");
                int totalLinks = viewChangesLinks.count();
                System.out.println("Total View changes links: " + totalLinks);

                for (int j = 0; j < totalLinks; j++) {

                    // Close any open modal before clicking
                    Locator modal = page.locator(".pt-modal-wrapper");
                    if (modal.isVisible()) {
                        page.keyboard().press("Escape");
                    }

                    page.waitForTimeout(2000);
                    System.out.println("Clicking View changes link: " + (j + 1));

                    Locator currentLink = page.locator("//button[text()=' View changes ']").nth(j);
                    currentLink.scrollIntoViewIfNeeded();
                    currentLink.click();

                    page.waitForTimeout(3000);

                    Locator headerDataSpan = page.locator(
                            "//span[@data-t-id='count-on-tab' and contains(text(),' Header data ')]");

                    Locator txt_Nochanges = page.locator("//div[contains(text(),'No changes')]");
                    Locator txt_coding = page.locator("//span[contains(text(),' Coding ')]");

                    HeaderDataLookup lookup = new HeaderDataLookup();

                    // Process only if Header data tab found
                    if (headerDataSpan.count() > 1) {

                        lookup.load(page);

                        for (String field : lookup.fields()) {
                            newComment = "";
                            Suppliername_new = "";
                            Suppliername_old = "";

                            boolean is_mandatory = isFieldMandatory(
                                    ConfigReader.get("dcg.docx.path"), field);

                            System.out.println("IS_MANDATORY " + is_mandatory + " | Field: " + field);

                            if (is_mandatory == true) {

                                // Get old and new values from lookup
                                Field_values = lookup.getValues(field);
                                Old_value = Field_values[0];
                                New_value = Field_values[1];
                                Suppliername_old = Old_value;
                                Suppliername_new = New_value;

                                FieldValue_Old_foundInInvoice = pdfText
                                        .toLowerCase()
                                        .contains(Suppliername_old.trim().toLowerCase());

                                Field_Name_FoundInInvoice = pdfText
                                        .toLowerCase()
                                        .contains(field.trim().toLowerCase());

                                // =====================================================
                                // WRITE RESULTS TO EXCEL (Fillo — no cell objects,
                                // no column indexes, rows addressed by id)
                                // =====================================================

                                boolean result2 = Keyword_Found_in_CKG(field, pdfText, ckgRows);

                                if (result2 == true)//it means KEYWORD IS PRESENT IN CKG
                                {
                                    System.out.println("FIELD KEYWORD FOUND IN PDF >>> " + field);

                                    if (FieldValue_Old_foundInInvoice == false)
                                    {
                                        if ("Supplier name".equalsIgnoreCase(field))
                                        {
                                            if (getSupplierCountInExtracts(Suppliername_old) > 1)
                                            {
                                                newComment = "Multiple supplier found in extracts";
                                                Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                            }
                                            else
                                            {
                                                if ("-".equals(Suppliername_old) || Suppliername_old.contains("NON"))
                                                {
                                                    if (Field_Name_FoundInInvoice == true) {
                                                        newComment = field + " not captured from invoice";
                                                        Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                    }
                                                    else {
                                                        newComment = field + " not captured from invoice";
                                                        Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                    }
                                                }
                                                else {// This cond. is that if Field name is not found in invoice
                                                    if (Field_Name_FoundInInvoice == false) {
                                                        newComment = "Wrong " + field + " captured from invoice";
                                                        Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                    }
                                                }
                                            }

                                        }	//else if field name is NOT EQUAL TO SUPPLIER / OTHER THEN SUPPLIER
                                        else
                                        {
                                            if ("-".equals(Suppliername_old) || Suppliername_old.contains("NON"))
                                            {
                                                if (Field_Name_FoundInInvoice == true) {
                                                    newComment = field + " not captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                                else {
                                                    newComment = field + " Keyword not present on invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                            }
                                            else {// This cond. is that if Field name is not found in invoice
                                                if (Field_Name_FoundInInvoice == false)
                                                {
                                                    newComment = "Wrong " + field + " captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                                else {
                                                    newComment = "Wrong " + field + " captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                            }
                                        }
                                        System.out.println(
                                                "FIELD_1 = " + field +
                                                " | OLD = " + Suppliername_old +
                                                " | NEW = " + Suppliername_new +
                                                " | COMMENT = " + newComment
                                        );
                                    }
                                    else// when there is no _vs Keyword present in invoice
                                    {
                                        System.out.println("No Keyword " + field);

                                        if (!"Supplier name".equals(field))
                                        {
                                            if ("-".equals(Suppliername_old) || Suppliername_old.contains("NON"))
                                            {
                                                if (Field_Name_FoundInInvoice == true) {
                                                    newComment = field + " not captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                            }
                                            else {// This cond. is that if Field name is not found in invoice
                                                if (Field_Name_FoundInInvoice == false) {
                                                    newComment = "Wrong " + field + " captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                            }

                                        }	//else if field name is EQUAL TO SUPPLIER / OTHER THEN SUPPLIER
                                        else
                                        {
                                            if ("-".equals(Suppliername_old) || Suppliername_old.contains("NON"))
                                            {
                                                if (Field_Name_FoundInInvoice == true) {
                                                    newComment = field + " not captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                            }
                                            else {// This cond. is that if Field name is not found in invoice
                                                if (Field_Name_FoundInInvoice == false) {
                                                    newComment = "Wrong " + field + " captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                            }
                                        }

                                        System.out.println(
                                                "FIELD_2 = " + field +
                                                " | OLD = " + Suppliername_old +
                                                " | NEW = " + Suppliername_new +
                                                " | COMMENT = " + newComment
                                        );
                                    }
                                }//closing braces for Keyword found in ckg

                                else {
                                    System.out.println("Keyword is not found in CKG");
                                    if ("Supplier name".equalsIgnoreCase(field))
                                    {
                                        if (getSupplierCountInExtracts(Suppliername_old) > 1)
                                        {
                                            newComment = "Multiple supplier found in extracts";
                                            Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                        }
                                        else
                                        {
                                            if ("-".equals(Suppliername_old) || Suppliername_old.contains("NON"))
                                            {
                                                if (Field_Name_FoundInInvoice == true) {
                                                    newComment = field + " not captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                                else {
                                                    newComment = field + " not captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                            }
                                            else {// This cond. is that if Field name is not found in invoice
                                                if (Field_Name_FoundInInvoice == false) {
                                                    newComment = "Wrong " + field + " captured from invoice";
                                                    Add_comment(centersConn, id, newComment, Suppliername_new, Suppliername_old);
                                                }
                                            }
                                        }
                                    }
                                }
                                System.out.println("NEW COMMENT " + newComment);
                            }

                            // NOTE: no explicit save needed here — Fillo writes the
                            // file on every executeUpdate inside Add_comment

                        } // end for fields loop

                        System.out.println("Workbook saved.");
                        // Close popup
                        Locator closeBtn = page.locator(
                                "//button[contains(@class,'close') or contains(text(),'Close')]");
                        if (closeBtn.count() > 0) {
                            closeBtn.first().click();
                            page.waitForTimeout(2000);
                        }

                    } // end if headerDataSpan found

                    if (txt_Nochanges.count() == 1) {
                        page.keyboard().press("Escape");
                    } else if (txt_coding.count() == 2) {
                        System.out.println("Header data NOT FOUND but Coding found");
                        page.keyboard().press("Escape");
                        page.waitForTimeout(2000);
                    }

                    // Close any remaining close button on last iteration
                    if (j == totalLinks - 1) {
                        page.keyboard().press("Escape");
                    }

                    System.out.println("???????????? rec " + (i + 1) + " | link " + (j + 1) + " ????????????");

                } // end for j loop (View changes buttons)

                System.out.println("Page Title: " + page.title());
                page.waitForTimeout(5000);

            } // end for i loop (invoice rows)

            centersConn.close();

        } catch (Exception e) {
            System.out.println("Error occurred:");
            e.printStackTrace();
        }
    }
}
