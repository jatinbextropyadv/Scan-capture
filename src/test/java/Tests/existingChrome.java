package Tests;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.LoadState;
import com.microsoft.playwright.options.WaitForSelectorState;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.yaml.snakeyaml.tokens.Token;

import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class existingChrome {

    // Excel file path 
    String excelPath = ConfigReader.get("centers.excel.path");

    // Base URL
    String baseUrl = ConfigReader.get("base.url");

    String Suppliername_old = "";
   // String Suppliername_new = "";
    String New_value = "";
    String Old_value = "";

    static boolean count_print = true;
    static boolean counter_2 = true;

    public int idColumnIndex = -1;
    public int expectedValueColumnIndex = -1;
    public int actualValueColumnIndex = -1;
    int commentsColumnIndex = -1;

    public static String existingComment="";
    public static String existingExpected_value="";
    public static String existingActual_value="";
    
    public static String newComment="";
    public static String Suppliername_new="";
    
    public static Cell commentsCell;
    
    public static Cell expected_valueCell;
    public static Cell actual_valueCell;
    public static boolean Is_OCR=true;
    public static boolean Field_Value_Exist_In_Pdf = false;

    // =====================================================
    // SHARED STATE (opened once in @BeforeClass, used by every @Test invocation)
    // =====================================================
    Playwright playwright;
    Browser browser;
    BrowserContext context;
    Page page;

    Workbook workbook;
    Sheet sheet;
    Row headerRow;
    // =====================================================
    // HELPER METHODS
    // =====================================================
    

    /** True if the PDF had no usable text layer and OCR was used instead. */
 // just construct it and call the public methods — nothing else


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
    
    public static void Add_comment(String newComment,String existingComment,String existingExpected_value,String existingActual_value, Cell commentsCell, Cell expected_valueCell, Cell actual_valueCell, String Suppliername_new, String Suppliername_old)
	{
            // No finding for this field — don't write anything, otherwise the
            // Expected/Actual cells get values with no matching comment
//            if (newComment == null || newComment.trim().isEmpty()) {
//                return;
//            }

            // Empty old value is recorded as "Blank" so the cell is never empty
            String actualValue = (Suppliername_old == null || Suppliername_old.trim().isEmpty())
                    ? "Blank"
                    : Suppliername_old.trim();

            if (!existingComment.isEmpty() && !existingComment.equals(newComment)) {
                commentsCell.setCellValue(existingComment + " ," + newComment); 
            } else if (existingComment.isEmpty()) {
                commentsCell.setCellValue(newComment);
            }

            if (!existingExpected_value.isEmpty() && !existingExpected_value.equals(Suppliername_new)) {
                expected_valueCell.setCellValue(existingExpected_value + " ," + Suppliername_new);
            } else if (existingExpected_value.isEmpty()) {
                expected_valueCell.setCellValue(Suppliername_new);
            }

            if (!existingActual_value.isEmpty() && !existingActual_value.equals(actualValue)) {
                actual_valueCell.setCellValue(existingActual_value + " ," + actualValue);
            } else if (existingActual_value.isEmpty()) {
                actual_valueCell.setCellValue(actualValue);
            }
	}

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

                        for (int i = 0; i < cells.size(); i++) {
                            //System.out.println("Cell " + (i + 1) + " = [" + cells.get(i).getText() + "]");
                        }

                        if (hasCheckbox == true && checkboxValue == true) {
//                            System.out.println(
//                                    "fieldFound=" + fieldFound +
//                                    ", hasCheckbox=" + hasCheckbox +
//                                    ", checkboxValue=" + checkboxValue);
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

    private static String Keyword_Found_in_CKG(
            String searchText,
            String pdfText,
            Sheet sheet,
            DataFormatter formatter, String Suppliername_old) throws Exception {
    	

        int startRow = -1;
        String KeywordFoundinCKG="";
       // Find partial text in Column B
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);

            if (row == null) {
                continue;
            }

            String colB = formatter.formatCellValue(row.getCell(1)).trim();

            
            
            if (!colB.isEmpty() && colB.equalsIgnoreCase(searchText)) {
                startRow = i+1;// DELIBERATELY DO THIS plus ONE +1 EARLIER IT WASI ONLY
               System.out.println("STARTROW "+ startRow);
               System.out.println("colB "+colB);
               break;
            }
        }

        if (startRow == -1) {
//            System.out.println("Section not found : " + searchText);
//            return false;
            //return "Not Found";
        	return "Not Found";
        }

       // boolean keywordFound = false;

        // Iterate until blank row
        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);

            if (row == null) {
                break;
            }
            
            //Here values in columns C,D,E ARE stored in variables of type String
            
            String c = formatter.formatCellValue(row.getCell(2)).trim();
            String d = formatter.formatCellValue(row.getCell(3)).trim();
            String e = formatter.formatCellValue(row.getCell(4)).trim();

            // Stop when C, D, E all blank
            if (c.isEmpty() && d.isEmpty() && e.isEmpty()) {
                break;
            }

            String[] keywords = {c, d, e};

            for (String keyword : keywords) 
            {

                if (keyword.isEmpty()) {
                    continue;
                }
               
                String str = pdfText.toLowerCase();
                
                String keyword_l = keyword.toLowerCase();

                if (str.contains(keyword_l)) 
                {
                    System.out.println("MATCH FOUND => " + keyword_l);
                    KeywordFoundinCKG = keyword_l;
                    return KeywordFoundinCKG;
                }    
            }
        }
		return KeywordFoundinCKG="Not Found"; 
    }

    public static boolean compareFieldWithCKG(String popupField, String pdfText) throws Exception {

        String ckgPath = ConfigReader.get("ckg.excel.path");

        try (FileInputStream fis = new FileInputStream(ckgPath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();

            boolean fieldMatched = false;
            boolean allKeywordsMatched = true;

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {

                Row row = sheet.getRow(i);

                if (row == null) {
                    continue;
                }

                String fieldName = formatter.formatCellValue(row.getCell(1)).trim();

                if (!fieldMatched) {

                    if (!fieldName.isEmpty() && fieldName.equalsIgnoreCase(popupField.trim())) {

                        fieldMatched = true;
                        System.out.println("Matched Field >>> " + fieldName);
                        allKeywordsMatched = processKeywords(row, formatter, pdfText);
                    }
                } else {

                    if (!fieldName.isEmpty()) {
                        System.out.println("Next Field Encountered >>> " + fieldName);
                        break;
                    }

                    if (!processKeywords(row, formatter, pdfText)) {
                        allKeywordsMatched = false;
                    }
                }
            }

            return fieldMatched && allKeywordsMatched;
        }
    }

    private static boolean processKeywords(
            Row row,
            DataFormatter formatter,
            String pdfText) {

        boolean allKeywordsFound = true;

        // Columns C, D, E
        for (int col = 2; col <= 4; col++) {

            String keyword = formatter.formatCellValue(row.getCell(col)).trim();

            if (keyword.isEmpty()) {
                continue;
            }

           // System.out.println("Checking Keyword >>> " + keyword);

            if (pdfText.toLowerCase().contains(keyword.toLowerCase())) {
               // System.out.println("Keyword FOUND in PDF >>> " + keyword);
            } else {
              //  System.out.println("Keyword NOT FOUND in PDF >>> " + keyword);
                allKeywordsFound = false;
            }
        }

        return allKeywordsFound;
    }

    public static void checkKeywordInPdf(String keyword, String pdfText) {

        if (keyword == null || keyword.trim().isEmpty())
            return;

        if (pdfText.toLowerCase().contains(keyword.toLowerCase())) {
           // System.out.println("Keyword [" + keyword + "] present in PDF");
        } else {
          //  System.out.println("Keyword [" + keyword + "] NOT present in PDF");
        }
    }

    public static int getSupplierCountInExtracts(String supplierName) throws Exception {

        int count = 0;

        try (FileInputStream fis = new FileInputStream(ConfigReader.get("extracts.excel.path"));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();

            Row headerRow = sheet.getRow(0);
            int supplierColumnIndex = -1;

            for (Cell headerCell : headerRow) {

                String header = formatter.formatCellValue(headerCell).trim();

                if ("Supplier_Name".equalsIgnoreCase(header)) {
                    supplierColumnIndex = headerCell.getColumnIndex();
                    break;
                }
            }

            if (supplierColumnIndex == -1) {
                throw new RuntimeException("Supplier_Name column not found");
            }

            String searchSupplier = supplierName == null
                    ? ""
                    : supplierName.replace("\r", " ")
                                  .replace("\n", " ")
                                  .replaceAll("\\s+", " ")
                                  .trim();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {

                Row row = sheet.getRow(i);

                if (row == null) continue;

                Cell cell = row.getCell(supplierColumnIndex);

                if (cell == null) continue;

                String excelSupplier = formatter
                        .formatCellValue(cell)
                        .replaceAll("\\s+", " ")
                        .trim();

                if (excelSupplier.equalsIgnoreCase(searchSupplier)) {
                    count++;
                }
            }
        }
        return count;
    }

    // =====================================================
    // TESTNG LIFECYCLE
    // =====================================================

    /**
     * Opens the Excel workbook (kept in memory for the whole run), resolves the
     * required column indexes, and connects to the already-running Chrome.
     */
    @BeforeClass
    public void setUp() throws Exception {

        playwright = Playwright.create();

        FileInputStream fis = new FileInputStream(excelPath);
        workbook = new XSSFWorkbook(fis);
        fis.close(); // XSSFWorkbook is fully in-memory, so we can save back to the same path later

        // Connect to already opened Chrome browser
        browser = playwright.chromium().connectOverCDP("http://127.0.0.1:9224");
        System.out.println("Connected to existing Chrome browser");

        Thread.sleep(5000);

        // Get existing browser context
        context = browser.contexts().get(0);

        if (context.pages().size() > 0) {
            page = context.pages().get(0);
        } else {
            page = context.newPage();
        }

        // Read first sheet
        sheet = workbook.getSheetAt(0);
        headerRow = sheet.getRow(0);

        idColumnIndex = -1;
        expectedValueColumnIndex = -1;
        actualValueColumnIndex = -1;
        commentsColumnIndex = -1;

        // Find required columns
        try {
            for (Cell cell : headerRow) {

                String columnName = cell.getStringCellValue().trim();

                if (columnName.equalsIgnoreCase("id")) {
                    idColumnIndex = cell.getColumnIndex();
                }
                if (columnName.equalsIgnoreCase("Comments")) {
                    commentsColumnIndex = cell.getColumnIndex();
                }
                if (columnName.equalsIgnoreCase("Expected value")) {
                    expectedValueColumnIndex = cell.getColumnIndex();
                }
                if (columnName.equalsIgnoreCase("Actual value")) {
                    actualValueColumnIndex = cell.getColumnIndex();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Validate id column
        if (idColumnIndex == -1) {
            throw new RuntimeException("Column 'id' not found in Excel");
        }

        // Create columns if not exist
        if (commentsColumnIndex == -1) {
            commentsColumnIndex = headerRow.getLastCellNum();
            headerRow.createCell(commentsColumnIndex).setCellValue("Comments");
        }
        if (expectedValueColumnIndex == -1) {
            expectedValueColumnIndex = headerRow.getLastCellNum();
            headerRow.createCell(expectedValueColumnIndex).setCellValue("Expected value");
        }
        if (actualValueColumnIndex == -1) {
            actualValueColumnIndex = headerRow.getLastCellNum();
            headerRow.createCell(actualValueColumnIndex).setCellValue("Actual value");
        }
    }

    /**
     * Reads every non-empty id from the Excel sheet and supplies them one-by-one
     * to the test method. Opens its own stream so it does not depend on @BeforeClass
     * ordering.
     */
    @DataProvider(name = "invoiceIds")
    public Object[][] invoiceIds() throws Exception {

        List<Object[]> ids = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(ConfigReader.get("centers.excel.path"));
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sh = wb.getSheetAt(0);
            Row header = sh.getRow(0);

            int idCol = -1;
            for (Cell cell : header) {
                if (cell.getStringCellValue().trim().equalsIgnoreCase("id")) {
                    idCol = cell.getColumnIndex();
                    break;
                }
            }

            if (idCol == -1) {
                throw new RuntimeException("Column 'id' not found in Excel");
            }

            for (int i = 1; i <= sh.getLastRowNum(); i++) {
                Row row = sh.getRow(i);
                if (row == null) continue;

                Cell cell = row.getCell(idCol);
                if (cell == null) continue;

                String id = cell.toString().trim();
                if (id.isEmpty()) continue;

                ids.add(new Object[] { id });
            }
        }

        return ids.toArray(new Object[0][]);
    }

    /** Finds the sheet row whose id column matches the given id. */
    private Row findRowById(String id) {
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell cell = row.getCell(idColumnIndex);
            if (cell == null) continue;

            if (cell.toString().trim().equals(id)) {
                return row;
            }
        }
        return null;
    }

    /**
     * Saves the workbook and closes the browser / Playwright once all test
     * invocations have finished.
     */
    @AfterClass(alwaysRun = true)
    public void tearDown() throws Exception {

        if (workbook != null) {
            try (FileOutputStream fos = new FileOutputStream(excelPath)) {
                workbook.write(fos);
            }
            workbook.close();
            System.out.println("Workbook saved and closed.");
        }

        if (browser != null) {
            browser.close();
        }
        if (playwright != null) {
            playwright.close();
        }
    }

    // =====================================================
    // TEST METHOD — runs once per id from the Excel sheet
    // =====================================================

    @Test(dataProvider = "invoiceIds")
    public void Scan_And_Capture(String id) throws Exception {

        String baseUrl = ConfigReader.get("base.url");

        String Suppliername_old = "";
        String Suppliername_new = "";
        String New_value = "";
        String Old_value = "";

        String[] Field_values = {""};

        boolean FieldValue_Old_foundInInvoice = false;
        boolean Field_Name_FoundInInvoice = false;
        String pdfText = "";
        boolean Is_keyword_Found=false;
        String KeywordFound="";

        File latestPdf=null;

        try {

            // Locate this id's row so results can be written back to it
            Row row = findRowById(id);
            if (row == null) {
                System.out.println("Row not found for id: " + id);
                return;
            }

            Locator nextButton = page.locator("//div[@class='pt-btn-group pt-next-prev-page-btn-group']/button[@title='Next']");

            if (nextButton.isVisible() && nextButton.isEnabled()) {
                System.out.println("Next button is clickable");
            }

            {
                // Clear results from previous runs so Comments / Expected /
                // Actual stay aligned and don't accumulate stale entries
                Cell oldResultCell = row.getCell(commentsColumnIndex);
                if (oldResultCell != null) oldResultCell.setCellValue("");
                oldResultCell = row.getCell(expectedValueColumnIndex);
                if (oldResultCell != null) oldResultCell.setCellValue("");
                oldResultCell = row.getCell(actualValueColumnIndex);
                if (oldResultCell != null) oldResultCell.setCellValue("");

                String finalUrl = baseUrl + id;
                System.out.println("\nOpening URL: " + finalUrl);

                // Navigate to invoice page
                page.navigate(finalUrl);  //HIDE this line If i am running thru PAGINATION in AP 
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
                                //System.out.println("Trying iframe selector: " + iframeSel);
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
                       // System.out.println("Trying download button directly on page (no iframe)...");
                        Locator candidate = page.locator("button[data-block-id='download-button']");
                        candidate.waitFor(new Locator.WaitForOptions()
                                .setState(WaitForSelectorState.VISIBLE)
                                .setTimeout(15000));
                        downloadBtn = candidate;
                        //System.out.println("Download button found on page.");
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
                	System.out.println("SAVED pATH  >>> "+savedPath);
                
                
                // OCR the saved PDF
                PdfKeywordValueFinder finder = null;
                if (savedPath != null) {
                	//File
                    latestPdf = savedPath.toFile();

                    System.out.println("latestPdf>>>>>>>>>>>>>>>> "+latestPdf);
                    
                    if (!latestPdf.exists()) 
                    {
                        System.out.println("PDF file not found after save. Skipping ID: " + id);
                    } else 
                    {
                        finder = new PdfKeywordValueFinder(latestPdf);

                        if (finder.usedOcr()) {
                            System.out.println("Scanned PDF — values came from OCR");
                            Is_OCR=true;
                        } else {
                            System.out.println("Text-based PDF — exact values");
                            Is_OCR=false;
                        }
                    }
                    
                    System.out.println("Is_OCR "+Is_OCR);
                    
                    if (!latestPdf.exists()) {
                        System.out.println("PDF file not found after save. Skipping ID: " + id);
                    } else {
                        try {
                            ScannedPdfReader2 sf = new ScannedPdfReader2();
                            pdfText = sf.extractTextFromScannedPdf(latestPdf.getAbsolutePath());
                            
                            System.out.println("PDF text extracted successfully for ID: " + id);
                            //System.out.println("PDF EXTRACTED ????  "+pdfText);
                            //System.out.println(" -- ");
                        }catch (Exception ocrEx) {
                            System.out.println("OCR failed for ID " + id + ": " + ocrEx.getMessage());
                        }

                        // Delete PDF after reading
                        boolean deleted = latestPdf.delete();
                        System.out.println(deleted ? "PDF deleted." : "Failed to delete PDF.");
                    }
                }
                
                //System.out.println("PDF TEXT ++++++++++++++++++++++++++++++ "+pdfText +" ++++++++++++++++++++++++++++++++++++++++++");
                // =====================================================
                // STEP 4: WAIT FOR PAGE TO LOAD & SELECT FILTER
                // =====================================================
                
                page.waitForLoadState(LoadState.DOMCONTENTLOADED);
                page.waitForSelector("select[data-t-id='history-filter']");
                page.selectOption("select[data-t-id='history-filter']", "0");
                page.waitForSelector("//button[@data-t-rel='history-dialog-link-button']");  //REMOVED this line for running code using PAGINATION

                // =====================================================
                // STEP 5: OPEN CKG WORKBOOK ONCE PER INVOICE
                // =====================================================
                FileInputStream ckgFis = new FileInputStream(
                        ConfigReader.get("ckg.excel.path"));
                Workbook ckgWorkbook = new XSSFWorkbook(ckgFis);
                Sheet ckgSheet = ckgWorkbook.getSheet("English");
                DataFormatter formatter = new DataFormatter();

                // =====================================================
                // STEP 6: ITERATE "VIEW CHANGES" BUTTONS
                // =====================================================
                Locator viewChangesLinks = page.locator("//button[text()=' View changes ']");
                int totalLinks = viewChangesLinks.count();
                System.out.println("Total View changes links: " + totalLinks);

                if(totalLinks>0)
                {
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

                        for (String field : lookup.fields()) 
                        {
                        	newComment = "";
                        	Suppliername_new="";
                        	Suppliername_old="";
                        	
                            boolean is_mandatory = isFieldMandatory(
                                    ConfigReader.get("dcg.docx.path"), field);

                            System.out.println("IS_MANDATORY " + is_mandatory + " | Field: " + field);

                          if (is_mandatory == true) {
                        	//check in DCG that Keyword/field is mandatory or NOT
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

                               // System.out.println("Field text/name Found In PDF: " + Field_Name_FoundInInvoice + " | " + field);
                                

                             // =====================================================
                             // FIND COLUMN INDEXES FROM HEADER ROW
                             // =====================================================
                             for (Cell headerCell : headerRow) {
                                 String col = headerCell.getStringCellValue().trim();
                                 if (col.equalsIgnoreCase("Comments"))
                                     commentsColumnIndex = headerCell.getColumnIndex();
                                 if (col.equalsIgnoreCase("Expected value"))
                                     expectedValueColumnIndex = headerCell.getColumnIndex();
                                 if (col.equalsIgnoreCase("Actual value"))
                                     actualValueColumnIndex = headerCell.getColumnIndex();
                             }

                             if (commentsColumnIndex == -1) {
                                 commentsColumnIndex = headerRow.getLastCellNum();
                                 headerRow.createCell(commentsColumnIndex).setCellValue("Comments");
                             }

                             // Get or create cells
                             commentsCell = row.getCell(commentsColumnIndex);
                             if (commentsCell == null)
                                 commentsCell = row.createCell(commentsColumnIndex);

                             expected_valueCell = row.getCell(expectedValueColumnIndex);
                             if (expected_valueCell== null)
                                 expected_valueCell = row.createCell(expectedValueColumnIndex);

                             actual_valueCell = row.getCell(actualValueColumnIndex);
                             if (actual_valueCell == null)
                                 actual_valueCell = row.createCell(actualValueColumnIndex);

                             // Read existing cell values
                             commentsCell.setCellType(CellType.STRING);
                             existingComment = commentsCell.getStringCellValue().trim();

                             expected_valueCell.setCellType(CellType.STRING);
                             existingExpected_value = expected_valueCell.getStringCellValue().trim();

                             actual_valueCell.setCellType(CellType.STRING);
                             existingActual_value = actual_valueCell.getStringCellValue().trim();
                             
                             
                            // =====================================================
                            // WRITE RESULTS TO EXCEL
                            // =====================================================
                            //System.out.println("Field Value_Old_foundInInvoice >>>> "+ Suppliername_old +" "+ FieldValue_Old_foundInInvoice);

                             KeywordFound = Keyword_Found_in_CKG(field, pdfText, ckgSheet, formatter, Suppliername_old);
                             
                             System.out.println("KeywordFound in INvoice "+KeywordFound);
                             
                             if(!"Not Found".equalsIgnoreCase(KeywordFound)) 
	                             {
	                            	 Is_keyword_Found=true;
	                             }
	                        	 else 
	                        	 {
	                        		 Is_keyword_Found=false;
	                        	 } 
                             
	                         if (Is_keyword_Found == true) //It means KEYWORD IS PRESENT IN CKGas well as in INVOICE
	                         {
	                        	 //Value of field is compared here..

	                        	 if (finder != null) 
                        	 		{
	                        		 	System.out.println("It finds the VALUE corresponds to Keyword found in CKG "+ KeywordFound);
	                        		 	
	                                    String right_val = finder.findValueRightOf(KeywordFound);
//	                                    String below_val = finder.findValueBelow(KeywordFound);
	                                    String below_val = finder.findValueInColumnBelow(KeywordFound);  // → "75.00"
	                                    String near_val = finder.findValueNear(KeywordFound);
	                                    System.out.println("right_val === "+right_val);
	                                    System.out.println("below_val === "+below_val);
	                                    System.out.println("near_val === "+near_val);
	                                    
	                                    if(right_val == Suppliername_old || below_val==Suppliername_old || near_val==Suppliername_old) 
	                                    {
	                                    	System.out.println("It finds the VALUE corresponds to Keyword found in CKG "+ KeywordFound);
	                                    	//this variable finds/returns the field value in invoice in right,Bottom,near 
	                                    	//by the field keyword
	                                    	System.out.println("RIGHT SIDE !!! "+ right_val);
	                                    	System.out.println("Below SIDE !!! "+ below_val);
	                                    	System.out.println("Near SIDE !!! "+ near_val); 
	                                    	
	                                    	FieldValue_Old_foundInInvoice = true;
	                                    }
                        	 		}    
	                        	 
                                 		System.out.println("FIELD KEYWORD FOUND IN PDF >>> " + field);
                             
                                 if (FieldValue_Old_foundInInvoice == false) 
	                                {                                
	                                	if("Supplier name".equalsIgnoreCase(field))
	                                	{
	                                		if (getSupplierCountInExtracts(Suppliername_old) > 1) 
	                                		{
	                                			newComment = "Multiple supplier found in extracts";
	                                			Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);
	                                		}
	                                		else 
	                                		{
	                                    		if ("-".equals(Suppliername_old) || Suppliername_old.contains("NON")) 
	                                    		{   
	                                                    newComment = field + " not captured from invoice";
	                                                    Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);
	                                            }
	                                    		else {
	                                                    newComment = "Wrong "+ field + " captured from invoice";
	                                                    Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);
	                                    			} 
	                                    	}
	                                	}	//else if field name is NOT EQUAL TO SUPPLIER / OTHER THEN SUPPLIER
	                                	else 
		                                	{
		                                		if ("-".equals(Suppliername_old) || Suppliername_old.contains("NON")) 
		                                		{
		                                            //how do u know that there is supplier in invoice
		                                			
		                                                newComment = field + " not captured from invoice";
		                                                Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);
		                                         }
		                                		else {// This cond. is that if Field name is not found in invoice
			                                		
			                                				newComment = "Wrong "+field +" captured from invoice";
			                                                Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);		
		                                			}
		                            		}
		                                	System.out.println(
		                                		    "FIELD_name = " + field +
		                                		    " | OLD = " + Suppliername_old +
		                                		    " | NEW = " + Suppliername_new +
		                                		    " | COMMENT = " + newComment
		                                		);
	                            	}
                                else 
                                {
                                	if("Supplier name".equals(field))
                                	{
                                		if ("-".equals(Suppliername_old) || Suppliername_old.contains("NON"))  
                                    		{
                                                    newComment = field + " not captured from invoice";
                                                    Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);    
                                            }
                                    		else {// This cond. is that if Field name is not found in invoice

                                                    newComment = "Wrong "+ field + " captured from invoice";
                                                    Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);
                                    		}

                                	}	//else if field name is EQUAL TO SUPPLIER / OTHER THEN SUPPLIER
                                	else 
                                	{
                                		if ("-".equals(Suppliername_old) || Suppliername_old.contains("NON"))  
                                		{
                                                newComment = field + " not captured from invoice";
                                                Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);
                                             
                                        }
                                		else {// This cond. is that if Field name is not found in invoice
                                			
                                                newComment = "Wrong "+ field + " captured from invoice";
                                                Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);
                                            } 
                                		
                                	}
                                	
                                	System.out.println(
                                		    "FIELD_2 = " + field +
                                		    " | OLD = " + Suppliername_old +
                                		    " | NEW = " + Suppliername_new +
                                		    " | COMMENT = " + newComment
                                		);
                                }

	                         }//closing braces for Keyword found in CKG
	                         
	                         else 
	                         {	 System.out.println("FIELDDDDD   "+field);

	                        	 if (!"Supplier name".equals(field)) {//NOT operator is used here because Supplier name keyword has is never present on invoice
		                        	 newComment="There is no +ve Keyword found in invoice against "+field;
		                        	 Add_comment(newComment,existingComment,existingExpected_value,existingActual_value,commentsCell,expected_valueCell,actual_valueCell, Suppliername_new,Suppliername_old);
	                        	 }
	                         }
                                		System.out.println("NEW COMMENT " + newComment);
                        	 }
                        
                            // Save updated Excel after each header data popup
                          FileOutputStream fos = new FileOutputStream(excelPath);
                          workbook.write(fos);
                          fos.close();                        
                        
                        } // end for fields loop

                        // Save updated Excel after each header data popup
//                        FileOutputStream fos = new FileOutputStream(excelPath);
//                        workbook.write(fos);
                      //  fos.close();
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

                    System.out.println("???????????? id " + id + " | link " + (j + 1) + " ????????????");


                }

                } // end for j loop (View changes buttons)
                else
                {
                    if (nextButton.isVisible() && nextButton.isEnabled()) {
                        System.out.println("Next button is clickable");
                        nextButton.click();
                    }
                }
                // Close CKG workbook after all j iterations for this invoice
                ckgWorkbook.close();
                ckgFis.close();

                System.out.println("Page Title: " + page.title());
                page.waitForTimeout(5000);

//                if (nextButton.isVisible() && nextButton.isEnabled()) {
//                    System.out.println("Next button is clickable");
//                }
//
            } // end of per-id block

        } catch (Exception e) {
            System.out.println("Error occurred:");
            e.printStackTrace();
        }
    }
}