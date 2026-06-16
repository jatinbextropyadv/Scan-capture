package Tests;

import com.microsoft.playwright.Browser;
import com.microsoft.playwright.BrowserContext;
import com.microsoft.playwright.BrowserType;
import com.microsoft.playwright.Locator;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.Playwright;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Reads the "Header data" panel from the Basware "Changes made to invoice"
 * dialog and stores every row as Field -> [oldValue, newValue].
 *
 * DOM contract (from the page HTML):
 *   active panel : mat-tab-body.mat-mdc-tab-body-active
 *   row          : div.pt-grid-list-item
 *     field      : div.pt-col-2
 *     old value  : div.pt-col-3
 *     new value  : div.pt-col-4
 *
 * Typical usage (no hard-coded field names):
 *     HeaderDataLookup lookup = new HeaderDataLookup();
 *     lookup.load(page);
 *
 *     for (String field : lookup.fields()) {
 *         String[] values   = lookup.getValues(field);
 *         String   oldValue = values[0];
 *         String   newValue = values[1];
 *         // ... use oldValue / newValue
 *     }
 */
public class HeaderDataLookup {

    /** key = Field name, value = [oldValue, newValue]. Preserves on-screen order. */
    private final Map<String, String[]> headerData = new LinkedHashMap<>();

    // ----------------------------------------------------------------------
    // Load
    // ----------------------------------------------------------------------

    /** Scrape the active Header data panel on the given page. */
//    public void load(Page page) {
//        headerData.clear();
//
//        // Scope to the active tab body (Header data)
//        Locator panel = page.locator("mat-tab-body.mat-mdc-tab-body-active").first();
//        panel.waitFor();
//
//        
//        // Every data row in the grid
//        Locator rows = panel.locator("//div[@class='pt-grid-list-item list-item ng-star-inserted']/div");
//        rows.waitFor();
//        int count = rows.count(); 
//         
//        System.out.println("COUNT = "+count);
//        for (int i = 0; i < count; i++) {
//            Locator row = rows.nth(i);
//
//            String field    = textOf(row.locator("div.pt-col-2"));
//            String oldValue = textOf(row.locator("div.pt-col-3"));
//            String newValue = textOf(row.locator("div.pt-col-4"));
//
//            if (field.isEmpty()) continue;   // safety guard
//
//            headerData.put(field, new String[]{ normalize(oldValue), normalize(newValue) });
//        }
//    }

    public void load(Page page) {
        headerData.clear();

        // Scope to the Header data component (not the generic active tab body)
        Locator panel = page.locator("ia-history-header-data-changes").first();
        panel.waitFor();

        // Each row of the grid
        Locator rows = panel.locator("div.pt-grid-list-item");
        int count = rows.count();
 
        for (int i = 0; i < count; i++) {
            Locator row = rows.nth(i);

            String field    = textOf(row.locator("div.pt-col-2"));
            String oldValue = textOf(row.locator("div.pt-col-3"));
            String newValue = textOf(row.locator("div.pt-col-4"));

            if (field.isEmpty()) continue;

            headerData.put(field, new String[]{ normalize(oldValue), normalize(newValue) });
        }
    }

    private static String textOf(Locator loc) {
        if (loc.count() == 0) return "";
        String t = loc.first().innerText();
        return t == null ? "" : t.replaceAll("\\s+", " ").trim();
    }

    private static String normalize(String v) {
        return (v == null) ? "" : v;
    }
    // ----------------------------------------------------------------------
    // Fields column (read from the page — no hard-coding required)
    // ----------------------------------------------------------------------

    /** All Field-column values, in the order they appeared on screen. */
    public List<String> fields() {
        return new ArrayList<>(headerData.keySet());
    }

    /** Same data as a Set if you only need membership checks. */
    public Set<String> fieldSet() {
        return Collections.unmodifiableSet(headerData.keySet());
    }

    // ----------------------------------------------------------------------
    // The single accessor — pass a Field, get Old & New back
    // ----------------------------------------------------------------------

    /**
     * Returns the row values for the given Field as [oldValue, newValue].
     * Returns null if the field is not present.
     *
     * Example: 
     *     String[] v = lookup.getValues("Net total");
     *     String oldValue = v[0];
     *     String newValue = v[1];
     */
    public String[] getValues(String field) {
        return headerData.get(field);
    }

    // Convenience wrappers (optional)
//    public String oldValueOf(String field) {//change made by Jatin from string field to String[]
    public String oldValueOf(String values) {//change made by Jatin
        String[] v = headerData.get(values);
        return v == null ? null : v[0];
    }
//    public String newValueOf(String field) {//change made here by JaTin Bakshi
    public String newValueOf(String values) {//change made here by JaTin Bakshi
        String[] v = headerData.get(values);
        return v == null ? null : v[1];
    }

    public boolean contains(String field) { return headerData.containsKey(field); }
    public int     size()                 { return headerData.size(); }

    /** Read-only view of the underlying map. */
    public Map<String, String[]> all() {
        return Collections.unmodifiableMap(headerData);
    }

    // ----------------------------------------------------------------------
    // Demo
    // ----------------------------------------------------------------------

    public static void main(String[] args) throws InterruptedException {
        try (Playwright pw = Playwright.create()) {
        	
        	// Connect to already opened Chrome browser
            Browser browser = pw.chromium()
            		.connectOverCDP("http://127.0.0.1:9224"); 

            System.out.println("Connected to existing Chrome browser");

            
            
            Thread.sleep(5000);
            // Get existing browser context
            BrowserContext context = browser.contexts().get(0);

            Page page;
        	
            // Use existing tab if available
            if (context.pages().size() > 0) {
                page = context.pages().get(0);
            } else {
                page = context.newPage();
            }
        	
        	
        	
//            Browser browser = pw.chromium().launch(
//                    new BrowserType.LaunchOptions().setHeadless(false));
//            Page page = browser.newContext().newPage();
//
//            page.navigate("https://cenbu.p2p.basware.com/ap/invoice/details"
//                    + "?docId=407c4504a8a9429da3e2cc48cfb2fc55");

            // (Open the "Changes made to invoice" dialog before this call.)
            HeaderDataLookup lookup = new HeaderDataLookup();
            lookup.load(page); 

            // ---------------------------------------------------------------
            // Fetch the Fields column FROM THE CODE (no hard-coded names),
            // then pass each field to getValues() to retrieve Old / New.
            // ---------------------------------------------------------------
            for (String field : lookup.fields()) {
                String[] values   = lookup.getValues(field);
                String   oldValue = values[0];
                String   newValue = values[1];

                System.out.println(field
                        + " | old='" + oldValue + "'"
                        + " | new='" + newValue + "'");
            }

            //browser.close();
        }
    }
}