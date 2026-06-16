package Tests;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExtentManager {

    private static ExtentReports extent;
    private static String reportPath;

    public static synchronized ExtentReports getInstance() {
        if (extent == null) {
            String timestamp = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss").format(new Date());

            String reportDir = System.getProperty("user.dir") + File.separator + "reports";
            new File(reportDir).mkdirs();

            reportPath = reportDir + File.separator + "ExtentReport_" + timestamp + ".html";

            ExtentSparkReporter spark = new ExtentSparkReporter(reportPath);
            spark.config().setTheme(Theme.STANDARD);
            spark.config().setDocumentTitle("Invoice Validation Report");
            spark.config().setReportName("Scan And Capture");
            spark.config().setTimeStampFormat("dd-MM-yyyy HH:mm:ss");
            spark.config().setCss(".test-name { font-weight: bold !important; }"); // bold fix

            extent = new ExtentReports();
            extent.attachReporter(spark);

            extent.setSystemInfo("OS", System.getProperty("os.name"));
            extent.setSystemInfo("Java Version", System.getProperty("java.version"));
            extent.setSystemInfo("User", System.getProperty("user.name"));

            // ✅ Flush report on JVM shutdown (covers force-stop, crashes, normal exit)
            Runtime.getRuntime().addShutdownHook(new Thread(() -> {
                if (extent != null) {
                    extent.flush();
                    System.out.println("✅ Extent Report flushed via shutdown hook: " + reportPath);
                }
            }));
        }
        return extent;
    }

    public static String getReportPath() {
        return reportPath;
    }
}