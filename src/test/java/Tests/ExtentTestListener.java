package Tests;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;

import org.testng.ITestContext;
import org.testng.ITestListener;
import org.testng.ITestResult;

public class ExtentTestListener implements ITestListener {

    private static final ExtentReports extent = ExtentManager.getInstance();
    private static final ThreadLocal<ExtentTest> currentTest = new ThreadLocal<ExtentTest>();

    // Lets test classes log custom steps: ExtentTestListener.getTest().info("...")
    public static ExtentTest getTest() {
        return currentTest.get();
    }

//    @Override
//    public void onTestStart(ITestResult result) {
//        // Append the DataProvider id so each invoice gets its own node
//        // (e.g. "Scan_And_Capture [id=12345]") instead of all sharing one name.
//        String name = result.getMethod().getMethodName();
//        Object[] params = result.getParameters();
//        if (params != null && params.length > 0 && params[0] != null) {
//            name += " [id=" + params[0] + "]";
//        }
//
//        // Wrap in <b> so the entry shows in bold in the report's test list
//        ExtentTest test = extent.createTest("<b>" + name + "</b>", result.getMethod().getDescription());
//        currentTest.set(test);
//    }
    
    
    @Override
    public void onTestStart(ITestResult result) {
        String name = result.getMethod().getMethodName();
        Object[] params = result.getParameters();
        if (params != null && params.length > 0 && params[0] != null) {
            name += " [id=" + params[0] + "]";
        }
        // ✅ CHANGE THIS LINE — remove the <b> tags
        ExtentTest test = extent.createTest(name, result.getMethod().getDescription());  // ← AFTER
        currentTest.set(test);
    }

    @Override
    public void onTestSuccess(ITestResult result) {
        currentTest.get().log(Status.PASS, "Test passed");
    }

    @Override
    public void onTestFailure(ITestResult result) {
        currentTest.get().log(Status.FAIL, "Test failed");
        currentTest.get().fail(result.getThrowable());
    }  

//    @Override
//    public void onTestSkipped(ITestResult result) {
//        ExtentTest test = currentTest.get();
//        if (test == null) {
//            test = extent.createTest("<b>"+result.getMethod().getMethodName()+"</b>");
//            currentTest.set(test);
//        }
//        test.log(Status.SKIP, "Test skipped");
//        if (result.getThrowable() != null) {
//            test.skip(result.getThrowable());
//        }
//    }

    @Override
    public void onTestSkipped(ITestResult result) {
        ExtentTest test = currentTest.get();
        if (test == null) {
            // ✅ CHANGE THIS LINE — remove the <b> tags
            test = extent.createTest(result.getMethod().getMethodName());  // ← AFTER
            currentTest.set(test);
        }
        test.log(Status.SKIP, "Test skipped");
        if (result.getThrowable() != null) {
            test.skip(result.getThrowable());
        }
    }
    
    @Override
    public void onFinish(ITestContext context) {
        extent.flush();
        System.out.println("Extent Report generated at: " + ExtentManager.getReportPath());
    }
}