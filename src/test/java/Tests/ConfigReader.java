package Tests;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Properties;

public class ConfigReader {

    private static final Properties props = new Properties();

    static {
        try {
            // Try classpath first (src/test/resources is on the test classpath)
            InputStream is = ConfigReader.class.getClassLoader()
                    .getResourceAsStream("config.properties");

            if (is != null) {
                try (InputStream in = is) {
                    props.load(in);
                }
            } else {
                // Fallback: load directly from the project folder
                try (FileInputStream fis = new FileInputStream("src/test/resources/config.properties")) {
                    props.load(fis);
                }
            }
        } catch (Exception e) {
            throw new RuntimeException("Could not load config.properties", e);
        }
    }

    public static String get(String key) {
        String value = props.getProperty(key);
        if (value == null || value.trim().isEmpty()) {
            throw new RuntimeException("Property '" + key + "' is missing in config.properties");
        }
        return value.trim();
    }
//    
    
    
    public static String get2(String key) {
        // 1. JVM system property: -Dcenters.excel.path=...
        String value = System.getProperty(key);

        // 2. Environment variable: CENTERS_EXCEL_PATH=...
        if (value == null || value.trim().isEmpty()) {
            String envKey = key.toUpperCase().replace('.', '_');
            value = System.getenv(envKey);
        }

        // 3. Fallback: config.properties
        if (value == null || value.trim().isEmpty()) {
            value = props.getProperty(key);
        }

        if (value == null || value.trim().isEmpty()) {
            throw new RuntimeException("Property '" + key + "' not set (system property, env var, or config.properties)");
        }
        return value.trim();
    }
}
