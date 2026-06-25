package Tests;

import javax.crypto.Cipher;
import javax.crypto.spec.GCMParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.util.Base64;

/**
 * AES-GCM string encryption for storing secrets (e.g. a password) as ciphertext
 * instead of plain text in config/code.
 *
 * SECURITY MODEL
 * --------------
 *   - The ENCRYPTED value may live in the source / config.properties.
 *   - The KEY must NOT. It is read from the environment variable APP_SECRET_KEY.
 *     Anyone with the code but WITHOUT the key cannot recover the password.
 *
 * One-time setup:
 *   1. Choose a strong passphrase and set it as an env var on every machine that
 *      runs the tests (and as a GitHub Actions secret):
 *          Windows:  setx APP_SECRET_KEY "your-long-random-passphrase"
 *   2. Run this class's main() once with your real password to print the
 *      ciphertext, then paste that ciphertext into config.properties:
 *          app.password.enc=<printed value>
 *   3. In the test, decrypt at runtime:
 *          String pwd = Crypto.decrypt(ConfigReader.get("app.password.enc"));
 */
public final class Crypto {

    private static final String ALGO = "AES/GCM/NoPadding";
    private static final int IV_LENGTH = 12;     // 96-bit nonce, recommended for GCM
    private static final int TAG_LENGTH = 128;   // auth tag bits
    private static final String KEY_ENV = "APP_SECRET_KEY";

    private Crypto() {}

    /** Derives a 256-bit AES key from the APP_SECRET_KEY passphrase via SHA-256. */
    private static SecretKeySpec key() throws Exception {
        String passphrase = System.getenv(KEY_ENV);
        if (passphrase == null || passphrase.trim().isEmpty()) {
            throw new IllegalStateException(
                "Environment variable " + KEY_ENV + " is not set. "
                + "Set it before running (it is the decryption key).");
        }
        // Trim so a stray trailing space/newline in the GitHub secret (the most
        // common cause of "Tag mismatch") does NOT change the derived key.
        passphrase = passphrase.trim();
        byte[] hash = MessageDigest.getInstance("SHA-256")
                .digest(passphrase.getBytes(StandardCharsets.UTF_8));
        return new SecretKeySpec(hash, "AES");
    }

    /**
     * A short, NON-SECRET fingerprint of the current APP_SECRET_KEY (first 4
     * bytes of its SHA-256). Lets you compare the key on GitHub vs the key you
     * encrypted with — WITHOUT revealing the key. If two environments print the
     * same fingerprint, they have the same key.
     */
    public static String keyFingerprint() {
        try {
            String passphrase = System.getenv(KEY_ENV);
            if (passphrase == null || passphrase.trim().isEmpty()) return "NO_KEY_SET";
            byte[] hash = MessageDigest.getInstance("SHA-256")
                    .digest(passphrase.trim().getBytes(StandardCharsets.UTF_8));
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < 4; i++) sb.append(String.format("%02x", hash[i]));
            return sb.toString();
        } catch (Exception e) {
            return "ERR";
        }
    }

    /** Encrypts plain text -> Base64(iv || ciphertext+tag). */
    public static String encrypt(String plainText) {
        try {
            byte[] iv = new byte[IV_LENGTH];
            new SecureRandom().nextBytes(iv);

            Cipher cipher = Cipher.getInstance(ALGO);
            cipher.init(Cipher.ENCRYPT_MODE, key(), new GCMParameterSpec(TAG_LENGTH, iv));
            byte[] cipherText = cipher.doFinal(plainText.getBytes(StandardCharsets.UTF_8));

            byte[] combined = new byte[iv.length + cipherText.length];
            System.arraycopy(iv, 0, combined, 0, iv.length);
            System.arraycopy(cipherText, 0, combined, iv.length, cipherText.length);

            return Base64.getEncoder().encodeToString(combined);
        } catch (Exception e) {
            throw new RuntimeException("Encryption failed", e);
        }
    }

    /** Decrypts Base64(iv || ciphertext+tag) -> plain text. */
    public static String decrypt(String encrypted) {
        try {
            byte[] combined = Base64.getDecoder().decode(encrypted);

            byte[] iv = new byte[IV_LENGTH];
            System.arraycopy(combined, 0, iv, 0, IV_LENGTH);

            byte[] cipherText = new byte[combined.length - IV_LENGTH];
            System.arraycopy(combined, IV_LENGTH, cipherText, 0, cipherText.length);

            Cipher cipher = Cipher.getInstance(ALGO);
            cipher.init(Cipher.DECRYPT_MODE, key(), new GCMParameterSpec(TAG_LENGTH, iv));
            return new String(cipher.doFinal(cipherText), StandardCharsets.UTF_8);
        } catch (Exception e) {
            throw new RuntimeException("Decryption failed (wrong APP_SECRET_KEY?) "
                    + "keyFingerprint=" + keyFingerprint(), e);
        }
    }

    /**
     * Run ONCE to turn a plain password into the ciphertext you paste into
     * config.properties. APP_SECRET_KEY must already be set in the environment.
     *
     *   mvn -q exec:java -Dexec.mainClass=Tests.Crypto -Dexec.args="MyPasswordHere"
     * or just run it from your IDE with the password as a program argument.
     */
    public static void main(String[] args) {
        if (args.length == 0) {
            System.out.println("Usage: provide the plain text to encrypt as the first argument.");
            return;
        }
        String enc = encrypt(args[0]);
        System.out.println("Encrypted value (paste into config.properties):");
        System.out.println("app.password.enc=" + enc);
        System.out.println();
        System.out.println("Round-trip check -> decrypts to: " + decrypt(enc));
    }
}