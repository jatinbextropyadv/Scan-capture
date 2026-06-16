package Tests;
import net.sourceforge.tess4j.ITessAPI;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.Word;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;

import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * Finds the text/value located to the RIGHT of a keyword (same visual line)
 * or BELOW a keyword (nearest lower line that is horizontally aligned with it).
 *
 * Works on BOTH kinds of PDFs:
 *   1. Text-based PDFs  -> words + coordinates come from PDFBox's text layer.
 *   2. Scanned/image PDFs (no text layer) -> pages are rendered to images and
 *      OCR'd with Tess4J (Tesseract); word bounding boxes come from the OCR.
 *
 * Either way the page is reduced to word tokens with x/y coordinates (in PDF
 * points), so the same geometry logic answers "right of" and "below".
 *
 * Dependencies: org.apache.pdfbox:pdfbox:3.0.5, net.sourceforge.tess4j:tess4j:5.13.0
 * plus an eng.traineddata file in the tessdata directory.
 */
public class pdf {

    /** One word on the page with its position (PDF points, y grows downward). */
    private record Token(String text, float x, float xEnd, float y, int page) {}

    /** A visual line: tokens on the same page with ~equal Y, left to right. */
    private static class Line {
        final int page;
        final float y;
        final List<Token> tokens = new ArrayList<>();

        Line(Token first) { this.page = first.page(); this.y = first.y(); tokens.add(first); }

        /** Line text: words joined by single spaces. */
        String text() {
            StringBuilder sb = new StringBuilder();
            for (Token t : tokens) {
                if (sb.length() > 0) sb.append(' ');
                sb.append(t.text());
            }
            return sb.toString();
        }

        float startX() { return tokens.get(0).x(); }
        float endX()   { return tokens.get(tokens.size() - 1).xEnd(); }

        /**
         * Approximate X of the character at charIndex of text(), interpolating
         * linearly inside each word (exact per-char positions are not available
         * from OCR, and word-level interpolation is accurate enough here).
         */
        float xAt(int charIndex) {
            int pos = 0;
            for (Token t : tokens) {
                int len = t.text().length();
                if (charIndex < pos + len) {
                    return t.x() + (t.xEnd() - t.x()) * (charIndex - pos) / (float) len;
                }
                pos += len;
                if (charIndex == pos) return t.xEnd(); // the joining space
                pos += 1;
            }
            return endX();
        }

        /** Approximate right edge of the character at charIndex of text(). */
        float xEndAt(int charIndex) { return xAt(charIndex + 1); }
    }

    private static final int   OCR_DPI            = 300;
    private static final float OCR_MIN_CONFIDENCE = 40f;  // drop noise words (0-100)
    private static final float MAX_VALUE_GAP      = 20f;  // points; bigger gap = next column
    private static final float COLUMN_TOLERANCE   = 15f;  // points; column x-overlap slack
    private static final float PDF_LINE_TOLERANCE = 2.0f; // points
    private static final float OCR_LINE_TOLERANCE = 6.0f; // points (OCR boxes are noisier)
    private static final int   MIN_TEXT_LAYER_CHARS = 20; // below this, treat as scanned

    private final List<Line> lines = new ArrayList<>();
    private final boolean usedOcr;

    public pdf(File pdfFile) throws Exception {
        this(pdfFile, new File("tessdata"));
    }

    /** @param tessdataDir directory containing eng.traineddata (only used for scanned PDFs) */
    public pdf(File pdfFile, File tessdataDir) throws Exception {
        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            List<Token> tokens = extractTextLayerTokens(document);

            int totalChars = tokens.stream().mapToInt(t -> t.text().length()).sum();
            usedOcr = totalChars < MIN_TEXT_LAYER_CHARS;
            if (usedOcr) {
                tokens = extractOcrTokens(document, tessdataDir);
            }
            groupIntoLines(tokens, usedOcr ? OCR_LINE_TOLERANCE : PDF_LINE_TOLERANCE);
        }
    }

    /** True if the PDF had no usable text layer and OCR was used instead. */
    public boolean usedOcr() { return usedOcr; }

    // ---------------------------------------------------------- extraction

    /** Words + coordinates from the PDF's native text layer (PDFBox). */
    private static List<Token> extractTextLayerTokens(PDDocument document) throws IOException {
        List<Token> tokens = new ArrayList<>();
        PDFTextStripper stripper = new PDFTextStripper() {
            @Override
            protected void writeString(String string, List<TextPosition> textPositions) {
                // split each positioned fragment into words using per-char positions
                StringBuilder word = new StringBuilder();
                float wordStart = 0, wordEnd = 0, wordY = 0;
                for (TextPosition p : textPositions) {
                    String ch = p.getUnicode();
                    if (ch.isBlank()) {
                        if (word.length() > 0) {
                            tokens.add(new Token(word.toString(), wordStart, wordEnd, wordY, getCurrentPageNo()));
                            word.setLength(0);
                        }
                        continue;
                    }
                    if (word.length() == 0) {
                        wordStart = p.getXDirAdj();
                        wordY = p.getYDirAdj();
                    }
                    word.append(ch);
                    wordEnd = p.getXDirAdj() + p.getWidthDirAdj();
                }
                if (word.length() > 0) {
                    tokens.add(new Token(word.toString(), wordStart, wordEnd, wordY, getCurrentPageNo()));
                }
            }
        };
        stripper.setSortByPosition(true);
        stripper.getText(document); // triggers writeString callbacks
        return tokens;
    }

    /** Words + bounding boxes from OCR (render page -> Tesseract). */
    private static List<Token> extractOcrTokens(PDDocument document, File tessdataDir) throws Exception {
        if (!new File(tessdataDir, "eng.traineddata").isFile()) {
            throw new IllegalStateException("eng.traineddata not found in " + tessdataDir.getAbsolutePath());
        }
        Tesseract tesseract = new Tesseract();
        tesseract.setDatapath(tessdataDir.getAbsolutePath());
        tesseract.setLanguage("eng");

        float pixelsToPoints = 72f / OCR_DPI;
        PDFRenderer renderer = new PDFRenderer(document);
        List<Token> tokens = new ArrayList<>();

        for (int pageIdx = 0; pageIdx < document.getNumberOfPages(); pageIdx++) {
            BufferedImage image = renderer.renderImageWithDPI(pageIdx, OCR_DPI, ImageType.GRAY);
            image = binarize(image); // drop faint backgrounds/highlights that confuse OCR
            if (System.getProperty("saveRender") != null) {
                javax.imageio.ImageIO.write(image, "png", new File("ocr-input-page" + (pageIdx + 1) + ".png"));
            }
            List<Word> words = tesseract.getWords(image, ITessAPI.TessPageIteratorLevel.RIL_WORD);
            for (Word w : words) {
                String text = w.getText() == null ? "" : w.getText().trim();
                if (text.isEmpty() || w.getConfidence() < OCR_MIN_CONFIDENCE) continue;
                // grid lines / scan artifacts come back as tokens like "|", "[+", "--"
                if (text.matches("[|\\[\\]+\\-_=~]+")) continue;
                Rectangle b = w.getBoundingBox();
                tokens.add(new Token(
                        text,
                        b.x * pixelsToPoints,
                        (b.x + b.width) * pixelsToPoints,
                        (float) (b.y + b.height / 2.0) * pixelsToPoints, // vertical center
                        pageIdx + 1));
            }
        }
        return tokens;
    }

    /** Hard black/white threshold: keeps crisp scan text, drops faint JPEG layers. */
    private static BufferedImage binarize(BufferedImage gray) {
        BufferedImage bw = new BufferedImage(gray.getWidth(), gray.getHeight(), BufferedImage.TYPE_BYTE_BINARY);
        for (int y = 0; y < gray.getHeight(); y++) {
            for (int x = 0; x < gray.getWidth(); x++) {
                int lum = gray.getRGB(x, y) & 0xFF;
                bw.setRGB(x, y, lum < 128 ? 0x000000 : 0xFFFFFF);
            }
        }
        return bw;
    }

    private void groupIntoLines(List<Token> tokens, float yTolerance) {
        tokens.sort(Comparator.comparingInt(Token::page)
                              .thenComparing(Token::y)
                              .thenComparing(Token::x));
        for (Token t : tokens) {
            Line last = lines.isEmpty() ? null : lines.get(lines.size() - 1);
            if (last != null && last.page == t.page() && Math.abs(last.y - t.y()) <= yTolerance) {
                last.tokens.add(t);
            } else {
                lines.add(new Line(t));
            }
        }
        for (Line line : lines) line.tokens.sort(Comparator.comparing(Token::x));
    }

    // ------------------------------------------------------------- search

    /**
     * Text immediately to the RIGHT of the keyword on the same visual line,
     * e.g. "Due date" -> "Apr 1, 2026". Null if not found.
     *
     * The first word after the keyword is always taken (labels are often far
     * from their value); after that, a horizontal gap larger than MAX_VALUE_GAP
     * ends the value, so an unrelated next column is not dragged in.
     */
    public String findValueRightOf(String keyword) {
        for (Line line : lines) {
            int[] span = findKeyword(line.text(), keyword);
            if (span == null) continue;

            // first token that starts strictly after the keyword's last char
            int charPos = 0, firstTokenIdx = -1;
            for (int i = 0; i < line.tokens.size(); i++) {
                int end = charPos + line.tokens.get(i).text().length() - 1;
                if (charPos > span[1]) { firstTokenIdx = i; break; }
                charPos = end + 2; // +1 past token, +1 joining space
            }
            if (firstTokenIdx < 0) continue; // keyword is the last thing on the line

            StringBuilder value = new StringBuilder();
            Token prev = null;
            for (int i = firstTokenIdx; i < line.tokens.size(); i++) {
                Token t = line.tokens.get(i);
                if (prev != null && t.x() - prev.xEnd() > MAX_VALUE_GAP) break;
                if (value.length() > 0) value.append(' ');
                value.append(t.text());
                prev = t;
            }
            String rest = value.toString().trim().replaceFirst("^[:\\-]\\s*", "");
            if (!rest.isEmpty()) return rest;
        }
        return null;
    }

    /**
     * For TABLE COLUMN headers: returns only the words on the nearest line
     * below the keyword that fall inside the keyword's x-column,
     * e.g. header "Qty" -> "1" (not the whole item row). Null if not found.
     */
    public String findValueInColumnBelow(String keyword) {
        Line keywordLine = null;
        float kwStartX = 0, kwEndX = 0;
        for (Line line : lines) {
            int[] span = findKeyword(line.text(), keyword);
            if (span == null) continue;
            keywordLine = line;
            kwStartX = line.xAt(span[0]);
            kwEndX   = line.xEndAt(span[1]);
            break;
        }
        if (keywordLine == null) return null;

        Line bestLine = null;
        List<Token> bestTokens = null;
        for (Line line : lines) {
            if (line.page != keywordLine.page) continue;
            if (line.y <= keywordLine.y + PDF_LINE_TOLERANCE) continue;
            List<Token> inColumn = new ArrayList<>();
            for (Token t : line.tokens) {
                float center = (t.x() + t.xEnd()) / 2;
                if (center >= kwStartX - COLUMN_TOLERANCE && center <= kwEndX + COLUMN_TOLERANCE) {
                    inColumn.add(t);
                }
            }
            if (inColumn.isEmpty()) continue;
            if (bestLine == null || line.y < bestLine.y) { bestLine = line; bestTokens = inColumn; }
        }
        if (bestTokens == null) return null;
        StringBuilder sb = new StringBuilder();
        for (Token t : bestTokens) {
            if (sb.length() > 0) sb.append(' ');
            sb.append(t.text());
        }
        return sb.toString();
    }

    /**
     * Text on the nearest line BELOW the keyword that is horizontally aligned
     * with it (column overlap), e.g. "Notes" -> the line under it. Null if not found.
     */
    public String findValueBelow(String keyword) {
        Line keywordLine = null;
        float kwStartX = 0, kwEndX = 0;

        for (Line line : lines) {
            int[] span = findKeyword(line.text(), keyword);
            if (span == null) continue;
            keywordLine = line;
            kwStartX = line.xAt(span[0]);
            kwEndX   = line.xEndAt(span[1]);
            break;
        }
        if (keywordLine == null) return null;

        Line best = null;
        for (Line line : lines) {
            if (line.page != keywordLine.page) continue;
            if (line.y <= keywordLine.y + PDF_LINE_TOLERANCE) continue; // must be below
            boolean overlapsColumn = line.endX() >= kwStartX - 2 && line.startX() <= kwEndX + 2;
            if (!overlapsColumn) continue;
            if (best == null || line.y < best.y) best = line; // nearest below
        }
        return best == null ? null : best.text().trim();
    }

    /** Convenience: try right-of first, fall back to below. */
    public String findValueNear(String keyword) {
        String v = findValueRightOf(keyword);
        return v != null ? v : findValueBelow(keyword);
    }

    /**
     * Locates the keyword in a line, ignoring case AND spacing differences
     * (OCR may read "PO #" as "PO#" or split words oddly).
     * Returns [startIndex, endIndexInclusive] in the original line text, or null.
     */
    private static int[] findKeyword(String lineText, String keyword) {
        String needle = keyword.toLowerCase().replaceAll("\\s+", "");
        StringBuilder norm = new StringBuilder();
        List<Integer> map = new ArrayList<>(); // normalized index -> original index
        for (int i = 0; i < lineText.length(); i++) {
            char c = lineText.charAt(i);
            if (!Character.isWhitespace(c)) {
                norm.append(Character.toLowerCase(c));
                map.add(i);
            }
        }
        int idx = norm.indexOf(needle);
        if (idx < 0 || needle.isEmpty()) return null;
        return new int[]{ map.get(idx), map.get(idx + needle.length() - 1) };
    }

    // ------------------------------------------------------------------ demo
    public static void main(String[] args) throws Exception {
        if (args.length == 0) {
            // no program arguments given -> use built-in defaults
            args = new String[] {"C:\\Users\\abcom\\Downloads\\IMAGE-20415_22700-1768494801-20260406_193941_CENTERSLIN_EMAIL_202604060243125351_153492522.pdf",
                "Account#", "Invoice#", "PO#", "Delivery Date", "Terms", "Qty"
            };
        }
        File pdf = new File(args[0]);
        if (!pdf.isFile()) {
            System.err.println("PDF not found: " + pdf.getAbsolutePath());
            System.err.println("Usage: java Tests.pdf <pdf-path> <keyword1> [keyword2] ...");
            return;
        }

        PdfKeywordValueFinder finder = new PdfKeywordValueFinder(pdf);
        System.out.println("File: " + pdf.getName());
        System.out.println("Mode: " + (finder.usedOcr() ? "OCR (scanned PDF)" : "text layer"));
        for (int i = 1; i < args.length; i++) {
            String kw = args[i];
            System.out.printf("%-14s right=%-26s below=%-38s column=%s%n",
                    kw,
                    finder.findValueRightOf(kw),
                    finder.findValueBelow(kw),
                    finder.findValueBelow(kw));
        }
    }
}