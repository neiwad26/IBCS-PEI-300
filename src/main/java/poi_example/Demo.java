package poi_example;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {

    private static final DataFormatter FORMATTER = new DataFormatter(true);

    static final class Firm {
        final int peiRank;
        final String name;
        final String region;
        final String primaryFocus;
        final double capitalRaisedM;
        final double latestFundSizeB;
        final double aumB;

        Firm(int rank, String name, String region, String primaryFocus, double capM, double lfsB, double aumB) {
            this.peiRank = rank;
            this.name = name;
            this.region = region;
            this.primaryFocus = primaryFocus;
            this.capitalRaisedM = capM;
            this.latestFundSizeB = lfsB;
            this.aumB = aumB;
        }
    }

    static final class Criteria {
        final String regionEquals;
        final double minAumB;
        final double minLatestFundSizeB;
        final double minCapitalRaisedM;
        final String focusContains;
        final List<String> priority;

        Criteria(String regionEquals, double minAumB, double minLatestFundSizeB, double minCapitalRaisedM,
                 String focusContains, List<String> priority) {
            this.regionEquals = regionEquals;
            this.minAumB = minAumB;
            this.minLatestFundSizeB = minLatestFundSizeB;
            this.minCapitalRaisedM = minCapitalRaisedM;
            this.focusContains = focusContains;
            this.priority = priority;
        }
    }

    public static void main(String[] args) throws Exception {
        Path inputPath = Paths.get("PEI300_SampleInput.xlsx");
        String dataSheetName = "PEI 300";

        // Step 1: Enrich missing firms in the input file using web scraping (Jsoup).
        // This looks for rows where key numeric fields like AUM (USD B) are blank and
        // attempts to fill them from Wikipedia.
        int maxToFill = 50; // number of firms to auto-fill per run; adjust as needed
        fillMissingFirmsInPlace(inputPath, dataSheetName, maxToFill);

        // Step 2: Proceed as before: read firms, prompt for criteria, filter, sort, and write results.
        List<Firm> all = readFirms(inputPath, dataSheetName);
        Criteria c = promptCriteria(all);

        List<Firm> filtered = filter(all, c);
        List<Firm> sorted = sort(filtered, c.priority);

        Path out = Paths.get("PEI300_SortedFile.xlsx");
        writeResults(out, c, sorted, inputPath, dataSheetName);

        tryOpen(out);

        System.out.println("\nMatches: " + sorted.size());
        sorted.stream().limit(12).forEach(f -> System.out.printf(
                "%3d | %-32s | %-16s | AUM(B): %.2f | LFS(B): %.2f | CR(M): %.0f | %s%n",
                f.peiRank, f.name, trim(f.region, 16), f.aumB, f.latestFundSizeB, f.capitalRaisedM, f.primaryFocus));
        System.out.println("Wrote: " + out.toAbsolutePath());
    }

    // ----------------------------------------------------------------------
    // NEW: Enrichment / web-scraping helpers
    // ----------------------------------------------------------------------

    private static void fillMissingFirmsInPlace(Path xlsx, String sheetName, int maxToFill) throws Exception {
        try (FileInputStream fis = new FileInputStream(xlsx.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            Sheet sh = wb.getSheet(sheetName);
            if (sh == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            Row header = sh.getRow(sh.getFirstRowNum());
            if (header == null) {
                throw new IllegalStateException("Header row is missing for sheet: " + sheetName);
            }

            Map<String, Integer> col = toHeaderIndexMap(header);

            Integer cFirm = col.get(normalizeHeader("Firm Name"));
            Integer cRegion = col.get(normalizeHeader("Region"));
            Integer cFocus = col.get(normalizeHeader("Primary Focus"));
            Integer cAum = col.get(normalizeHeader("AUM (USD B)"));

            if (cFirm == null || cRegion == null || cFocus == null || cAum == null) {
                throw new IllegalStateException("Expected columns not found when enriching input sheet.");
            }

            int cWiki = ensureColumn(sh, header, "Wikipedia URL");
            int cWebsite = ensureColumn(sh, header, "Website (wiki)");
            int cFounded = ensureColumn(sh, header, "Founded Year (wiki)");
            int cLast = ensureColumn(sh, header, "Last Enriched (UTC)");

            int filled = 0;
            int first = sh.getFirstRowNum() + 1;
            int last = sh.getLastRowNum();

            for (int r = first; r <= last && filled < maxToFill; r++) {
                Row row = sh.getRow(r);
                if (row == null) {
                    continue;
                }

                // Skip rows without a firm name
                String firm = getString(row, col, "Firm Name");
                if (firm.isBlank()) {
                    continue;
                }

                // "Needs filling" = AUM (USD B) blank
                String aumStr = getString(row, col, "AUM (USD B)");
                if (!aumStr.isBlank()) {
                    // Already has AUM, assume row is filled
                    continue;
                }

                System.out.println("Enriching firm from web: " + firm);
                try {
                    // Polite rate limiting between requests
                    Thread.sleep(350);

                    java.util.Optional<WebScraper.Result> opt = WebScraper.fetch(firm);
                    if (opt.isEmpty()) {
                        continue;
                    }
                    WebScraper.Result res = opt.get();

                    // AUM
                    double aumB = parseAumToUsdB(res.aumRaw);
                    if (aumB > 0) {
                        Cell aumOut = row.getCell(cAum);
                        if (aumOut == null) {
                            aumOut = row.createCell(cAum);
                        }
                        aumOut.setCellValue(aumB);
                    }

                    // Region: derive from HQ if region blank
                    String regionExisting = getString(row, col, "Region");
                    if (regionExisting.isBlank() && !res.hqCountry.isBlank()) {
                        Cell regionCell = row.getCell(cRegion);
                        if (regionCell == null) {
                            regionCell = row.createCell(cRegion);
                        }
                        regionCell.setCellValue(mapCountryToRegion(res.hqCountry));
                    }

                    // Primary Focus: guess if blank
                    String focusExisting = getString(row, col, "Primary Focus");
                    if (focusExisting.isBlank() && !res.primaryFocusGuess.isBlank()) {
                        Cell focusCell = row.getCell(cFocus);
                        if (focusCell == null) {
                            focusCell = row.createCell(cFocus);
                        }
                        focusCell.setCellValue(res.primaryFocusGuess);
                    }

                    // Tracking / metadata columns
                    if (!res.wikiUrl.isBlank()) {
                        Cell wikiCell = row.getCell(cWiki);
                        if (wikiCell == null) {
                            wikiCell = row.createCell(cWiki);
                        }
                        wikiCell.setCellValue(res.wikiUrl);
                    }

                    if (!res.website.isBlank()) {
                        Cell webCell = row.getCell(cWebsite);
                        if (webCell == null) {
                            webCell = row.createCell(cWebsite);
                        }
                        webCell.setCellValue(res.website);
                    }

                    if (res.foundedYear != null) {
                        Cell foundedCell = row.getCell(cFounded);
                        if (foundedCell == null) {
                            foundedCell = row.createCell(cFounded);
                        }
                        foundedCell.setCellValue(res.foundedYear);
                    }

                    Cell lastCell = row.getCell(cLast);
                    if (lastCell == null) {
                        lastCell = row.createCell(cLast);
                    }
                    lastCell.setCellValue(LocalDateTime.now().toString());

                    filled++;
                } catch (Exception e) {
                    System.out.println("Failed to enrich '" + firm + "': " + e.getMessage());
                }
            }

            for (int i = 0; i <= header.getLastCellNum(); i++) {
                sh.autoSizeColumn(i);
            }

            try (FileOutputStream fos = new FileOutputStream(xlsx.toFile())) {
                wb.write(fos);
            }

            System.out.println("Web enrichment completed. Filled " + filled + " firm(s) in " + xlsx.getFileName());
        }
    }

    private static int ensureColumn(Sheet sh, Row header, String name) {
        Map<String, Integer> col = toHeaderIndexMap(header);
        String norm = normalizeHeader(name);
        Integer idx = col.get(norm);
        if (idx != null) {
            return idx;
        }
        int newIdx = header.getLastCellNum();
        if (newIdx < 0) {
            newIdx = 0;
        }
        Cell cell = header.getCell(newIdx);
        if (cell == null) {
            cell = header.createCell(newIdx);
        }
        cell.setCellValue(name);
        return newIdx;
    }

    private static double parseAumToUsdB(String aumRaw) {
        if (aumRaw == null || aumRaw.isBlank()) {
            return 0.0;
        }

        String s = aumRaw.toLowerCase(Locale.ROOT)
                .replaceAll(",", "")
                .replaceAll("\\(.*?\\)", "") // remove bracketed notes
                .trim();

        double multiplier = 1.0;
        if (s.contains("billion") || s.contains("bn") || s.matches(".*\\b\\d+(\\.\\d+)?\\s*b\\b.*")) {
            multiplier = 1.0; // billions already
        } else if (s.contains("million") || s.contains("m")) {
            multiplier = 0.001; // million to billion
        }

        Matcher m = Pattern.compile("(\\d+(?:\\.\\d+)?)").matcher(s);
        if (!m.find()) {
            return 0.0;
        }

        try {
            double val = Double.parseDouble(m.group(1));
            return val * multiplier;
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private static String mapCountryToRegion(String country) {
        if (country == null || country.isBlank()) {
            return "";
        }
        String c = country.toLowerCase(Locale.ROOT);
        if (c.contains("united states") || c.equals("canada")) {
            return "North America";
        }
        if (c.contains("united kingdom") || c.contains("germany") || c.contains("france")
                || c.contains("switzerland") || c.contains("luxembourg") || c.contains("ireland")
                || c.contains("spain") || c.contains("italy")) {
            return "Europe";
        }
        if (c.contains("china") || c.contains("japan") || c.contains("singapore")
                || c.contains("india") || c.contains("hong kong")) {
            return "Asia-Pacific";
        }
        return "Other";
    }

    // ----------------------------------------------------------------------
    // Existing methods below (unchanged from your original, except imports)
    // ----------------------------------------------------------------------

    private static List<Firm> readFirms(Path xlsx, String sheetName) throws Exception {
        try (FileInputStream fis = new FileInputStream(xlsx.toFile());
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet s = wb.getSheet(sheetName);
            if (s == null)
                throw new IllegalArgumentException("Sheet not found: " + sheetName);

            Row header = s.getRow(s.getFirstRowNum());
            Map<String, Integer> col = toHeaderIndexMap(header);

            require(col, "Rank");
            require(col, "Firm Name");
            require(col, "Region");
            require(col, "Primary Focus");
            require(col, "Capital Raised (USD M, 2020–24)");
            require(col, "Latest Fund Size (USD B)");
            require(col, "AUM (USD B)");

            int first = s.getFirstRowNum() + 1;
            int last = s.getLastRowNum();
            List<Firm> out = new ArrayList<>();

            for (int r = first; r <= last; r++) {
                Row row = s.getRow(r);
                if (row == null)
                    continue;

                int rank = (int) getNumeric(row, col, "Rank");
                String name = getString(row, col, "Firm Name");
                if (name.isBlank())
                    continue;

                String region = getString(row, col, "Region");
                String focus = getString(row, col, "Primary Focus");
                double capM = getNumeric(row, col, "Capital Raised (USD M, 2020–24)");
                double lfsB = getNumeric(row, col, "Latest Fund Size (USD B)");
                double aumB = getNumeric(row, col, "AUM (USD B)");

                out.add(new Firm(rank, name, region, focus, capM, lfsB, aumB));
            }
            return out;
        }
    }

    private static Map<String, Integer> toHeaderIndexMap(Row header) {
        Map<String, Integer> map = new HashMap<>();
        if (header == null)
            return map;
        short last = header.getLastCellNum();
        for (int c = 0; c < last; c++) {
            Cell cell = header.getCell(c);
            if (cell == null)
                continue;
            String key = normalizeHeader(FORMATTER.formatCellValue(cell));
            if (!key.isBlank())
                map.put(key, c);
        }
        return map;
    }

    private static String normalizeHeader(String s) {
        if (s == null)
            return "";
        return s.replace('\u2013', '-')
                .toLowerCase(Locale.ROOT)
                .replaceAll("[^a-z0-9\\s-]+", "")
                .replaceAll("\\s+", " ")
                .trim();
    }

    private static void require(Map<String, Integer> col, String headerName) {
        if (!col.containsKey(normalizeHeader(headerName))) {
            throw new IllegalStateException("Missing column header: '" + headerName + "'");
        }
    }

    private static String getString(Row row, Map<String, Integer> col, String header) {
        Integer idx = col.get(normalizeHeader(header));
        if (idx == null)
            return "";
        Cell cell = row.getCell(idx);
        return (cell == null) ? "" : FORMATTER.formatCellValue(cell);
    }

    private static double getNumeric(Row row, Map<String, Integer> col, String header) {
        Integer idx = col.get(normalizeHeader(header));
        if (idx == null)
            return 0.0;
        Cell cell = row.getCell(idx);
        if (cell == null)
            return 0.0;

        if (cell.getCellType() == CellType.NUMERIC)
            return cell.getNumericCellValue();

        if (cell.getCellType() == CellType.FORMULA) {
            FormulaEvaluator eval = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            CellType t = eval.evaluateFormulaCell(cell);
            if (t == CellType.NUMERIC)
                return cell.getNumericCellValue();
            String s = FORMATTER.formatCellValue(cell, eval);
            return parseNumber(s);
        }
        String s = FORMATTER.formatCellValue(cell);
        return parseNumber(s);
    }

    private static double parseNumber(String s) {
        if (s == null)
            return 0.0;
        s = s.replace(",", " ").trim().replace(" ", "");
        if (s.isEmpty())
            return 0.0;
        try {
            return Double.parseDouble(s);
        } catch (Exception e) {
            return 0.0;
        }
    }

    private static Criteria promptCriteria(List<Firm> all) {
        Scanner sc = new Scanner(System.in);

        var regions = all.stream().map(f -> f.region)
                .filter(s -> s != null && !s.isBlank())
                .distinct().sorted().limit(12).collect(Collectors.toList());
        System.out.println("Examples of Region Values: " + regions);

        System.out.print("Regions equals (blank=ANY): ");
        String region = sc.nextLine().trim();

        System.out.print("Min AUM (USD B): ");
        double minAumB = parseDouble(sc.nextLine());

        System.out.print("Min Latest Fund Size (USD B): ");
        double minLfsB = parseDouble(sc.nextLine());

        System.out.print("Min Capital Raised (USD M, 2020–24): ");
        double minCapM = parseDouble(sc.nextLine());

        System.out.print("Primary Focus contains (blank = ANY; e.g., Buyout): ");
        String focus = sc.nextLine().trim();

        System.out.print("Priority CSV (default AUM,LATEST_FUND_SIZE,CAPITAL_RAISED,PEI_RANK): ");
        String csv = sc.nextLine().trim();
        List<String> pr = csv.isBlank()
                ? List.of("AUM", "LATEST_FUND_SIZE", "CAPITAL_RAISED", "PEI_RANK")
                : Arrays.stream(csv.split(",")).map(String::trim).filter(s -> !s.isBlank()).toList();

        return new Criteria(region, minAumB, minLfsB, minCapM, focus, pr);
    }

    private static double parseDouble(String s) {
        if (s == null || s.isBlank())
            return 0.0;
        try {
            return Double.parseDouble(s.replace(",", ""));
        } catch (Exception e) {
            return 0.0;
        }
    }

    private static List<Firm> filter(List<Firm> input, Criteria c) {
        return input.stream()
                .filter(f -> c.regionEquals == null || c.regionEquals.isBlank() ||
                        equalsIgnoreCaseTrim(f.region, c.regionEquals))
                .filter(f -> f.aumB >= c.minAumB)
                .filter(f -> f.latestFundSizeB >= c.minLatestFundSizeB)
                .filter(f -> f.capitalRaisedM >= c.minCapitalRaisedM)
                .filter(f -> c.focusContains == null || c.focusContains.isBlank() ||
                        containsIgnoreCase(f.primaryFocus, c.focusContains))
                .collect(Collectors.toList());
    }

    private static List<Firm> sort(List<Firm> list, List<String> priority) {
        Comparator<Firm> cmp = (a, b) -> 0;
        for (String p : priority) {
            String key = p.toUpperCase(Locale.ROOT).replace(" ", "_");
            Comparator<Firm> next = switch (key) {
                case "AUM" -> Comparator.comparingDouble((Firm f) -> f.aumB).reversed();
                case "LATEST_FUND_SIZE" -> Comparator.comparingDouble((Firm f) -> f.latestFundSizeB).reversed();
                case "CAPITAL_RAISED" -> Comparator.comparingDouble((Firm f) -> f.capitalRaisedM).reversed();
                case "PEI_RANK", "RANK" -> Comparator.comparingInt((Firm f) -> f.peiRank); // 1 is best
                default -> (x, y) -> 0;
            };
            cmp = cmp.thenComparing(next);
        }
        return list.stream().sorted(cmp).collect(Collectors.toList());
    }

    private static void writeResults(Path outPath, Criteria c, List<Firm> firms, Path inputPath, String sheetName) throws Exception {
        try (FileInputStream fis = new FileInputStream(inputPath.toFile());
             Workbook sourceWb = new XSSFWorkbook(fis);
             Workbook wb = new XSSFWorkbook()) {

            Sheet source = sourceWb.getSheet(sheetName);
            if (source == null) {
                throw new IllegalArgumentException("Sheet not found in source workbook: " + sheetName);
            }

            Row sourceHeader = source.getRow(source.getFirstRowNum());
            if (sourceHeader == null) {
                throw new IllegalStateException("Source sheet header row is missing: " + sheetName);
            }

            Map<String, Integer> col = toHeaderIndexMap(sourceHeader);
            Integer rankCol = col.get(normalizeHeader("Rank"));
            if (rankCol == null) {
                throw new IllegalStateException("Could not find 'Rank' column in source sheet.");
            }

            short lastCol = sourceHeader.getLastCellNum();

            // Build outputColumns list for desired column order:
            // 1) Rank first, 2) metrics as per user priority, 3) remaining columns in original order (no duplicates)
            List<Integer> outputColumns = new ArrayList<>();
            // 1. Add Rank first
            if (rankCol != null && !outputColumns.contains(rankCol)) {
                outputColumns.add(rankCol);
            }
            // 2. Add metrics in priority order
            for (String p : c.priority) {
                String key = p.toUpperCase(Locale.ROOT).replace(" ", "_");
                Integer idx = null;
                switch (key) {
                    case "AUM" -> idx = col.get(normalizeHeader("AUM (USD B)"));
                    case "LATEST_FUND_SIZE" -> idx = col.get(normalizeHeader("Latest Fund Size (USD B)"));
                    case "CAPITAL_RAISED" -> idx = col.get(normalizeHeader("Capital Raised (USD M, 2020–24)"));
                    case "PEI_RANK", "RANK" -> idx = rankCol;
                    default -> idx = col.get(normalizeHeader(p));
                }
                if (idx != null && !outputColumns.contains(idx)) {
                    outputColumns.add(idx);
                }
            }
            // 3. Add remaining columns in original order, skipping any already present
            for (int cIdx = 0; cIdx < lastCol; cIdx++) {
                if (!outputColumns.contains(cIdx)) {
                    outputColumns.add(cIdx);
                }
            }

            Sheet sh = wb.createSheet("Results");

            int r = 0;
            // Write criteria / filter summary at the top
            r = writeKV(sh, r, "Region equals", nullToEmpty(c.regionEquals));
            r = writeKV(sh, r, "Min AUM (USD B)", String.valueOf(c.minAumB));
            r = writeKV(sh, r, "Min Latest Fund Size (USD B)", String.valueOf(c.minLatestFundSizeB));
            r = writeKV(sh, r, "Min Capital Raised (USD M, 2020–24)", String.valueOf(c.minCapitalRaisedM));
            r = writeKV(sh, r, "Primary Focus contains", nullToEmpty(c.focusContains));
            r = writeKV(sh, r, "Priority", c.priority.toString());
            r++;

            // Header row: copy column headers from the source sheet in the desired order
            Row headerOut = sh.createRow(r++);
            for (int outIdx = 0; outIdx < outputColumns.size(); outIdx++) {
                int srcIdx = outputColumns.get(outIdx);
                Cell src = sourceHeader.getCell(srcIdx);
                String text = (src == null) ? "" : FORMATTER.formatCellValue(src);
                headerOut.createCell(outIdx).setCellValue(text);
            }

            // Build a map of Rank -> source Row for quick lookup
            Map<Integer, Row> byRank = new HashMap<>();
            int firstData = source.getFirstRowNum() + 1;
            int lastData = source.getLastRowNum();
            for (int rr = firstData; rr <= lastData; rr++) {
                Row row = source.getRow(rr);
                if (row == null) {
                    continue;
                }
                Cell rankCell = row.getCell(rankCol);
                if (rankCell == null || rankCell.getCellType() != CellType.NUMERIC) {
                    continue;
                }
                int rank = (int) rankCell.getNumericCellValue();
                byRank.put(rank, row);
            }

            // For each filtered/sorted firm, copy the original row using outputColumns order
            for (Firm f : firms) {
                Row srcRow = byRank.get(f.peiRank);
                Row outRow = sh.createRow(r++);

                if (srcRow == null) {
                    // Fallback: at least write rank and firm name
                    outRow.createCell(0).setCellValue(f.peiRank);
                    outRow.createCell(1).setCellValue(f.name);
                    continue;
                }

                for (int outIdx = 0; outIdx < outputColumns.size(); outIdx++) {
                    int srcIdx = outputColumns.get(outIdx);
                    Cell src = srcRow.getCell(srcIdx);
                    if (src == null) {
                        continue;
                    }
                    Cell dst = outRow.createCell(outIdx);
                    switch (src.getCellType()) {
                        case STRING -> dst.setCellValue(src.getStringCellValue());
                        case NUMERIC -> dst.setCellValue(src.getNumericCellValue());
                        case BOOLEAN -> dst.setCellValue(src.getBooleanCellValue());
                        case FORMULA -> {
                            // Evaluate formulas as text to avoid broken references in the new workbook
                            String txt = FORMATTER.formatCellValue(src);
                            dst.setCellValue(txt);
                        }
                        default -> {
                            String txt = FORMATTER.formatCellValue(src);
                            if (!txt.isEmpty()) {
                                dst.setCellValue(txt);
                            }
                        }
                    }
                }
            }

            for (int outIdx = 0; outIdx < outputColumns.size(); outIdx++) {
                sh.autoSizeColumn(outIdx);
            }

            try (FileOutputStream fos = new FileOutputStream(outPath.toFile())) {
                wb.write(fos);
            }
        }
    }

    private static int writeKV(Sheet sh, int r, String k, String v) {
        Row row = sh.createRow(r++);
        row.createCell(0).setCellValue(k);
        row.createCell(1).setCellValue(v);
        return r;
    }

    private static String nullToEmpty(String s) {
        return s == null ? "" : s;
    }

    private static boolean equalsIgnoreCaseTrim(String a, String b) {
        return a != null && b != null && a.trim().equalsIgnoreCase(b.trim());
    }

    private static boolean containsIgnoreCase(String hay, String needle) {
        if (hay == null || needle == null)
            return false;
        return hay.toLowerCase(Locale.ROOT).contains(needle.toLowerCase(Locale.ROOT).trim());
    }

    private static String trim(String s, int n) {
        if (s == null)
            return "";
        return s.length() <= n ? s : s.substring(0, n - 1) + "…";
    }

    private static void tryOpen(Path out) {
        File f = out.toFile();

        // 1) Desktop API
        try {
            if (Desktop.isDesktopSupported()) {
                Desktop.getDesktop().open(f);
                return;
            }
        } catch (Exception ignore) {
        }

        String os = System.getProperty("os.name").toLowerCase(Locale.ROOT);
        try {
            if (os.contains("mac")) {
                new ProcessBuilder("open", f.getAbsolutePath()).inheritIO().start();
                return;
            } else if (os.contains("win")) {
                new ProcessBuilder("cmd", "/c", "start", "\"\"", f.getAbsolutePath()).inheritIO().start();
                return;
            } else {
                new ProcessBuilder("xdg-open", f.getAbsolutePath()).inheritIO().start();
                return;
            }
        } catch (IOException ignore) {
        }

        try {
            if (os.contains("mac")) {
                new ProcessBuilder("open", "-a", "Microsoft Excel", f.getAbsolutePath())
                        .inheritIO().start();
                return;
            }
        } catch (IOException ignore) {
        }

        System.out.println("Saved to: " + f.getAbsolutePath());
        System.out.println("If it didn't open automatically, open it with Excel/Numbers or upload to Google Sheets.");
    }
}

