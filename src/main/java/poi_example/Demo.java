package poi_example;

import java.awt.Desktop;
import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.*;
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
        final String focusContains;
        final List<String> priority;

        Criteria(String regionEquals, double minAumB, double minLatestFundSizeB, String focusContains, List<String> priority) {
            this.regionEquals = regionEquals;
            this.minAumB = minAumB;
            this.minLatestFundSizeB = minLatestFundSizeB;
            this.focusContains = focusContains;
            this.priority = priority;
        }
    }

    public static void main(String[] args) throws Exception {
        Path inputPath = Paths.get("PEI300_SampleInput.xlsx");
        String dataSheetName = "PEI 300";

        List<Firm> all = readFirms(inputPath, dataSheetName);
        Criteria c = promptCriteria(all);

        List<Firm> filtered = filter(all, c);
        List<Firm> sorted   = sort(filtered, c.priority);

        Path out = Paths.get("PEI300_SortedFile.xlsx");
        writeResults(out, c, sorted);

        tryOpen(out);

        System.out.println("\nMatches: " + sorted.size());
        sorted.stream().limit(12).forEach(f ->
            System.out.printf(
                "%3d | %-32s | %-16s | AUM(B): %.2f | LFS(B): %.2f | CR(M): %.0f | %s%n",
                f.peiRank, f.name, trim(f.region, 16), f.aumB, f.latestFundSizeB, f.capitalRaisedM, f.primaryFocus
            )
        );
        System.out.println("Wrote: " + out.toAbsolutePath());
    }

    private static List<Firm> readFirms(Path xlsx, String sheetName) throws Exception {
        try (FileInputStream fis = new FileInputStream(xlsx.toFile());
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet s = wb.getSheet(sheetName);
            if (s == null) throw new IllegalArgumentException("Sheet not found: " + sheetName);

            Row header = s.getRow(s.getFirstRowNum());
            Map<String,Integer> col = toHeaderIndexMap(header);

            require(col, "Rank");
            require(col, "Firm Name");
            require(col, "Region");
            require(col, "Primary Focus");
            require(col, "Capital Raised (USD M, 2020–24)");
            require(col, "Latest Fund Size (USD B)");
            require(col, "AUM (USD B)");

            int first = s.getFirstRowNum() + 1;
            int last  = s.getLastRowNum();
            List<Firm> out = new ArrayList<>();

            for (int r = first; r <= last; r++) {
                Row row = s.getRow(r);
                if (row == null) continue;

                int rank      = (int) getNumeric(row, col, "Rank");
                String name   = getString(row, col, "Firm Name");
                if (name.isBlank()) continue;

                String region = getString(row, col, "Region");
                String focus  = getString(row, col, "Primary Focus");
                double capM   = getNumeric(row, col, "Capital Raised (USD M, 2020–24)");
                double lfsB   = getNumeric(row, col, "Latest Fund Size (USD B)");
                double aumB   = getNumeric(row, col, "AUM (USD B)");

                out.add(new Firm(rank, name, region, focus, capM, lfsB, aumB));
            }
            return out;
        }
    }

    private static Map<String,Integer> toHeaderIndexMap(Row header) {
        Map<String,Integer> map = new HashMap<>();
        if (header == null) return map;
        short last = header.getLastCellNum();
        for (int c = 0; c < last; c++) {
            Cell cell = header.getCell(c);
            if (cell == null) continue;
            String key = normalizeHeader(FORMATTER.formatCellValue(cell));
            if (!key.isBlank()) map.put(key, c);
        }
        return map;
    }

    private static String normalizeHeader(String s) {
        if (s == null) return "";
        return s.replace('\u2013','-')
                .toLowerCase(Locale.ROOT)
                .replaceAll("[^a-z0-9\\s-]+", "")
                .replaceAll("\\s+"," ")
                .trim();
    }

    private static void require(Map<String,Integer> col, String headerName) {
        if (!col.containsKey(normalizeHeader(headerName))) {
            throw new IllegalStateException("Missing column header: '" + headerName + "'");
        }
    }

    private static String getString(Row row, Map<String,Integer> col, String header) {
        Integer idx = col.get(normalizeHeader(header));
        if (idx == null) return "";
        Cell cell = row.getCell(idx);
        return (cell == null) ? "" : FORMATTER.formatCellValue(cell);
    }

    private static double getNumeric(Row row, Map<String,Integer> col, String header) {
        Integer idx = col.get(normalizeHeader(header));
        if (idx == null) return 0.0;
        Cell cell = row.getCell(idx);
        if (cell == null) return 0.0;

        if (cell.getCellType() == CellType.NUMERIC) return cell.getNumericCellValue();

        if (cell.getCellType() == CellType.FORMULA) {
            FormulaEvaluator eval = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            CellType t = eval.evaluateFormulaCell(cell);
            if (t == CellType.NUMERIC) return cell.getNumericCellValue();
            String s = FORMATTER.formatCellValue(cell, eval);
            return parseNumber(s);
        }
        String s = FORMATTER.formatCellValue(cell);
        return parseNumber(s);
    }

    private static double parseNumber(String s) {
        if (s == null) return 0.0;
        s = s.replace(","," ").trim().replace(" ","");
        if (s.isEmpty()) return 0.0;
        try { return Double.parseDouble(s); } catch (Exception e) { return 0.0; }
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

        System.out.print("Primary Focus contains (blank = ANY; e.g., Buyout): ");
        String focus = sc.nextLine().trim();

        System.out.print("Priority CSV (default AUM,LATEST_FUND_SIZE,CAPITAL_RAISED,PEI_RANK): ");
        String csv = sc.nextLine().trim();
        List<String> pr = csv.isBlank()
                ? List.of("AUM","LATEST_FUND_SIZE","CAPITAL_RAISED","PEI_RANK")
                : Arrays.stream(csv.split(",")).map(String::trim).filter(s -> !s.isBlank()).toList();

        return new Criteria(region, minAumB, minLfsB, focus, pr);
    }

    private static double parseDouble(String s) {
        if (s == null || s.isBlank()) return 0.0;
        try { return Double.parseDouble(s.replace(",","")); } catch (Exception e) { return 0.0; }
    }

    private static List<Firm> filter(List<Firm> input, Criteria c) {
        return input.stream()
                .filter(f -> c.regionEquals == null || c.regionEquals.isBlank() ||
                             equalsIgnoreCaseTrim(f.region, c.regionEquals))
                .filter(f -> f.aumB >= c.minAumB)
                .filter(f -> f.latestFundSizeB >= c.minLatestFundSizeB)
                .filter(f -> c.focusContains == null || c.focusContains.isBlank() ||
                             containsIgnoreCase(f.primaryFocus, c.focusContains))
                .collect(Collectors.toList());
    }

    private static List<Firm> sort(List<Firm> list, List<String> priority) {
        Comparator<Firm> cmp = (a, b) -> 0;
        for (String p : priority) {
            String key = p.toUpperCase(Locale.ROOT).replace(" ", "_");
            Comparator<Firm> next = switch (key) {
                case "AUM"              -> Comparator.comparingDouble((Firm f) -> f.aumB).reversed();
                case "LATEST_FUND_SIZE" -> Comparator.comparingDouble((Firm f) -> f.latestFundSizeB).reversed();
                case "CAPITAL_RAISED"   -> Comparator.comparingDouble((Firm f) -> f.capitalRaisedM).reversed();
                case "PEI_RANK", "RANK" -> Comparator.comparingInt((Firm f) -> f.peiRank); // 1 is best
                default                 -> (x, y) -> 0;
            };
            cmp = cmp.thenComparing(next);
        }
        return list.stream().sorted(cmp).collect(Collectors.toList());
    }

    private static void writeResults(Path outPath, Criteria c, List<Firm> firms) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sh = wb.createSheet("Results");

            int r = 0;
            r = writeKV(sh, r, "Region equals", nullToEmpty(c.regionEquals));
            r = writeKV(sh, r, "Min AUM (USD B)", String.valueOf(c.minAumB));
            r = writeKV(sh, r, "Min Latest Fund Size (USD B)", String.valueOf(c.minLatestFundSizeB));
            r = writeKV(sh, r, "Primary Focus contains", nullToEmpty(c.focusContains));
            r = writeKV(sh, r, "Priority", c.priority.toString());
            r++;

            String[] headers = {"Rank","Firm Name","Region","Primary Focus",
                                "AUM (USD B)","Latest Fund Size (USD B)","Capital Raised (USD M, 2020–24)"};
            Row h = sh.createRow(r++);
            for (int i = 0; i < headers.length; i++) h.createCell(i).setCellValue(headers[i]);

            for (Firm f : firms) {
                Row row = sh.createRow(r++);
                row.createCell(0).setCellValue(f.peiRank);
                row.createCell(1).setCellValue(f.name);
                row.createCell(2).setCellValue(f.region);
                row.createCell(3).setCellValue(f.primaryFocus);
                row.createCell(4).setCellValue(f.aumB);
                row.createCell(5).setCellValue(f.latestFundSizeB);
                row.createCell(6).setCellValue(f.capitalRaisedM);
            }

            for (int i = 0; i < headers.length; i++) sh.autoSizeColumn(i);
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

    private static String nullToEmpty(String s) { return s == null ? "" : s; }
    private static boolean equalsIgnoreCaseTrim(String a, String b) {
        return a != null && b != null && a.trim().equalsIgnoreCase(b.trim());
    }
    private static boolean containsIgnoreCase(String hay, String needle) {
        if (hay == null || needle == null) return false;
        return hay.toLowerCase(Locale.ROOT).contains(needle.toLowerCase(Locale.ROOT).trim());
    }
    private static String trim(String s, int n) {
        if (s == null) return "";
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
        } catch (Exception ignore) { }

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
        } catch (IOException ignore) { }

        try {
            if (os.contains("mac")) {
                new ProcessBuilder("open", "-a", "Microsoft Excel", f.getAbsolutePath())
                    .inheritIO().start();
                return;
            }
        } catch (IOException ignore) { }

        System.out.println("Saved to: " + f.getAbsolutePath());
        System.out.println("If it didn't open automatically, open it with Excel/Numbers or upload to Google Sheets.");
    }
}
