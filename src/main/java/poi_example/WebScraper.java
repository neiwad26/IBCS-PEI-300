package poi_example;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Locale;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WebScraper {

    static final class Result {
        final String wikiUrl;
        final String website;
        final String hqCountry;
        final Integer foundedYear;
        final String aumRaw;
        final String primaryFocusGuess;

        Result(String wikiUrl, String website, String hqCountry, Integer foundedYear,
               String aumRaw, String primaryFocusGuess) {
            this.wikiUrl = wikiUrl;
            this.website = website;
            this.hqCountry = hqCountry;
            this.foundedYear = foundedYear;
            this.aumRaw = aumRaw;
            this.primaryFocusGuess = primaryFocusGuess;
        }

        static Result empty() {
            return new Result("", "", "", null, "", "");
        }
    }

    private static final String UA = "Mozilla/5.0 (compatible; PEI-IA/1.0)";
    private static final Pattern YEAR = Pattern.compile("(19|20)\\d{2}");
    private static final String[] AUM_HEADERS = {
            "AUM",
            "Assets under management",
            "Assets Under Management",
            "Assets under management (AUM)",
            "Total assets"
    };

    /**
     * Fetch basic information about a private equity firm from Wikipedia.
     * It will:
     *  - search for "<firmName> private equity"
     *  - open the first result (or a direct /wiki/<firmName> page as fallback)
     *  - scrape the infobox for Website, Headquarters, Founded, and AUM
     *  - guess a primary focus from the page body text
     */
    public static Optional<Result> fetch(String firmName) {
        try {
            String q = URLEncoder.encode(firmName + " private equity", StandardCharsets.UTF_8);
            String searchUrl = "https://en.wikipedia.org/w/index.php?search=" + q;

            // 1) Search results page
            Document search = Jsoup.connect(searchUrl)
                    .userAgent(UA)
                    .timeout(12_000)
                    .get();

            Element first = search.selectFirst("div.mw-search-result-heading > a");
            String pageUrl = (first != null)
                    ? "https://en.wikipedia.org" + first.attr("href")
                    : "https://en.wikipedia.org/wiki/" + URLEncoder.encode(firmName, StandardCharsets.UTF_8);

            // 2) Actual firm page
            Document doc = Jsoup.connect(pageUrl)
                    .userAgent(UA)
                    .timeout(12_000)
                    .get();

            Element infobox = doc.selectFirst("table.infobox.vcard, table.infobox");

            String website = "";
            String hq = "";
            String aumText = "";
            Integer foundedYear = null;

            if (infobox != null) {
                website = cellText(infobox, "Website").orElseGet(() -> {
                    Element wb = infobox.selectFirst("tr:has(th:matches((?i)Website)) td a[href]");
                    return wb != null ? wb.absUrl("href") : "";
                });

                hq = cellText(infobox, "Headquarters").orElse("");
                String founded = cellText(infobox, "Founded").orElse("");
                foundedYear = firstYear(founded);

                // Try several common labels for AUM / assets, in order
                for (String header : AUM_HEADERS) {
                    aumText = firstRowContaining(infobox, header).orElse("");
                    if (aumText != null && !aumText.isBlank()) {
                        break;
                    }
                }
            }

            // Guess primary focus from body text.
            // This is a heuristic used only when no primary focus is provided in the input file.
            String body = doc.text().toLowerCase(Locale.ROOT);
            String focusGuess;
            if (body.contains("infrastructure")) {
                focusGuess = "Infrastructure";
            } else if (body.contains("real estate")) {
                focusGuess = "Real Estate";
            } else if (body.contains("venture capital")) {
                focusGuess = "Venture";
            } else if (body.contains("private credit") || body.contains("direct lending") || body.contains("credit fund")) {
                focusGuess = "Credit";
            } else if (body.contains("growth equity") || body.contains("growth capital") || body.contains("growth investments")) {
                focusGuess = "Growth";
            } else if (body.contains("buyout")) {
                focusGuess = "Buyout";
            } else {
                focusGuess = "";
            }

            return Optional.of(new Result(
                    pageUrl,
                    website,
                    extractCountry(hq),
                    foundedYear,
                    aumText,
                    focusGuess
            ));
        } catch (Exception e) {
            System.out.println("WebScraper error for '" + firmName + "': " + e.getMessage());
            return Optional.empty();
        }
    }

    private static Optional<String> cellText(Element infobox, String headerContains) {
        Element row = infobox.selectFirst("tr:has(th:matches((?i)" + Pattern.quote(headerContains) + "))");
        if (row == null) {
            return Optional.empty();
        }
        Element td = row.selectFirst("td");
        return Optional.ofNullable(td != null ? td.text() : null);
    }

    private static Optional<String> firstRowContaining(Element infobox, String needle) {
        Element row = infobox.selectFirst("tr:has(th:matches((?i)" + Pattern.quote(needle) + "))");
        if (row == null) {
            return Optional.empty();
        }
        Element td = row.selectFirst("td");
        return Optional.ofNullable(td != null ? td.text() : null);
    }

    private static Integer firstYear(String s) {
        if (s == null) {
            return null;
        }
        Matcher m = YEAR.matcher(s);
        return m.find() ? Integer.parseInt(m.group()) : null;
    }

    private static String extractCountry(String hq) {
        if (hq == null || hq.isBlank()) {
            return "";
        }
        String[] parts = hq.split(",");
        String last = parts[parts.length - 1].trim();
        String l = last.toLowerCase(Locale.ROOT);
        if (l.equals("us") || l.equals("u.s.") || l.equals("usa") || l.equals("u.s.a")) {
            return "United States";
        }
        if (l.equals("uk") || l.equals("u.k.")) {
            return "United Kingdom";
        }
        return last;
    }
}