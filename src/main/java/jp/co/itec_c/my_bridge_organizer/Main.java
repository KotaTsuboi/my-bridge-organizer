package jp.co.itec_c.my_bridge_organizer;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;


public class Main {
    static final DateTimeFormatter JP_DATE = DateTimeFormatter.ofPattern("yyyy年MM月dd日");
    static final DateTimeFormatter STD_DATE = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    static final String[] HEADERS = {
            "会社名", "名前", "部署", "役職", "電子メール", "郵便番号", "会社住所",
            "会社電話", "会社FAX", "携帯電話", "名刺交換日", "グループ", "メモ"
    };
    static final int[] COLUMN_INDICES = {1, 2, 3, 4, 7, 9};

    static final String REPOSITORY = "shiroi36/Drawing";
    static final String ISSUE_NUMBER = "547";
    static final String MARK = "Data from myBridge";

    public static void main(String[] args) throws Exception {
        if (args.length != 1) {
            throw new IllegalArgumentException("The length of arguments must be 1");
        }

        String in_excel_name = args[0];

        FileInputStream file = new FileInputStream(in_excel_name);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        Connection conn = DriverManager.getConnection("jdbc:h2:./mydb", "sa", "");

        String createTable = """
                    CREATE TABLE IF NOT EXISTS contacts (
                        company_name VARCHAR(255),
                        name VARCHAR(255),
                        department VARCHAR(255),
                        position VARCHAR(255),
                        email VARCHAR(255),
                        postal_code VARCHAR(255),
                        company_address VARCHAR(255),
                        company_phone VARCHAR(255),
                        company_fax VARCHAR(255),
                        mobile_phone VARCHAR(255),
                        business_card_date VARCHAR(255),
                        `group` VARCHAR(255),
                        note VARCHAR(255)
                    );
                """;

        conn.createStatement().execute(createTable);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue;

            String[] values = new String[13];
            for (int i = 0; i < 13; i++) {
                Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                values[i] = getCellStringValue(cell);

                if (i == 0) {
                    values[i] = values[i].replace("株式会社", "(株)");
                }
            }

            String company = values[0];
            String name = values[1];
            String department = values[2];
            String newDateStr = values[10];
            String newDateStd = toStdDate(newDateStr); // yyyy-MM-ddに変換

            String selectSql = """
                        SELECT business_card_date FROM contacts
                        WHERE company_name = ? AND name = ? AND department = ?
                    """;

            PreparedStatement checkStmt = conn.prepareStatement(selectSql);
            checkStmt.setString(1, company);
            checkStmt.setString(2, name);
            checkStmt.setString(3, department);
            ResultSet rs = checkStmt.executeQuery();

            boolean exists = false;
            boolean shouldUpdate = false;

            if (rs.next()) {
                exists = true;
                String oldDateStr = rs.getString("business_card_date");
                shouldUpdate = isNewer(newDateStd, oldDateStr);
            }
            rs.close();
            checkStmt.close();

            if (!exists) {
                String insertSql = """
                            INSERT INTO contacts (
                                company_name, name, department, position, email, postal_code,
                                company_address, company_phone, company_fax, mobile_phone,
                                business_card_date, `group`, note
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """;
                PreparedStatement ps = conn.prepareStatement(insertSql);
                for (int i = 0; i < 13; i++) {
                    if (i == 10) {
                        ps.setString(i + 1, newDateStd); // 変換済み
                    } else {
                        ps.setString(i + 1, values[i]);
                    }
                }
                ps.executeUpdate();
                ps.close();
            } else if (shouldUpdate) {
                String updateSql = """
                            UPDATE contacts SET
                                position = ?, email = ?, postal_code = ?, company_address = ?,
                                company_phone = ?, company_fax = ?, mobile_phone = ?,
                                business_card_date = ?, `group` = ?, note = ?
                            WHERE company_name = ? AND name = ? AND department = ?
                        """;
                PreparedStatement ps = conn.prepareStatement(updateSql);
                for (int i = 3; i < 13; i++) {
                    if (i == 10) {
                        ps.setString(i - 2, newDateStd);
                    } else {
                        ps.setString(i - 2, values[i]);
                    }
                }
                ps.setString(11, values[0]); // 会社名
                ps.setString(12, values[1]); // 名前
                ps.setString(13, values[2]); // 部署
                ps.executeUpdate();
                ps.close();
            }
        }

        workbook.close();

        // 出力フォルダ作成
        Files.createDirectories(Paths.get("output"));
        Files.createDirectories(Paths.get("card_book"));

        // 会社名一覧取得
        Statement st = conn.createStatement();
        ResultSet rs = st.executeQuery("SELECT DISTINCT company_name FROM contacts");

        while (rs.next()) {
            String company = rs.getString(1);
            exportCompanyExcel(conn, company, in_excel_name);
            exportCardBook(conn, company);
        }

        conn.close();
        System.out.println("✅ 取り込み＆会社別Excel出力完了");

//        deleteComments();
//        upload();
    }

    private static String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // 日付形式だったら yyyy年MM月dd日 にフォーマット（元の形式に合わせる）
                    return DateTimeFormatter.ofPattern("yyyy年MM月dd日")
                            .format(cell.getLocalDateTimeCellValue().toLocalDate());
                } else {
                    // 整数 or 小数を文字列に変換（不要な.0は削除）
                    double d = cell.getNumericCellValue();
                    return (d == Math.floor(d)) ? String.valueOf((long) d) : String.valueOf(d);
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula(); // あるいは evaluate してもよい
            case BLANK:
            case _NONE:
            case ERROR:
            default:
                return "";
        }
    }


    // 日付変換: yyyy年MM月dd日 → yyyy-MM-dd
    static String toStdDate(String jpDate) {
        try {
            LocalDate d = LocalDate.parse(jpDate, JP_DATE);
            return d.format(STD_DATE);
        } catch (DateTimeParseException e) {
            return "";
        }
    }

    // 日付比較（new > old）
    static boolean isNewer(String newDate, String oldDate) {
        try {
            LocalDate n = LocalDate.parse(newDate);
            LocalDate o = LocalDate.parse(oldDate);
            return n.isAfter(o);
        } catch (Exception e) {
            return false;
        }
    }

    // 会社ごとにExcel出力
    static void exportCompanyExcel(Connection conn, String company, String in_excel_name) throws Exception {
        String sql = "SELECT * FROM contacts WHERE company_name = ?";
        PreparedStatement ps = conn.prepareStatement(sql);
        ps.setString(1, company);
        ResultSet rs = ps.executeQuery();

        Workbook outBook = new XSSFWorkbook();
        Sheet sheet = outBook.createSheet("Contacts");

        try (FileInputStream inFile = new FileInputStream(in_excel_name)) {
            Workbook inBook = new XSSFWorkbook(inFile);
            Sheet originalSheet = inBook.getSheetAt(0);
            for (int i = 0; i < HEADERS.length; i++) {
                sheet.setColumnWidth(i, originalSheet.getColumnWidth(i));
            }
            inBook.close();
        }

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < HEADERS.length; i++) {
            headerRow.createCell(i).setCellValue(HEADERS[i]);
        }

        int rowNum = 1;
        while (rs.next()) {
            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < HEADERS.length; i++) {
                String value = rs.getString(i + 1);
                // 日付をyyyy年MM月dd日に戻す
                if (i == 10) {
                    try {
                        LocalDate d = LocalDate.parse(value);
                        value = d.format(JP_DATE);
                    } catch (Exception ignored) {
                    }
                }
                row.createCell(i).setCellValue(value != null ? value : "");
            }
        }

        rs.close();
        ps.close();

        String safeName = company.replaceAll("[\\\\/:*?\"<>|]", "_");
        try (FileOutputStream out = new FileOutputStream("output/" + safeName + ".xlsx")) {
            outBook.write(out);
        }
        outBook.close();
    }


    static void exportCardBook(Connection conn, String company) throws IOException, SQLException {
        String sql = "SELECT * FROM contacts WHERE company_name = ?";
        PreparedStatement ps = conn.prepareStatement(sql);
        ps.setString(1, company);
        ResultSet rs = ps.executeQuery();

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Contacts");

        // スタイル作成（テキスト折り返し用）
        CellStyle style = wb.createCellStyle();
        style.setWrapText(true);

        // フォント（太字用）
        Font boldFont = wb.createFont();
        boldFont.setBold(true);

        int rowIndex = 0;
        while (rs.next()) {
            rowIndex = writeOneCard(sheet, wb, rs, rowIndex);
            rowIndex++;
        }

        // 幅調整（例: A列を広めに）
        sheet.setColumnWidth(0, 12 * 256);
        sheet.setColumnWidth(1, 80 * 256);

        // 保存
        String safeName = company.replaceAll("[\\\\/:*?\"<>|]", "_");
        try (FileOutputStream out = new FileOutputStream("card_book/" + safeName + ".xlsx")) {
            wb.write(out);
        }
        wb.close();

        rs.close();
        ps.close();
    }

    private static int writeOneCard(Sheet sheet, Workbook wb, ResultSet rs, int startRow) throws SQLException {
        Font boldFont = wb.createFont();
        boldFont.setBold(true);

        CellStyle labelStyle = wb.createCellStyle();
        labelStyle.setAlignment(HorizontalAlignment.LEFT);

        for (int i = 0; i < HEADERS.length; i++) {
            Row row = sheet.createRow(startRow + i);

            // ラベル列（A列）
            Cell labelCell = row.createCell(0);
            labelCell.setCellValue(HEADERS[i]);
            labelCell.setCellStyle(labelStyle);

            // データ列（B列）
            String value = rs.getString(i + 1);

            // 日付をyyyy年MM月dd日に戻す
            if (i == 10) {
                try {
                    LocalDate d = LocalDate.parse(value);
                    value = d.format(JP_DATE);
                } catch (Exception ignored) {
                }
            }

            Cell valueCell = row.createCell(1);

            if (i == 1) {
                // 名前は太字で設定
                RichTextString richText = new XSSFRichTextString(value);
                richText.applyFont(boldFont);
                valueCell.setCellValue(richText);
            } else {
                valueCell.setCellValue(value);
            }
        }

        return startRow + HEADERS.length;
    }

    static String convertExcelToMarkdown(File file) throws IOException {
        Workbook wb = new XSSFWorkbook(new FileInputStream(file));
        Sheet sheet = wb.getSheetAt(0);

        // ヘッダー（1行目）を取得
        Row headerRow = sheet.getRow(0);
        int colCount = headerRow.getLastCellNum();

        // ヘッダー行をMarkdownに変換
        StringBuilder markdown = new StringBuilder();
        markdown.append("|");
        for (int i : COLUMN_INDICES) {
            markdown.append(HEADERS[i]).append("|");
        }
        markdown.append("\n|");
        for (int i = 0; i < COLUMN_INDICES.length; i++) {
            markdown.append("---|");
        }
        markdown.append("\n");

        // データ行
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            markdown.append("|");
            for (int j : COLUMN_INDICES) {
                Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                markdown.append(cell.toString().replace("\n", " ")).append("|");
            }
            markdown.append("\n");
        }
        wb.close();

        return markdown.toString();
    }

    static String convertToUTF8(String text) {
        byte[] bytes = text.getBytes(StandardCharsets.UTF_8);

        StringBuilder hex = new StringBuilder();
        for (byte b : bytes) {
            hex.append(String.format("%02x", b)); // 小文字16進
        }

        return hex.toString();
    }

    static void deleteComments() throws InterruptedException, IOException {
        // 既存の自分のコメント一覧を取得
        ProcessBuilder listComments = new ProcessBuilder(
                "gh", "api",
                "/repos/" + REPOSITORY + "/issues/" + ISSUE_NUMBER + "/comments"
        );

        // 一度にはすべてのコメントを列挙できないため繰り返し列挙と削除をする
        while (true) {
            Process process = listComments.start();

            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            StringBuilder json = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                json.append(line);
            }
            process.waitFor();

            // JSONを解析
            ObjectMapper mapper = new ObjectMapper();
            List<Map<String, Object>> comments = mapper.readValue(json.toString(), new TypeReference<>() {
            });

            if (comments.size() == 0) {
                break;
            }

            List<String> commentIdsToDelete = new ArrayList<>();

            for (Map<String, Object> comment : comments) {
                String id = String.valueOf(comment.get("id"));
                String body = (String) comment.get("body");

                // 目印付きのコメントだけ削除対象にする（推奨）
                if (body.contains(MARK)) {
                    commentIdsToDelete.add(id);
                }
            }

            // コメント削除
            for (String commentId : commentIdsToDelete) {
                ProcessBuilder deleteCmd = new ProcessBuilder(
                        "gh", "api", "--method", "DELETE",
                        "/repos/" + REPOSITORY + "/issues/comments/" + commentId
                );
                deleteCmd.inheritIO().start().waitFor();

                Thread.sleep(500);
            }
        }
    }

    static void upload() throws IOException, InterruptedException {
        List<String> companyNames = new ArrayList<>();
        Map<String, String> markdownPerCompany = new LinkedHashMap<>();

        for (File file : Objects.requireNonNull(new File("output").listFiles(f -> f.getName().endsWith(".xlsx")))) {
            String companyName = file.getName().replace(".xlsx", "");
            String tableMarkdown = convertExcelToMarkdown(file);

            companyNames.add(companyName);
            markdownPerCompany.put(companyName, tableMarkdown);
        }

        StringBuilder fullMd = new StringBuilder();
        fullMd.append("## 📋 目次\n\n");
        for (String name : companyNames) {
            fullMd.append("- [").append(name).append("](#").append(convertToUTF8(name)).append(")\n");
        }

        Files.writeString(Paths.get("body.md"), fullMd.toString(), StandardCharsets.UTF_8);

        ProcessBuilder pb = new ProcessBuilder(
                "gh", "issue", "edit", ISSUE_NUMBER,
                "--repo", REPOSITORY,
                "--body-file", "body.md"
        );

        pb.inheritIO().start().waitFor();

        for (Map.Entry<String, String> entry : markdownPerCompany.entrySet()) {
            StringBuilder md = new StringBuilder();
            String anchor = convertToUTF8(entry.getKey());
            md.append("<a id=\"").append(anchor).append("\"></a>\n");
            md.append("### ").append(entry.getKey()).append("\n\n");
            md.append(entry.getValue()).append("\n");
            md.append(MARK).append("\n");
            Files.writeString(Paths.get("comment.md"), md.toString(), StandardCharsets.UTF_8);

            pb = new ProcessBuilder(
                    "gh", "issue", "comment", ISSUE_NUMBER,
                    "--repo", REPOSITORY,
                    "--body-file", "comment.md"
            );

            pb.inheritIO().start().waitFor();

            Thread.sleep(500);
        }

    }

}
