package org.x.dbstructtoword;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

import java.io.*;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author OSP_TEAM
 */
@Component
public class PanWeiDbStructToWordTool {
    @Autowired
    private NamedParameterJdbcTemplate namedParameterJdbcTemplate;

    private final static String QUERY_TABLES_SQL =
            "SELECT tb.table_name, d.description as table_comment " +
                    "FROM information_schema.tables tb " +
                    "JOIN pg_class c ON c.relname = tb.table_name " +
                    "LEFT JOIN pg_description d ON d.objoid = c.oid AND d.objsubid = '0' " +
                    "WHERE tb.table_schema = :schemaName";

    private final static String QUERY_STRUCT_SQL =
            "SELECT col.column_name, col.data_type as column_type, col.is_nullable, d.description as column_comment " +
                    "FROM information_schema.columns col " +
                    "JOIN pg_class c ON c.relname = col.table_name " +
                    "LEFT JOIN pg_description d ON d.objoid = c.oid AND d.objsubid = col.ordinal_position " +
                    "WHERE col.table_schema = :schemaName and col.table_name=:tableName " +
                    "ORDER BY col.ordinal_position;";

    public void writeWord(String schemaName, String tableName, String tableComment) {
        XWPFDocument document = loadDocument();
        writeTableToWord(document, schemaName, tableName, tableComment);
        writeDocument(document);
    }

    public void writeWord(String schemaName) {
        Map<String, Object> param = new HashMap<>();
        param.put("schemaName", schemaName);

        List<Map<String, Object>> tableList = namedParameterJdbcTemplate.queryForList(QUERY_TABLES_SQL, param);
        if (tableList == null || tableList.isEmpty()) {
            System.out.println("Cannot find any table");
            return;
        }

        XWPFDocument document = loadDocument();
        for (Map<String, Object> tableInfo : tableList) {
            String tableName = "" + tableInfo.get("TABLE_NAME");
            String tableComment = "" + tableInfo.get("TABLE_COMMENT");
            if ("null".equals(tableComment)) {
                tableComment = tableName;
            }
            writeTableToWord(document, schemaName, tableName, tableComment);
        }
        writeDocument(document);
    }

    private void writeTableToWord(XWPFDocument document, String schemaName, String tableName, String tableComment) {
        Map<String, Object> param = new HashMap<>();
        param.put("schemaName", schemaName);
        param.put("tableName", tableName);

        List<Map<String, Object>> list = namedParameterJdbcTemplate.queryForList(QUERY_STRUCT_SQL, param);
        if (list == null || list.isEmpty()) {
            System.out.println("Cannot find tale: " + tableName);
            return;
        }

        String tableTitle = (StringUtils.isEmpty(tableComment)) ? tableName : tableComment + " (" + tableName + ")";
        writeTableToWord(document, list, tableTitle);
    }

    private void writeDocument(XWPFDocument document) {
        try (FileOutputStream out = new FileOutputStream(new File("/Users/hongyangxiao/xtemp/test.docx"))) {
            document.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void writeTableToWord(XWPFDocument document, List<Map<String, Object>> list, String tableTitle) {
        addTitleLine(document, tableTitle);

        XWPFTable table = document.createTable();
        CTTblWidth tableWidth = table.getCTTbl().addNewTblPr().addNewTblW();
        tableWidth.setType(STTblWidth.DXA);
        tableWidth.setW(BigInteger.valueOf(8000));
        writeTableHead(table);
        writeTableContent(table, list);

        addBlankLine(document);
    }

    private void addBlankLine(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun paragraphRun = paragraph.createRun();
        paragraphRun.setText("\r");
    }

    private void addTitleLine(XWPFDocument document, String title) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setStyle("1");
        XWPFRun paragraphRun = paragraph.createRun();
        paragraphRun.setText(title);
        //paragraphRun.setBold(true);
        //paragraphRun.setFontSize(15);
    }

    private void writeTableContent(XWPFTable table, List<Map<String, Object>> list) {
        for (Map<String, Object> columnInfo : list) {
            writeContentRow(table, columnInfo);
        }
    }

    private void writeContentRow(XWPFTable table, Map<String, Object> columnInfo) {
        XWPFTableRow row = table.createRow();
        row.getCell(0).setText("" + columnInfo.get("COLUMN_NAME"));
        row.getCell(1).setText(transColumnType("" + columnInfo.get("COLUMN_TYPE")));
        row.getCell(2).setText("" + columnInfo.get("COLUMN_COMMENT"));

        String isNullableStr = ("NO".equalsIgnoreCase("" + columnInfo.get("IS_NULLABLE"))) ? "是" : "否";
        if (null != columnInfo.get("COLUMN_DEFAULT")) {
            isNullableStr = isNullableStr + " (默认值：" + columnInfo.get("COLUMN_DEFAULT") + ")";
        }

        row.getCell(3).setText(isNullableStr);
    }

    private String transColumnType(String columnType) {
        if ("character varying".equalsIgnoreCase(columnType)) {
            return "varchar";
        } else if ("integer".equalsIgnoreCase(columnType)) {
            return "int";
        } else if ("timestamp without time zone".equalsIgnoreCase(columnType)) {
            return "timestamp";
        }

        return columnType;
    }

    private void writeTableHead(XWPFTable table) {
        XWPFTableRow row = table.getRow(0);
        writeHeadCell(row.getCell(0), "字段名称");
        writeHeadCell(row.addNewTableCell(), "字段类型");
        writeHeadCell(row.addNewTableCell(), "字段说明");
        writeHeadCell(row.addNewTableCell(), "是否必须");
    }

    private void writeHeadCell(XWPFTableCell c, String text) {
        c.removeParagraph(0);
        XWPFParagraph newPara = new XWPFParagraph(c.getCTTc().addNewP(), c);
        XWPFRun run = newPara.createRun();
        newPara.setAlignment(ParagraphAlignment.CENTER);
        //run.getCTR().addNewRPr().addNewColor().setVal("CCCCCC");
        run.setText(text);
        run.setBold(true);
        c.setColor("CCCCCC");
    }

    private XWPFDocument loadDocument() {
        try (InputStream is = new FileInputStream("/Users/hongyangxiao/xtemp/worddemo.docx")) {
            return new XWPFDocument(is);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }
}
