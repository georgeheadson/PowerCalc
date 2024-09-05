package ve.powercalc;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

import org.apache.poi.hssf.usermodel.*;
import static java.lang.System.exit;

public class PowerCount {

    public static String getPower(Integer id, String dateFrom, String dateTo, String filename) throws SQLException, IOException {
        String host = "10.181.10.55";
        int    port = 1521;
        String sid  = "enf0";
        String user = "enforce_dba";
        String pwd  = "jhfrk";

        try {
            Class.forName("oracle.jdbc.driver.OracleDriver");
        } catch (ClassNotFoundException e) {
            System.out.println("Oracle JDBC Driver is not found");
            exit (-1);
        }
        String url = String.format("jdbc:oracle:thin:@%s:%d:%s", host, port, sid);

        Connection connection = null;
        try {
            connection = DriverManager.getConnection(url, user, pwd);
        } catch (SQLException e) {
            System.out.println("Connection Failed : " + e.getMessage());
        }

        if (connection != null) {
            Statement statement = connection.createStatement();

            String sql = "SELECT idname FROM srez WHERE id = " + id;
            ResultSet resultSet = null;
            resultSet = statement.executeQuery(sql);
            resultSet.next();
            String name = resultSet.getString(1);


            FileOutputStream fileOut = null;
            fileOut = new FileOutputStream(filename);
            HSSFWorkbook workbook = new HSSFWorkbook();
            writeAnalyze(id, name, dateFrom, dateTo, statement, workbook);
            writePower(id, name, dateFrom, dateTo, statement, workbook);
            workbook.write(fileOut);
            fileOut.flush();
            fileOut.close();

            //workbook.close();

            connection.close();
        }
        else {
            System.out.println("Failed to make connection!");
        }

        return filename;
    }

    public static void writeAnalyze(int groupId, String groupName, String dateFrom, String dateTo, Statement statement, HSSFWorkbook workbook) throws SQLException, IOException {

        String sql = "select id_lo, idname_lo, " +
                "pk_enf_dop.get_count_avg_power_pik(id_lo, to_date('" + dateFrom + "', 'yyyy.mm.dd'), to_date('" + dateTo + "', 'yyyy.mm.dd')), " +
                "pk_enf_ask.potr_nachmon_obh(id_lo, to_date('" + dateTo + "', 'yyyy.mm.dd')) " +
                "from power_count_list where id_hi = " + groupId + " order by idname_lo";

        ResultSet resultSet = null;
        resultSet = statement.executeQuery(sql);

        HSSFSheet sheet = workbook.createSheet("Анализ");

        HSSFRow rowTitle = sheet.createRow((short)0);
        rowTitle.createCell(0).setCellValue("Анализ полноты данных");
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 18);
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        rowTitle.getCell(0).setCellStyle(cellStyle);

        HSSFRow rowName = sheet.createRow((short)1);
        rowName.createCell(0).setCellValue(groupName);
        rowName.getCell(0).setCellStyle(cellStyle);
        HSSFRow rowDate = sheet.createRow((short)2);
        rowDate.createCell(0).setCellValue(dateFrom + " - " + dateTo);
        rowDate.getCell(0).setCellStyle(cellStyle);

        HSSFRow rowHead = sheet.createRow((short)3);
        rowHead.createCell(0).setCellValue("ID");
        rowHead.createCell(1).setCellValue("Потребитель");
        rowHead.createCell(2).setCellValue("30 минут");
        rowHead.createCell(3).setCellValue("Расход");

        int i = 4;
        while(resultSet.next()) {
            int id = resultSet.getInt(1);
            String idname = resultSet.getString(2);
            long value1 = resultSet.getLong(3);
            long value2 = resultSet.getLong(4);

            HSSFRow row = sheet.createRow((short) i);
            row.createCell(0).setCellValue(id);
            row.createCell(1).setCellValue(idname);
            row.createCell(2).setCellValue(value1);
            row.createCell(3).setCellValue(value2);

            i++;
        }

        sheet.autoSizeColumn(1);
    }

    public static void writePower(int groupId, String groupName, String dateFrom, String dateTo, Statement statement, HSSFWorkbook workbook) throws SQLException, IOException {

        String sql = "select id_lo, idname_lo, " +
                "pk_enf_dop.get_avg_power_pik(id_lo, to_date('" + dateFrom + "', 'yyyy.mm.dd'), to_date('" + dateTo + "', 'yyyy.mm.dd')), " +
                "pk_enf_dop.get_avg_power_pik_all_days(id_lo, to_date('" + dateFrom + "', 'yyyy.mm.dd'), to_date('" + dateTo + "', 'yyyy.mm.dd')), " +
                "pk_enf_dop.get_avg_power_pik_max(id_lo, to_date('" + dateFrom + "', 'yyyy.mm.dd'), to_date('" + dateTo + "', 'yyyy.mm.dd')), " +
                "pk_enf_dop.get_avg_power_work_days(id_lo, to_date('" + dateFrom + "', 'yyyy.mm.dd'), to_date('" + dateTo + "', 'yyyy.mm.dd')), " +
                "pk_enf_dop.get_avg_power_work_days_max(id_lo, to_date('" + dateFrom + "', 'yyyy.mm.dd'), to_date('" + dateTo + "', 'yyyy.mm.dd')), " +
                "pk_enf_dop.get_avg_power_all_max(id_lo, to_date('" + dateFrom + "', 'yyyy.mm.dd'), to_date('" + dateTo + "', 'yyyy.mm.dd')) " +
                "from power_count_list where id_hi = " + groupId + " order by idname_lo";

        ResultSet resultSet = null;
        resultSet = statement.executeQuery(sql);

        HSSFSheet sheet = workbook.createSheet("Расчет");

        HSSFRow rowTitle = sheet.createRow((short)0);
        rowTitle.createCell(0).setCellValue("Расчет средней мощности");
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 18);
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        rowTitle.getCell(0).setCellStyle(cellStyle);

        HSSFRow rowName = sheet.createRow((short)1);
        rowName.createCell(0).setCellValue(groupName);
        rowName.getCell(0).setCellStyle(cellStyle);
        HSSFRow rowDate = sheet.createRow((short)2);
        rowDate.createCell(0).setCellValue(dateFrom + " - " + dateTo);
        rowDate.getCell(0).setCellStyle(cellStyle);

        HSSFRow rowHead = sheet.createRow((short)3);
        rowHead.createCell(0).setCellValue("ID");
        rowHead.createCell(1).setCellValue("Потребитель");
        rowHead.createCell(2).setCellValue("Ср. Мощность в часы контроля");
        rowHead.createCell(3).setCellValue("Кроме часов контроля");
        rowHead.createCell(4).setCellValue("Максимум в часы контроля");
        rowHead.createCell(5).setCellValue("Среднее - Все часы");
        rowHead.createCell(6).setCellValue("Максимум - Все часы");
        rowHead.createCell(7).setCellValue("Максимум - Включая нерабочие дни");

        int i = 4;
        while(resultSet.next()) {
            int id = resultSet.getInt(1);
            String idname = resultSet.getString(2);
            long value1 = resultSet.getLong(3);
            long value2 = resultSet.getLong(4);
            long value3 = resultSet.getLong(5);
            long value4 = resultSet.getLong(6);
            long value5 = resultSet.getLong(7);
            long value6 = resultSet.getLong(8);

            HSSFRow row = sheet.createRow((short) i);
            row.createCell(0).setCellValue(id);
            row.createCell(1).setCellValue(idname);
            row.createCell(2).setCellValue(value1);
            row.createCell(3).setCellValue(value2);
            row.createCell(4).setCellValue(value3);
            row.createCell(5).setCellValue(value4);
            row.createCell(6).setCellValue(value5);
            row.createCell(7).setCellValue(value6);

            i++;
        }

        sheet.autoSizeColumn(1);
    }

}
