package com.devdungeon.mysql2excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Handles the connection to the database and dumping the data to a .xlsx file.
 *
 * @author NanoDano <nanodano@devdungeon.com>
 */
public class MysqlDumper {


    /**
     * Processing date and time delta
     * @param outputFileName
     * @return replaced outputFileName string
     */
    static String fileNameProcess(String outputFileName) {
        SimpleDateFormat dateSdf = new SimpleDateFormat("yyyyMMdd");
        SimpleDateFormat timeSdf = new SimpleDateFormat("HHmmss");
        String datePattern = "~date([+-]?\\d+)~";
        String timePattern = "~time([+-]?\\d+)~";
        Pattern datePat = Pattern.compile(datePattern);
        Pattern timePat = Pattern.compile(timePattern);
        Matcher dateMatcher = datePat.matcher(outputFileName);
        Matcher timeMatcher = timePat.matcher(outputFileName);
        int dateDelta = 0;
        int timeDelta = 0;
        if (dateMatcher.find()) {
            dateDelta = Integer.parseInt(dateMatcher.group(1));
        }
        if (timeMatcher.find()) {
            timeDelta = Integer.parseInt(timeMatcher.group(1));
        }

        Calendar cal = Calendar.getInstance();
        cal.setTime(new Date());
        cal.add(Calendar.DAY_OF_MONTH, dateDelta);
        cal.add(Calendar.HOUR, timeDelta);

        String dateStr = dateSdf.format(cal.getTime());
        String timeStr = timeSdf.format(cal.getTime());

        String result = outputFileName.replaceAll(datePattern, dateStr).replaceAll(timePattern, timeStr);
        return result;
    }

    /**
     * Given database credentials and table name, this function connects to the
     * database and then dumps all of the data to a spreadsheet file including a
     * header row.
     *
     * @param outputFileName Name of output file (e.g. output.xlsx)
     * @param dbHost         Database host (e.g. localhost)
     * @param dbName         Name of database
     * @param dbUser         Database username
     * @param dbPass         Database user password
     * @param tableNames     Multi-table names in database
     * @param tableCondition Condition for all tables
     */
    static void dumpMysqlToExcelFile(String outputFileName, String dbHost, String dbName, String dbUser, String dbPass, String tableNames, String tableCondition) {
        // Create spreadsheet
        XSSFWorkbook workbook = new XSSFWorkbook();
        // Connect to database
        Connection conn = null;
        Statement stmt = null;
        try {
            conn = DriverManager.getConnection("jdbc:mysql://" + dbHost + ":3306/" + dbName + "?zeroDateTimeBehavior=convertToNull", dbUser, dbPass);
        } catch (SQLException ex) {
            Log.logException("Error connecting to database.", ex);
        }
        if (conn == null) {
            Log.logSevere("No connection established. Exiting.");
            System.exit(1);
        }
        Log.log("Connected.");

        String[] allTables = tableNames.replaceAll(" ", "").split(",");
        // Get list of column names and write sheet headers

        try {
            stmt = conn.createStatement();

            for (String tableName : allTables) {
                int cellNum = 0;
                ResultSet results = stmt.executeQuery("DESCRIBE " + tableName);
                ArrayList<String> columnNames = new ArrayList();
                XSSFSheet mySheet = workbook.createSheet(tableName);
                Row row = mySheet.createRow(0); // Header row
                while (results.next()) {
                    // Each column name in the table
                    Cell cell = row.createCell(cellNum++);
                    cell.setCellValue(results.getString(1));
                    columnNames.add(results.getString(1));
                }
                Log.log("Columns found: " + columnNames.size());
                Log.log(columnNames);

                // Get list of all data and dump
                Log.log("Dumping data...");
                int rowNum = 1;

                stmt = conn.createStatement();
                String querySql = "SELECT * FROM " + tableName;
                if (tableCondition.trim().length() > 0) {
                    querySql += " where " + tableCondition;
                }
                results = stmt.executeQuery(querySql);
                // write each field to db column and append row
                while (results.next()) {
                    row = mySheet.createRow(rowNum++); // 0-based index
                    int cellnum = 0;
                    for (String colName : columnNames) {
                        Cell cell = row.createCell(cellnum++);
                        cell.setCellValue(results.getString(colName));
                    }
                }
            }
        } catch (SQLException ex) {
            Log.logException("Error making SQL query to check table column headers.", ex);
        }
        if (stmt == null) {
            Log.logSevere("Statement is null when it shouldn't be. Exiting.");
            System.exit(1);
        }

        // Save spreadsheet
        FileOutputStream os;
        try {
            outputFileName = fileNameProcess(outputFileName);
            int lastSep = outputFileName.lastIndexOf(File.separator);

            String outputDir = null;
            if (lastSep > 0) {
                outputDir = outputFileName.substring(0, lastSep);
                File outputDirs = new File(outputDir);
                outputDirs.mkdirs();
                Log.log("Create directory: " + outputDir);
            }

            os = new FileOutputStream(outputFileName);
            workbook.write(os);
            Log.log("Writing on XLSX file Finished ...");
        } catch (IOException ex) {
            Log.logException("Error writing Excel file.", ex);
        }
        // Clean up db stuff
        try {
            stmt.close();
            conn.close();
        } catch (SQLException ex) {
            Log.logException("Error closing statement and database. Exiting", ex);
            System.exit(1);
        }
    }

}
