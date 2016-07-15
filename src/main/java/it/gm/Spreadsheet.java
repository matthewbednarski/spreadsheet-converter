package it.gm;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Created by matthew on 14.07.16.
 */
public class Spreadsheet {
    Logger logger = LoggerFactory.getLogger(this.getClass());

    public Map<String, String> convertToCsv(Path spreadsheet, char separator){
        if(Files.exists(spreadsheet)) {
            Workbook wb = null;
            try (InputStream is = new FileInputStream(spreadsheet.toFile())){
                wb = WorkbookFactory.create(is);
                return this.convertToCsv(wb, separator);
            } catch (InvalidFormatException e) {
                logger.warn(e.getLocalizedMessage());
                if (logger.isDebugEnabled()) {
                    logger.debug(e.getLocalizedMessage(), e);
                }
            } catch (FileNotFoundException e) {
                logger.warn(e.getLocalizedMessage());
                if (logger.isDebugEnabled()) {
                    logger.debug(e.getLocalizedMessage(), e);
                }
            } catch (IOException e) {
                logger.warn(e.getLocalizedMessage());
                if (logger.isDebugEnabled()) {
                    logger.debug(e.getLocalizedMessage(), e);
                }
            }
        }
        return null;
    }
    protected Map<String, String> convertToCsv(Workbook workbook, char separator) {
        Map<String, String> result = new HashMap<>();
        try {
            if (workbook != null && workbook.getNumberOfSheets() > 0) {
                for (int indexSheet = 0; indexSheet < workbook.getNumberOfSheets(); indexSheet++) {
                    Sheet sheet = workbook.getSheetAt(indexSheet);
                    StringBuffer data = new StringBuffer();
                    if (sheet != null) {
                        String name = sheet.getSheetName();
                        Row row;
                        Cell cell;
                        // Iterate through each rows from first sheet
                        Iterator<Row> rowIterator = sheet.iterator();
                        while (rowIterator.hasNext()) {
                            row = rowIterator.next();
                            for (int cellCount = 0; cellCount < row.getLastCellNum(); cellCount++) {
                                cell = row.getCell(cellCount, Row.CREATE_NULL_AS_BLANK);
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_ERROR:
                                        data.append(cell.getErrorCellValue());
                                        data.append(separator);
                                        break;
                                    case Cell.CELL_TYPE_FORMULA:
                                        data.append(cell.getCachedFormulaResultType());
                                        data.append(separator);
                                        break;
                                    case Cell.CELL_TYPE_BOOLEAN:
                                        data.append(cell.getBooleanCellValue());
                                        data.append(separator);
                                        break;
                                    case Cell.CELL_TYPE_NUMERIC:
                                        data.append(cell.getNumericCellValue());
                                        data.append(separator);
                                        break;
                                    case Cell.CELL_TYPE_STRING:
                                        data.append(cell.getStringCellValue());
                                        data.append(separator);
                                        break;
                                    case Cell.CELL_TYPE_BLANK:
                                        data.append("");
                                        data.append(separator);
                                        break;
                                    default:
                                        data.append(cell);
                                        data.append(separator);
                                }
                            }
                            data.append(System.lineSeparator());
                        }
                        result.put(name, data.toString());
                    }
                }
            }
        } catch (Exception e) {
            logger.warn(e.getLocalizedMessage());
            if (logger.isDebugEnabled()) {
                logger.debug(e.getLocalizedMessage(), e);
            }
        }
        return result;
    }
}
