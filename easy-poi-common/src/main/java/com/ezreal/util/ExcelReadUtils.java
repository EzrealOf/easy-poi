package com.ezreal.util;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;

@Slf4j
@UtilityClass
public class ExcelReadUtils {

    /**
     * 数值后缀
     */
    private static final String NUMBER_SUFFIX = ".0";

    /**
     * 从第二行开始读取excel文件内容
     */
    public static List<List<Object>> getExcelContextByPath(String path) {
        return getExcelContextByPathAndStartNumber(path, 1, null);
    }

    /**
     * 根据路径和开始读取行获取excel文件内容
     *
     * @param path        文件路径
     * @param startNumber 从第几行开始读，如0、1、2、3
     * @return excel读取结果
     */
    public static List<List<Object>> getExcelContextByPathAndStartNumber(String path, int startNumber, Integer rowCountIndex) {
        rowCountIndex = Optional.ofNullable(rowCountIndex).orElseGet(() -> getRowNum(path));

        Workbook wb = getWorkbook(path);
        Sheet sheet = wb.getSheetAt(0);
        List<List<Object>> rowList = readRows(sheet, startNumber, rowCountIndex);

        if (rowList.size() == 0) {
            throw new NullPointerException("该文件中没有内容，请检查文件是否为空");
        }
        return rowList;
    }

    /**
     * 获取第一行内容
     *
     * @param path 文件路径
     * @return 第一行内容
     */
    public static List<Object> getContextByFirstRow(String path) {
        return getContextByHeadNumber(path, 0).get(0);
    }

    /**
     * 获取前面几行内容
     *
     * @param path          文件路径
     * @param rowCountIndex 总行数
     * @return 前几行内容
     */
    public static List<List<Object>> getContextByHeadNumber(String path, int rowCountIndex) {
        return getExcelContextByPathAndStartNumber(path, 0, rowCountIndex);
    }

    /**
     * 具体获取数据 3
     */
    public static List<List<Object>> readRows(Sheet sheet, int startRowIndex, int rowCountIndex) {
        List<List<Object>> rowList = new ArrayList<>();
        int totalRows = startRowIndex + rowCountIndex;
        int maxCellIndex = getMaxCellIndex(sheet);
        // 取的是下标，所以要+1
        for (int i = startRowIndex; i <= totalRows; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            List<Object> cellList = new ArrayList<>();
            // 取的是下标，所以要+1
            for (int j = 0; j <= maxCellIndex; j++) {
                Cell cell = row.getCell(j);
                Object cellValue;
                if (isMergedRegion(sheet, i, j)) {
                    cellValue = getMergedRegionValue(sheet, i, j);
                } else {
                    cellValue = readCell(cell);
                }
                cellList.add(cellValue == null ? "" : cellValue);
            }
            rowList.add(cellList);
        }
        return rowList;
    }

    /**
     * 获取excel文件中的所有sheet页
     */
    public static Map<String, List<List<Object>>> readAllRows(String excelFile) {
        try (FileInputStream is = new FileInputStream(excelFile)) {
            return readAllRows(is);
        } catch (Exception e) {
            throw new ServiceException(e);
        }
    }

    public static Map<String, List<List<Object>>> readAllRows(MultipartFile file) {
        try (InputStream is = file.getInputStream()) {
            return readAllRows(is);
        } catch (Exception e) {
            log.error("Excel解析失败", e);
            throw new ServiceException(e.getMessage());
        }
    }

    public static Map<String, List<List<Object>>> readAllRows(InputStream is) {
        Workbook wb = getWorkbook(is);
        Map<String, List<List<Object>>> map = Maps.newHashMap();
        // 获取每个Sheet表
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            List<List<Object>> rows = readRows(sheet);
            if (rows.size() == 0) {
                continue;
            }
            map.put(sheet.getSheetName(), rows);
        }
        if (map.size() == 0) {
            throw new ServiceException("请检查文件内容是否存在");
        }
        return map;
    }

    public static List<List<Object>> readRows(String excelFile) {
        try (FileInputStream is = new FileInputStream(excelFile)) {
            return readRows(is);
        } catch (Exception e) {
            throw new ServiceException(e);
        }
    }

    public static List<List<Object>> readRows(InputStream is) {
        Workbook wb = getWorkbook(is);
        Sheet sheet = wb.getSheetAt(0);
        return readRows(sheet);
    }

    /**
     * 获取当前sheet里的最大列数
     */
    public static int getMaxCellIndex(Sheet sheet) {
        int maxCellIndex = 0;
        // 取的是下标，所以要+1
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            short index = row.getLastCellNum();
            if (index > maxCellIndex) {
                maxCellIndex = index;
            }
        }
        return maxCellIndex;
    }

    public static List<List<Object>> readRows(Sheet sheet) {
        int rowCount = sheet.getLastRowNum();
        return readRows(sheet, 0, rowCount);
    }

    /**
     * 根据路径读取指定sheet页的Excel内容
     */
    public static List<List<Object>> readRowsByPathAndSheetName(String excelFile, String sheetName) {
        Workbook wb = getWorkbook(excelFile);
        Sheet sheet = Optional.ofNullable(wb.getSheet(sheetName)).orElseThrow(() -> new ServiceException(sheetName + "，该sheet页名称没找到"));
        return readRows(sheet);
    }

    /**
     * 读取行数
     */
    public static int getRowNum(String excelFile) {
        Workbook wb = getWorkbook(excelFile);
        Sheet sheet = wb.getSheetAt(0);
        // 取的是下标，所以要+1
        return sheet.getLastRowNum() + 1;
    }

    /**
     * 获取列数
     */
    public static int getColNum(String excelFile) {
        Workbook wb = getWorkbook(excelFile);
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(0);
        return row.getLastCellNum();
    }

    /**
     * 从Excel读Cell
     */
    private static Object readCell(Cell cell) {
        if (cell == null) {
            return null;
        }
        try {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    String str = cell.getRichStringCellValue().getString();
                    return StringUtils.trim(str);
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    }
                    return cell.getNumericCellValue();
                case Cell.CELL_TYPE_BOOLEAN:
                    return cell.getBooleanCellValue();
                case Cell.CELL_TYPE_FORMULA:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    }
                    return cell.getCellFormula();
                case Cell.CELL_TYPE_BLANK:
                    return "";
                case Cell.CELL_TYPE_ERROR:
                    return "";
                default:
                    System.err.println("Data error for cell of excel: " + cell.getCellType());
                    return "";
            }
        } catch (Exception e) {
            log.error("文件中可能包含公式", e);
            throw new ServiceException("文件中可能包含公式，请核查！");
        }
    }

    /**
     * 获取列值
     */
    public static String getCell(Object object) {
        if (object == null) {
            return null;
        }
        String str = String.valueOf(object).trim();
        if (str.isEmpty()) {
            return str;
        }
        return str;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     *
     * @param row    行下标
     * @param column 列下标
     */
    public static boolean isMergedRegion(Sheet sheet, int row, int column) {
        return (boolean) checkOrGetMergedRegion(sheet, row, column, true);
    }

    /**
     * 校验是否是合并单元格或者获取合并单元格内容
     *
     * @param sheet   sheet
     * @param row     行下标
     * @param column  列下标
     * @param isCheck 是否是校验
     * @return 校验结果或者合并单元格的值
     */
    private static Object checkOrGetMergedRegion(Sheet sheet, int row, int column, boolean isCheck) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    if (isCheck) {
                        return true;
                    }
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return readCell(fCell);
                }
            }
        }
        if (isCheck) {
            return false;
        }
        return null;
    }

    /**
     * 获取合并区域值
     */
    public static Object getMergedRegionValue(Sheet sheet, int row, int column) {
        return checkOrGetMergedRegion(sheet, row, column, false);
    }

    public static <T> List<T> parseFromExcel(MultipartFile uploadFile, Class<T> aimClass) {
        return parseFromExcel(uploadFile, 1, aimClass);
    }

    public static <T> List<T> parseFromExcel(String filePath, int firstIndex, Class<T> aimClass) {
        return parseFromExcel(new File(filePath), firstIndex, aimClass);
    }

    /**
     * 根据Excel输入流转换成对象列表
     *
     * @param uploadFile 上传上来的Excel文件
     * @param firstIndex 起始行
     * @param aimClass   对象class
     * @param <T>        对象
     * @return 对象列表
     */
    public static <T> List<T> parseFromExcel(MultipartFile uploadFile, int firstIndex, Class<T> aimClass) {
        return parseObjectFromExcelInputStream(getWorkbook(uploadFile), firstIndex, aimClass);
    }

    /**
     * 根据Excel输入流转换成对象列表
     *
     * @param file       需要解析的Excel文件
     * @param firstIndex 起始行
     * @param aimClass   对象class
     * @param <T>        对象
     * @return 对象列表
     */
    public static <T> List<T> parseFromExcel(File file, int firstIndex, Class<T> aimClass) {
        return parseObjectFromExcelInputStream(getWorkbook(file), firstIndex, aimClass);
    }

    /**
     * 获取工作簿
     */
    public static Workbook getWorkbook(MultipartFile uploadFile) {
        try (InputStream is = uploadFile.getInputStream()) {
            return WorkbookFactory.create(is);
        } catch (Exception e) {
            log.error("Excel转换失败", e);
            throw new ServiceException(e.getMessage());
        }
    }

    /**
     * 获取工作簿
     */
    public static Workbook getWorkbook(File file) {
        try (InputStream is = new FileInputStream(file)) {
            return WorkbookFactory.create(is);
        } catch (Exception e) {
            log.error("Excel转换失败", e);
            throw new ServiceException(e.getMessage());
        }
    }

    /**
     * 获取工作簿
     */
    public static Workbook getWorkbook(String excelFile) {
        return getWorkbook(new File(excelFile));
    }

    /**
     * 获取工作簿
     */
    public static Workbook getWorkbook(InputStream inputStream) {
        try (InputStream is = inputStream) {
            return WorkbookFactory.create(is);
        } catch (Exception e) {
            log.error("Excel转换失败", e);
            throw new ServiceException(e.getMessage());
        }
    }

    /**
     * 根据Excel输入流转换成对象列表
     *
     * @param workbook   工作浦
     * @param firstIndex 起始行
     * @param aimClass   对象class
     * @param <T>        对象
     * @return 对象列表
     */
    public static <T> List<T> parseObjectFromExcelInputStream(Workbook workbook, int firstIndex, Class<T> aimClass) {
        // 对excel文档的第一页,即sheet1进行操作
        Sheet sheet = workbook.getSheetAt(0);
        int lastRaw = sheet.getLastRowNum();
        List<T> result = Lists.newArrayList();
        for (int i = firstIndex; i <= lastRaw; i++) {
            result.add(getEntityByRow(sheet.getRow(i), aimClass));
        }
        if (0 == result.size()) {
            throw new ServiceException("没有从文件中读取到内容，请核查！");
        }
        return result;
    }

    public static <T> T getEntityByRow(Row row, Class<T> ct) {
        return getEntityByObjectRow(row, ct);
    }

    public static <T> T getEntityByRow(List row, Class<T> ct) {
        return getEntityByObjectRow(row, ct);
    }

    public static <T> T getEntityByRow(Map row, Class<T> ct) {
        return getEntityByObjectRow(row, ct);
    }

    private static <T> T getEntityByObjectRow(Object row, Class<T> ct) {
        T entity;
        try {
            entity = ct.newInstance();
        } catch (Exception e) {
            log.error(ct.getSimpleName() + "实例化失败！", e);
            throw new ServiceException(ct.getSimpleName() + "实例化失败！" + e.getMessage());
        }

        List<String> keys = null;
        List<String> values = null;
        if (row instanceof Map) {
            @SuppressWarnings("unchecked")
            Map<String, String> map = (Map<String, String>) row;
            keys = Lists.newArrayList(map.keySet());
            values = Lists.newArrayList(map.values());
        }

        Field[] fields = ct.getDeclaredFields();
        for (int j = 0, ii = 0; j < fields.length; j++) {
            Object value = null;
            if (row instanceof Map) {
                String cellPrefix = getExcelCellPrefixByIndex(j);
                if (!keys.get(ii).startsWith(cellPrefix)) {
                    continue;
                }
                value = values.get(ii);
                ii++;
            } else if (row instanceof List) {
                value = ((List) row).get(j);
            } else if (row instanceof Row) {
                value = readCell(((Row) row).getCell(j));
            }

            if (null == value || StringUtils.isBlank(value.toString())) {
                continue;
            }
            Field field = fields[j];
            ReflectionUtil.setFieldValue(entity, field.getName(), initValueByType(value.toString(), field.getType()));
        }
        return entity;
    }

    /**
     * 移除错误标志
     *
     * @param str 字符串
     * @return 移除后的字符串
     */
    public static String removeErrorSign(String str) {
        if (StringUtils.isBlank(str)) {
            throw new ServiceException("字符串不能为空");
        }
        if (str.endsWith(NUMBER_SUFFIX)) {
            str = str.substring(0, str.length() - NUMBER_SUFFIX.length());
        }
        return str;
    }

    /**
     * 根据类型初始化值
     *
     * @param str       字符串
     * @param typeClass 字段类型
     * @return 对象
     */
    public static Object initValueByType(String str, Class<?> typeClass) {
        if (StringUtils.isBlank(str)) {
            return null;
        }
        str = StringUtils.trim(str);

        Object value;
        if (typeClass.equals(int.class) || typeClass.equals(Integer.class)) {
            try {
                value = Integer.parseInt(removeErrorSign(str));
            } catch (Exception e) {
                throw new ServiceException(str + "该值不能转换成Integer类型");
            }
        } else if (typeClass.equals(long.class) || typeClass.equals(Long.class)) {
            try {
                value = Long.parseLong(removeErrorSign(str));
            } catch (Exception e) {
                throw new ServiceException(str + "该值不能转换成Long类型");
            }
        } else if (typeClass.equals(double.class) || typeClass.equals(Double.class)) {
            try {
                value = Double.parseDouble(str);
            } catch (Exception e) {
                throw new ServiceException(str + "该值不能转换成Double类型");
            }
        } else if (typeClass.equals(float.class) || typeClass.equals(Float.class)) {
            try {
                value = Float.parseFloat(str);
            } catch (Exception e) {
                throw new ServiceException(str + "该值不能转换成Float类型");
            }
            //} else if (typeClass.equals(Date.class)) {
            //value = DateFormatUtil.parseDate(str);
            //if (null == value) {
            //    throw new ServiceException(str + "该值不能转换成Date类型");
            //}
        } else if (typeClass.equals(BigDecimal.class)) {
            try {
                value = new BigDecimal(str);
            } catch (Exception e) {
                throw new ServiceException(str + "该值不能转换成BigDecimal类型");
            }
        } else {
            value = str;
        }
        return value;
    }

    /**
     * 读取Excel文件
     *
     * @param filename 文件路径及名称
     * @param handler  中间处理
     */
    public static void readExcelByFile(String filename, XSSFSheetXMLHandler.SheetContentsHandler handler) {
        try (OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ)) {
            XSSFReader xssfReader = new XSSFReader(pkg);
            try (InputStream is = xssfReader.getSheetsData().next()) {
                StylesTable styles = xssfReader.getStylesTable();
                ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
                processSheet(styles, strings, is, handler);
            }
        } catch (ServiceException e) {
            throw e;
        } catch (Exception e) {
            throw new ServiceException(e);
        }
    }

    /**
     * sheet处理
     */
    private static void processSheet(StylesTable styles, ReadOnlySharedStringsTable strings, InputStream sheetInputStream, XSSFSheetXMLHandler.SheetContentsHandler handler) throws SAXException, ParserConfigurationException, IOException {
        XMLReader sheetParser = SAXHelper.newXMLReader();

        if (handler != null) {
            sheetParser.setContentHandler(new XSSFSheetXMLHandler(styles, strings, handler, false));
        } else {
            sheetParser.setContentHandler(new XSSFSheetXMLHandler(styles, strings, new SimpleSheetContentsHandler(), false));
        }

        sheetParser.parse(new InputSource(sheetInputStream));
    }

    /**
     * 根据下标生成excel列前缀
     *
     * @param index 下标
     * @return 列前缀
     */
    public static String getExcelCellPrefixByIndex(int index) {
        int max = 26;
        int begin = 65;
        if (index >= max) {
            String colName = getExcelCellPrefixByIndex(index / max - 1);
            colName += (char) (begin + index % max);
            return colName;
        }
        return String.valueOf((char) (begin + index));
    }

    /**
     * sheet处理对象
     */
    public static class SimpleSheetContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

        private Map<String, String> rowMap;

        public Map<String, String> getRowMap() {
            return rowMap;
        }

        @Override
        public void startRow(int rowNum) {
            rowMap = Maps.newLinkedHashMap();
        }

        @Override
        public void endRow(int rowNum) {
            log.info(rowNum + " : " + rowMap);
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            if (null != formattedValue) {
                formattedValue = formattedValue.trim();
            }

            rowMap.put(cellReference, formattedValue);
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
            log.info(text + "\t" + isHeader + "\t" + tagName);
        }
    }
}
