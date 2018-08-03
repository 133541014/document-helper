package pers.chao.document.helper.excel;

import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import pers.chao.document.helper.annontation.ExcelColumn;
import pers.chao.document.helper.common.Consts;
import pers.chao.document.helper.exception.DataSetException;
import pers.chao.document.helper.exception.ExcelHandleException;
import pers.chao.document.helper.exception.FileNameException;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author: WangYichao
 * @description: ExcelUtil
 * @date: 2018/8/3 22:19
 */
public class ExcelUtils {

    private static final Logger log = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 根据数据集导出excel
     *
     * @param dataset  数据集
     * @param fileName 导出文件名
     * @param <T>      实体类型
     */
    public static <T> void exportForAnno(List<T> dataset, String fileName, HttpServletResponse response) throws ExcelHandleException, IOException {

        if (dataset == null || dataset.size() == 0) {
            throw new DataSetException("数据集为空，没有可导出数据");
        } else if (fileName == null || fileName == "") {
            throw new FileNameException("文件名为空");
        }

        //获取标题
        Class<?> type = dataset.get(0).getClass();

        //获取表头列名
        Field[] fields = type.getDeclaredFields();
        List<String> keyList = new ArrayList<>();
        List<String> columnList = new ArrayList<>();
        List<Integer> orderList = new ArrayList<>();
        String key = null;
        String column = null;
        for (Field field : fields) {

            field.setAccessible(true);

            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);

            if (excelColumn == null) {
                continue;
            }

            key = field.getName();
            column = excelColumn.value();
            orderList.add(excelColumn.order());
            keyList.add(key);
            columnList.add(column);

        }

        //根据排序号进行排序
        sortColumn(keyList, columnList, orderList);

        String[] keys = new String[keyList.size()];
        String[] columnNames = new String[columnList.size()];
        keyList.toArray(keys);
        columnList.toArray(columnNames);

        if (columnNames.length == 0) {
            throw new DataSetException("没有需要导出的列");
        }

        exportForArray(dataset, keys, columnNames, fileName, response);

    }

    /**
     * 导出excel
     *
     * @param list        数据集合
     * @param keys        表头
     * @param columnNames 列名
     * @param fileName    文件名
     * @param response    response
     * @throws IOException IOException
     */
    public static void exportForArray(List list, String[] keys, String[] columnNames, String fileName, HttpServletResponse response) throws IOException {
        try {
            Workbook wb = createWorkBook(list, keys, columnNames);
            writeWorkBook(fileName, response, wb);
        } catch (IOException e) {
            log.error("导出excel出错!", e);
        }

    }

    /**
     * 创建excel文档
     *
     * @param list        数据
     * @param keys        list中map的key数组集合
     * @param columnNames excel的列
     */
    private static Workbook createWorkBook(List<Map<String, Object>> list, String[] keys, String[] columnNames) {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        for (int i = 0; i < keys.length; i++) {
            sheet.setColumnWidth((short) i, (short) (Consts.NUM_40 * Consts.NUM_120));
        }
        Row row = sheet.createRow((short) 0);
        CellStyle cs = wb.createCellStyle();
        CellStyle cs2 = wb.createCellStyle();
        Font f = wb.createFont();
        Font f2 = wb.createFont();
        f.setFontHeightInPoints((short) Consts.NUM_10);
        f.setColor(IndexedColors.BLACK.getIndex());
        f2.setFontHeightInPoints((short) Consts.NUM_10);
        for (int i = 0; i < columnNames.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(columnNames[i]);
            cell.setCellStyle(cs);
        }
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        for (short i = 0; i < list.size(); i++) {
            // Row Cell 方格 , Row Cell 都是计数
            // 创建，在页sheet
            Row row1 = sheet.createRow(i + 1);
            // 在row行上创建方格
            for (short j = 0; j < keys.length; j++) {
                Cell cell = row1.createCell(j);
                if (list.get(i).get(keys[j]) instanceof Date) {
                    cell.setCellValue(sdf.format(list.get(i).get(keys[j])));
                } else {
                    cell.setCellValue(list.get(i).get(keys[j]) == null ? " " : list.get(i).get(keys[j]).toString());
                }
                cell.setCellStyle(cs2);
            }
        }
        return wb;
    }

    /**
     * 内容写入工作薄
     *
     * @param sheetName sheet页名称
     * @param response  response
     * @param wb        工作薄
     * @throws IOException IOException
     */
    private static void writeWorkBook(String sheetName, HttpServletResponse response, Workbook wb) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {

            wb.write(os);
            byte[] content = os.toByteArray();
            InputStream is = new ByteArrayInputStream(content);

            response.reset();
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename="
                    + new String(sheetName.getBytes("GB2312"), "ISO8859-1") + ".xls");

            ServletOutputStream out = response.getOutputStream();
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(out);
            byte[] buff = new byte[Consts.NUM_2048];
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }

        } catch (IOException e) {
            log.error("数据写入excel出错!", e);
        } finally {
            if (bis != null) {
                bis.close();
            }
            if (bos != null) {
                bos.close();
            }
        }
    }

    /**
     * 总行数
     */
    private int totalRows = 0;
    /**
     * 总列数
     */
    private int totalCells = 0;

    public Map<String, List<List<String>>> read(InputStream inputStream, String fileName) {
        Map<String, List<List<String>>> maps = new HashMap<>(16);
        if (fileName == null || !fileName.matches("^.+\\.(?i)((xls)|(xlsx))$")) {
            return maps;
        }
//        File file = new File(fileName);
        if (inputStream == null) {
            return maps;
        }
        try {
            Workbook wb = WorkbookFactory.create(inputStream);
            maps = read(wb);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return maps;
    }

    public int getTotalRows() {
        return totalRows;
    }

    public int getTotalCells() {
        return totalCells;
    }

    private Map<String, List<List<String>>> read(Workbook wb) {
        Map<String, List<List<String>>> maps = new HashMap<>(16);
        int number = wb.getNumberOfSheets();
        if (number > 0) {
            // 循环每个工作表
            for (int i = 0; i < number; i++) {
                List<List<String>> list = new ArrayList<>();
                // 第一页去除行数
                int delnumber = 0;
                Sheet sheet = wb.getSheetAt(i);
                // 获取工作表中行数
                this.totalRows = sheet.getPhysicalNumberOfRows() - delnumber;
                if (this.totalRows >= 1 && sheet.getRow(delnumber) != null) {
                    // 得到当前行的所有单元格
                    this.totalCells = sheet.getRow(0)
                            .getPhysicalNumberOfCells();
                    for (int j = 0; j < totalRows; j++) {
                        List<String> rowLst = new ArrayList<>();
                        for (int f = 0; f < totalCells; f++) {
                            if (totalCells > 0) {
                                String value = getCell(sheet.getRow(j).getCell(f));
                                rowLst.add(value);
                            }
                        }
                        list.add(rowLst);
                    }
                }
                maps.put(sheet.getSheetName(), list);
            }
        }
        return maps;
    }

    private String getCell(Cell cell) {
        String cellValue = null;
        /*
         * if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) { if
         * (HSSFDateUtil.isCellDateFormatted(cell)) { cellValue =
         * getRightStr(cell.getDateCellValue() + ""); } else {
         *
         * cellValue = getRightStr(cell.getNumericCellValue() + ""); } } else if
         * (Cell.CELL_TYPE_STRING == cell.getCellType()) { cellValue =
         * cell.getStringCellValue(); } else if (Cell.CELL_TYPE_BOOLEAN ==
         * cell.getCellType()) { cellValue = cell.getBooleanCellValue() + ""; }
         * else { cellValue = cell.getStringCellValue(); }
         */
        HSSFDataFormatter hSSFDataFormatter = new HSSFDataFormatter();
        // 使用EXCEL原来格式的方式取得值
        cellValue = hSSFDataFormatter.formatCellValue(cell);
        return cellValue;
    }

    /**
     * 列排序
     *
     * @param keys        字段名
     * @param columnNames 列名
     * @param orders      排序号
     */
    private static void sortColumn(List<String> keys, List<String> columnNames, List<Integer> orders) {
        Map<String, String> key2Column = new HashMap<>();
        for (int i = 0; i < keys.size(); i++) {
            key2Column.put(keys.get(i), columnNames.get(i));
        }

        Map<String, Integer> orderTree = new HashMap<>();
        for (int i = 0; i < keys.size(); i++) {
            orderTree.put(keys.get(i), orders.get(i));
        }

        List<Map.Entry<String, Integer>> list = new ArrayList<>(orderTree.entrySet());

        Collections.sort(list, ((o1, o2) -> {
            return o1.getValue().compareTo(o2.getValue());
        }));

        keys.clear();
        columnNames.clear();
        for (Map.Entry<String, Integer> entry : list) {
            keys.add(entry.getKey());
            columnNames.add(key2Column.get(entry.getKey()));
        }

    }

}
