package zyr.learn.poi.util;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by zhouweitao on 2016/12/4.
 */
public class ExcelTemplate {
    public final static String DATA_LINE = "datas";
    public final static String DEFAULT_STYLE = "defaultStyles";
    public final static String STYLE = "styles";
    public final static String NO = "no";

    private static ExcelTemplate et = new ExcelTemplate();


    private Workbook wb;
    private Sheet sheet;

    /**
     * 初始化列数
     */
    private int initColIndex;
    /**
     * 初始化行数
     */
    private int initRowIndex;
    /**
     * 当前列数
     */
    private int curColIndex;
    /**
     * 当前行数
     */
    private int curRowIndex;

    /**
     * 当前行对象
     */
    private Row curRow;

    /**
     * 最后一行的数据
     */
    private int lastRowIndex;

    /**
     * 默认样式
     */
    private CellStyle defaultStyle;

    /**
     * 默认行高
     */
    private float rowHeight;

    /**
     * 存储每一行对应的样式
     */
    private Map<Integer, CellStyle> styles;

    /**
     * 序号的列
     */
    private int noColIndex;

    private ExcelTemplate() {

    }

    public static ExcelTemplate newInstance() {
        return et;
    }

    /**
     * 文件相对路径
     * */
    public ExcelTemplate readTemplateByClasspath(String path) {
        try {
            wb = WorkbookFactory.create(ExcelTemplate.class.getResourceAsStream(path));
            initTemplate();
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("读取的模板不存在！");
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            throw new RuntimeException("读取的模板格式错误！");
        }
        return this;
    }


    /**
     * 文件绝对路径
     * */
    public ExcelTemplate readTemplateByPath(String path) {
        try {
            wb = WorkbookFactory.create(new File(path));
            initTemplate();
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("读取的模板不存在！");
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            throw new RuntimeException("读取的模板格式错误！");
        }
        return this;
    }

    private void initTemplate() {
        sheet = wb.getSheetAt(0);
        initConfigData();
        lastRowIndex = sheet.getLastRowNum();
        curRow = sheet.createRow(curRowIndex);
    }

    /***
     *
     * 初始化数据
     * */
    private void initConfigData() {
        boolean flag = false;
        boolean findNo = false;
        for (Row row : sheet) {
            if (flag) break;
            for (Cell cell : row) {
                if (cell.getCellType() != cell.CELL_TYPE_STRING) continue;
                String str = cell.getStringCellValue().trim();
                if (str.equals(NO)) {
                    noColIndex = cell.getColumnIndex();
                    findNo = true;
                }
                if (str.equals(DATA_LINE)) {
                    initColIndex = cell.getColumnIndex();
                    initRowIndex = row.getRowNum();
                    curColIndex = initColIndex;
                    curRowIndex = initRowIndex;
                    defaultStyle = cell.getCellStyle();
                    rowHeight = row.getHeightInPoints();
                    initStyles();
                    flag = true;
                    break;
                }
            }
        }
        if (!findNo) {
            initNO();
        }
    }


    private void initStyles() {
        styles = new HashMap<Integer, CellStyle>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() != cell.CELL_TYPE_STRING) continue;
                String str = cell.getStringCellValue();
                if (str.equals(DEFAULT_STYLE)) {
                    defaultStyle = cell.getCellStyle();
                }
                if (str.equals(STYLE)) {
                    styles.put(cell.getColumnIndex(), cell.getCellStyle());
                }
            }
        }
    }

    public void createCell(String value) {
        Cell cell = curRow.createCell(curColIndex);
        setCellStyle(cell);
        cell.setCellValue(value);
        curColIndex++;
    }

    public void createCell(int value) {
        Cell cell = curRow.createCell(curColIndex);
        setCellStyle(cell);
        cell.setCellValue((int) value);
        curColIndex++;
    }

    public void createCell(Date value) {
        Cell cell = curRow.createCell(curColIndex);
        setCellStyle(cell);
        cell.setCellValue(value);
        curColIndex++;
    }

    public void createCell(double value) {
        Cell cell = curRow.createCell(curColIndex);
        setCellStyle(cell);
        cell.setCellValue(value);
        curColIndex++;
    }

    public void createCell(boolean value) {
        Cell cell = curRow.createCell(curColIndex);
        setCellStyle(cell);
        cell.setCellValue(value);
        curColIndex++;
    }

    public void createCell(Calendar value) {
        Cell cell = curRow.createCell(curColIndex);
        setCellStyle(cell);
        cell.setCellValue(value);
        curColIndex++;
    }

    public void createNewRow() {
        if (lastRowIndex > curRowIndex && curRowIndex != initRowIndex) {
            sheet.shiftRows(curRowIndex, lastRowIndex, 1, true, true);
            lastRowIndex++;
        }
        curRow = sheet.createRow(curRowIndex);
        curRow.setHeightInPoints(rowHeight);
        curRowIndex++;
        curColIndex = initColIndex;
    }

    public void writeToFIle(String filePath) {
        FileOutputStream fos = null;

        try {
            fos = new FileOutputStream(filePath);
            wb.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new RuntimeException("写入的文件不存在！" + e.getMessage());
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("写入数据失败失败！" + e.getMessage());
        } finally {
            try {
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    public void writeToFIle(OutputStream outputStream) {
        try {
            wb.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("写入流失败！" + e.getMessage());
        }
    }

    /**
     * 根据map替换相应常量
     */
    public void replaceFinalData(Map<String, String> datas) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() != cell.CELL_TYPE_STRING) continue;
                String str = cell.getStringCellValue();
                if (str.startsWith("#")) {
                    if (datas.containsKey(str.substring(1))) {
                        cell.setCellValue(datas.get(str.substring(1)));
                    }
                }
            }
        }
    }

    /**
     * 插入序号
     */

    public void insertNO() {
        int index = 1;
        Row row = null;
        Cell cell = null;
        for (int i = initRowIndex; i < curRowIndex; i++) {
            row = sheet.getRow(i);
            cell = row.createCell(noColIndex);
            setCellStyle(cell);
            cell.setCellValue(index++);
        }
    }

    private void setCellStyle(Cell cell) {
        if (styles.containsKey(curColIndex)) {
            cell.setCellStyle(styles.get(curColIndex));
        } else {
            cell.setCellStyle(defaultStyle);
        }
    }

    private void initNO() {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() != cell.CELL_TYPE_STRING) continue;
                String str = cell.getStringCellValue().trim();
                if (str.equals(NO)) {
                    noColIndex = cell.getColumnIndex();
                }
            }
        }
    }
}
