package zyr.learn.poi.util;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

/**
 * Created by zhouweitao on 2016/12/4.
 */
public class ExcelUtil {
    private static ExcelUtil eu = new ExcelUtil();

    private ExcelUtil() {

    }

    public static ExcelUtil newInstance() {
        return eu;
    }


    public ExcelTemplate obj2Excel(Map map, String template, List objs, Class clazz, boolean isClasspath) {
        ExcelTemplate et = ExcelTemplate.newInstance();
        try {
            et = ExcelTemplate.newInstance();
            if (isClasspath) {
                et.readTemplateByClasspath(template);
            } else {
                et.readTemplateByPath(template);
            }
            List<ExcelHeader> headers = getHeaderList(clazz);
            Collections.sort(headers);
            et.createNewRow();
            /**
             * 创建表头
             * */
            for (ExcelHeader header : headers) {
                et.createCell(header.getTitle());
            }
            /**
             * 设置表头，表尾信息
             * */

            if (map != null) {
                et.replaceFinalData(map);
            }

            /**
             * 向表中插入值
             * */
            for (Object obj : objs) {
                et.createNewRow();
                for (ExcelHeader eh : headers) {
                    String mn = eh.getMethodName().substring(3);
                    mn = mn.substring(0, 1).toLowerCase() + mn.substring(1);
                    et.createCell(BeanUtils.getProperty(obj, mn));
                }
            }
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }
        return et;
    }

    private String getMethodName(ExcelHeader eh) {
        String mn = eh.getMethodName().substring(3);
        mn = mn.substring(0, 1).toLowerCase() + mn.substring(1);
        return mn;
    }

    /**
     * 输出流
     */
    public OutputStream exportObj2ExcelByTemplate(Map map, String template, OutputStream outputStream, List objs, Class clazz, boolean isClasspath) {
        ExcelTemplate et = obj2Excel(map, template, objs, clazz, isClasspath);
        et.writeToFIle(outputStream);
        return outputStream;
    }

    /**
     * 输出文件路径
     */
    public void exportObj2ExcelByTemplate(Map map, String template, String outPath, List objs, Class clazz, boolean isClasspath) {
        ExcelTemplate et = obj2Excel(map, template, objs, clazz, isClasspath);
        et.writeToFIle(outPath);
    }

    private List<ExcelHeader> getHeaderList(Class clazz) {
        List<ExcelHeader> headers = new ArrayList<ExcelHeader>();
        Method[] ms = clazz.getDeclaredMethods();
        for (Method m : ms) {
            String mn = m.getName();
            if (mn.startsWith("get")) {
                if (m.isAnnotationPresent(ExcelResources.class)) {
                    ExcelResources er = m.getAnnotation(ExcelResources.class);
                    headers.add(new ExcelHeader(er.title(), er.order(), mn));
                }
            }
        }
        return headers;
    }


    private Workbook obj2Excel(List objs, Class clazz, boolean isXssf) {
        Workbook wb = null;
        try {
            if (isXssf) {
                wb = new XSSFWorkbook();
            } else {
                wb = new HSSFWorkbook();
            }
            Sheet sheet = wb.createSheet();
            Row row = sheet.createRow(0);

            List<ExcelHeader> headers = getHeaderList(clazz);
            /**
             * 排序
             * */
            Collections.sort(headers);
            /**
             * 创建表头
             * */
            for (int i = 0; i < headers.size(); i++) {
                row.createCell(i).setCellValue(headers.get(i).getTitle());
            }
            /**
             * 向表中写数据
             * */
            Object obj = null;
            for (int i = 0; i < objs.size(); i++) {
                row = sheet.createRow(i + 1);
                for (int j = 0; j < headers.size(); j++) {
                    String mn = headers.get(j).getMethodName().substring(3);
                    mn = mn.substring(0, 1).toLowerCase() + mn.substring(1);
                    row.createCell(j).setCellValue(BeanUtils.getProperty(objs.get(i), mn));
                }
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
        return wb;
    }


    public void exportObj2Excel(String outPath, List objs, Class clazz, boolean isXssf) {
        Workbook wb = obj2Excel(objs, clazz, isXssf);
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(outPath);
            wb.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public void exportObj2Excel(OutputStream os, List objs, Class clazz, boolean isXssf) {
        try {
            Workbook wb = obj2Excel(objs, clazz, isXssf);
            wb.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 读 表中excel中数据
     */

    private Map<Integer, String> getHeaderMap(Row titleRow, Class clazz) {
        Map<Integer, String> maps = new HashMap<>();
        List<ExcelHeader> headers = getHeaderList(clazz);
        for (Cell cell : titleRow) {
            String title = cell.getStringCellValue();
            for (ExcelHeader eh : headers) {
                if (eh.getTitle().equals(title)) {
                    maps.put(cell.getColumnIndex(), eh.getMethodName().replace("get", "set"));
                    break;
                }
            }
        }
        return maps;
    }

    public List<Object> readExcel2ObjsByClassPath(String path, Class clazz, int readLine, int tailLine) {
        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(ExcelUtil.class.getResourceAsStream(path));
            return handlerExcel2Objs(wb, clazz, readLine, tailLine);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return null;
    }

    public List<Object> readExcel2ObjsByFilePath(String path, Class clazz, int readLine, int tailLine) {
        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(new File(path));
            return handlerExcel2Objs(wb, clazz, readLine, tailLine);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return null;
    }

    public List<Object> readExcel2ObjsByClassPath(String path, Class clazz) {
        return readExcel2ObjsByClassPath(path, clazz, 0, 0);
    }

    public List<Object> readExcel2ObjsByFilePath(String path, Class clazz) {
        return readExcel2ObjsByFilePath(path, clazz, 0, 0);
    }

    private List<Object> handlerExcel2Objs(Workbook wb, Class clazz, int readLine, int tailLine) {
        List<Object> objects = new ArrayList<>();
        try {
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(readLine);
            Map<Integer, String> maps = getHeaderMap(row, clazz);
            if (maps == null||maps.size()<=0) {
                throw new RuntimeException("excel格式不正确，检查是否设定了格式行！请检查，并设定读取范围");
            }
            for (int i = 0; i <= sheet.getLastRowNum() - tailLine; i++) {
                row = sheet.getRow(i);
                Object obj = clazz.newInstance();
                for (Cell c : row) {
                    int ci = c.getColumnIndex();
                    String mn = maps.get(ci).substring(3);
                    mn = mn.substring(0, 1).toLowerCase() + mn.substring(1);
                    BeanUtils.copyProperty(obj, mn, getCellValue(c));
                }
                objects.add(obj);
            }
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }
        return objects;
    }

    private String getCellValue(Cell c) {
        String s = null;
        switch (c.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                s = "";
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                s = String.valueOf(c.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                s = String.valueOf(c.getCellFormula());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                s = String.valueOf(c.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING:
                s = c.getStringCellValue();
                break;
            default:
                s = null;
                break;
        }
        return s;
    }
}
