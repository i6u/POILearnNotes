package poi;

/**
 * Created by zhouweitao on 2016/12/3.
 */

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.text.Format;
import java.util.Date;

public class POITest {

    /**
     * 读取excel文档
     */
    @Test
    public void test01() throws Exception {
        Workbook workbook = WorkbookFactory.create(new File("/Users/zhouweitao/Desktop/temp/1.xls"));
        Sheet sheet = workbook.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();
        int firstRowNum = sheet.getFirstRowNum();
        System.out.println(firstRowNum + "--" + lastRowNum);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue());

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            for (int j = 0; j <row.getLastCellNum() ; j++) {
                System.out.print(row.getCell(j) + "\t");
            }
            System.out.println();
        }

        System.out.println("----------第二种方式----------");

        for (Row row1 : sheet) {
            for (Cell cell1 : row1) {
                System.out.print(cell1+"\t");
            }
            System.out.println();
        }
    }


    @Test
    public void test02() {
        FileOutputStream file = null;
        Workbook workbook = new HSSFWorkbook();
//        new XSSFWorkbook();

        try {
            file = new FileOutputStream("/Users/zhouweitao/Desktop/temp/2112.xls");
            Sheet sheet = workbook.createSheet("测试01");
            Row row = sheet.createRow(0);
            row.setHeightInPoints(30);


            CellStyle style = workbook.createCellStyle();
            style.setBorderBottom(BorderStyle.DOUBLE);
            style.setBottomBorderColor(HSSFColor.PINK.index);

            Cell cell = row.createCell(0);
            cell.setCellValue("部门编号");
            cell.setCellStyle(style);

            CellStyle style1 = workbook.createCellStyle();
            style1.setBorderBottom(BorderStyle.DOTTED);
            style1.setBottomBorderColor(HSSFColor.ORANGE.index);

            Cell cell1 = row.createCell(1);
            cell1.setCellValue("部门名称");
            cell1.setCellStyle(style1);

            workbook.write(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                file.close();
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    /***
     *
     * 网上的例子
     * 以下第二种方式可以读出完整的表格内容
     * */

    @Test
    public void demo() throws Exception {
        String path = "/Users/zhouweitao/Desktop/temp/1.xls";
        extract(path);
        readWorkbook(path);
    }



    /** * 直接抽取excel中的数据 */
    public static void extract(String path) {
        InputStream inp = null;
        Workbook workbook = null;
        ExcelExtractor extractor = null;
        XSSFExcelExtractor xssfExtractor = null;
        String text = "";
        try {
            inp = new FileInputStream(path);
            workbook = WorkbookFactory.create(inp);
            if (workbook instanceof HSSFWorkbook) {
                extractor = new ExcelExtractor((HSSFWorkbook) workbook);
                extractor.setFormulasNotResults(true);
                extractor.setIncludeSheetNames(false);
                text = extractor.getText();
            } else if (workbook instanceof XSSFWorkbook) {
                xssfExtractor = new XSSFExcelExtractor((XSSFWorkbook) workbook);
                xssfExtractor.setFormulasNotResults(true);
                xssfExtractor.setIncludeSheetNames(false);
                text = xssfExtractor.getText();
            } else {
                return;
            }
            System.out.println(text);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (EncryptedDocumentException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (extractor != null) {
                try {
                    extractor.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (xssfExtractor != null) {
                try {
                    xssfExtractor.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (inp != null) {
                try {
                    inp.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }



    /** * 原样返回数值单元格的内容 */
    public static String formatNumericCell(Double value, Cell cell) {
        if(cell.getCellTypeEnum() != CellType.NUMERIC && cell.getCellTypeEnum() != CellType.FORMULA) {
            return null;
        }
        //isCellDateFormatted判断该单元格是"时间格式"或者该"单元格的公式算出来的是时间格式"
        if(DateUtil.isCellDateFormatted(cell)) {
            //cell.getDateCellValue()碰到单元格是公式,会自动计算出Date结果
            Date date = cell.getDateCellValue();
            DataFormatter dataFormatter = new DataFormatter();
            Format format = dataFormatter.createFormat(cell);
            return format.format(date);
        } else {
            DataFormatter dataFormatter = new DataFormatter();
            Format format = dataFormatter.createFormat(cell);
            return format.format(value);

        }
    }


    public static void readWorkbook(String path) {
        InputStream inp = null;
        Workbook workbook = null;
        try {
            inp = new FileInputStream(path);
            workbook = WorkbookFactory.create(inp);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            // for/in是iterator的简写, 最终会被编译器编译为iterator
            for (Sheet sheet : workbook) {
                System.out.println("----------" + sheet.getSheetName() + "----------");
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        switch (cell.getCellTypeEnum()) {
                            case _NONE:
                                System.out.print("_NONE" + "\t");
                                break;
                            case BLANK:
                                System.out.print("BLANK" + "\t");
                                break;
                            case BOOLEAN:
                                System.out.print(cell.getBooleanCellValue() + "\t");
                                break;
                            case ERROR:
                                System.out.print("ERROR(" + cell.getErrorCellValue() + ")" + "\t");
                                break;
                            case FORMULA:
                                // 会打印出原本单元格的公式
                                // System.out.print(cell.getCellFormula() + "\t");
                                // NumberFormat nf = new DecimalFormat("#.#");
                                // String value = nf.format(cell.getNumericCellValue());
                                CellValue cellValue = evaluator.evaluate(cell);
                                switch (cellValue.getCellTypeEnum()) {
                                    case _NONE:
                                        System.out.print("_NONE" + "\t");
                                        break;
                                    case BLANK:
                                        System.out.print("BLANK" + "\t");
                                        break;
                                    case BOOLEAN:
                                        System.out.print(cellValue.getBooleanValue() + "\t");
                                        break;
                                    case ERROR:
                                        System.out.print("ERROR(" + cellValue.getErrorValue() + ")" + "\t");
                                        break;
                                    case NUMERIC:
                                        System.out.print(formatNumericCell(cellValue.getNumberValue(), cell) + "\t");
                                        break;
                                    case STRING:
                                        System.out.print(cell.getStringCellValue() + "\t");
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            case NUMERIC:
                                System.out.print(formatNumericCell(cell.getNumericCellValue(), cell) + "\t");
                                break;
                            case STRING:
                                System.out.print(cell.getStringCellValue() + "\t");
                                break;
                        }
                    }
                    System.out.println();
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (EncryptedDocumentException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (inp != null) {
                try {
                    inp.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }


}
