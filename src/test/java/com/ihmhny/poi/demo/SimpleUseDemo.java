package com.ihmhny.poi.demo;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;


public class SimpleUseDemo {

    @Test
    public void testNewWorkbook() throws Exception {
        Workbook wb = new HSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testXSSFWorkbook() throws Exception {
        Workbook wb = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testSheet() throws Exception {
        Workbook wb = new HSSFWorkbook();
        wb.createSheet("new sheet1");
        wb.createSheet("second sheet2");
        String safeSheetName = WorkbookUtil.createSafeSheetName("O'Brien's sales*?]");
        wb.createSheet(safeSheetName);
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testCreateCells() throws Exception {
        //创建一个.xls格式的excel文件
        Workbook wb = new HSSFWorkbook();
        CreationHelper creationHelper = wb.getCreationHelper();
        //创建一个名称为new sheet的sheet页
        Sheet sheet1 = wb.createSheet("new sheet");
        Row row = sheet1.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(1);

        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(creationHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testCreateDateCells() throws Exception {
        Workbook wb = new HSSFWorkbook();
        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");

        //这种方式不行
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(new Date());

        //通过设计cell样式，设置date
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("m/d/yy h:mm"));
        cell = row.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        //通过Calendar生成日期
        cell = row.createCell(2);
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testDifTypeOfCells() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        Row row = sheet.createRow(2);
        row.createCell(0).setCellValue(1.1);
        row.createCell(1).setCellValue(new Date());
        row.createCell(2).setCellValue(Calendar.getInstance());
        row.createCell(3).setCellValue("a string");
        row.createCell(4).setCellValue(true);
        row.createCell(5).setCellType(CellType.ERROR);

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testFileVSInputStreams() throws Exception {
        Workbook wb = WorkbookFactory.create(new File("MyExcel.xls"));

        Workbook wb2 = WorkbookFactory.create(new FileInputStream("MyExcel.xls"));


        NPOIFSFileSystem fs = new NPOIFSFileSystem(new File("file.xls"));
        HSSFWorkbook wb3 = new HSSFWorkbook(fs.getRoot(), true);
        fs.close();

    }

    @Test
    public void testAlignment() throws Exception {

        Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();

        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow((short) 2);
        row.setHeightInPoints(30);

        createCell(wb, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);
        createCell(wb, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM);
        createCell(wb, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
        createCell(wb, row, 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER);
        createCell(wb, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY);
        createCell(wb, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
        createCell(wb, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("xssf-align.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * Creates a cell and aligns it a certain way.
     *
     * @param wb     the workbook
     * @param row    the row to create the cell in
     * @param column the column number to create the cell in
     * @param halign the horizontal alignment for the cell.
     */
    private static void createCell(Workbook wb, Row row, int column, HorizontalAlignment halign, VerticalAlignment valign) {
        Cell cell = row.createCell(column);
        cell.setCellValue("Align It");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);
    }

    @Test
    public void testWorkingBorders() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Row row = sheet.createRow(1);
        Cell cell = row.createCell(1);
        cell.setCellValue(4);

        //设置单元格样式
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());
        cellStyle.setBorderTop(BorderStyle.MEDIUM_DASHED);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(cellStyle);

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testColors() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow((short) 1);

        // Aqua background
        CellStyle style = wb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        Cell cell = row.createCell((short) 1);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        // Orange "foreground", foreground being the fill foreground not the font color.
        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell = row.createCell((short) 2);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testMergeCells() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Row row = sheet.createRow((short) 1);
        Cell cell = row.createCell((short) 1);
        cell.setCellValue("This is a test of merging");

        sheet.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)
                1, //last row  (0-based)
                1, //first column (0-based)
                2  //last column  (0-based)
        ));

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testWorkingWithFont() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(1);

        // Create a new font and alter it.
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 24);
        font.setFontName("Courier New");
        font.setItalic(true);
        font.setStrikeout(true);

        // Fonts are set into a style so create a new one to use.
        CellStyle style = wb.createCellStyle();
        style.setFont(font);

        // Create a cell and put a value in it.
        Cell cell = row.createCell(1);
        cell.setCellValue("This is a test of fonts");
        cell.setCellStyle(style);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void testCustomColors() throws Exception {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet();
        HSSFRow row = sheet.createRow((short) 0);
        HSSFCell cell = row.createCell((short) 0);
        cell.setCellValue("Default Palette");

        //apply some colors from the standard palette,
        // as in the previous examples.
        //we'll use red text on a lime background

        HSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(HSSFColor.LIME.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.RED.index);
        style.setFont(font);

        cell.setCellStyle(style);

        //save with the default palette
        FileOutputStream out = new FileOutputStream("default_palette.xls");
        wb.write(out);
        out.close();

        //now, let's replace RED and LIME in the palette
        // with a more attractive combination
        // (lovingly borrowed from freebsd.org)

        cell.setCellValue("Modified Palette");

        //creating a custom palette for the workbook
        HSSFPalette palette = wb.getCustomPalette();

        //replacing the standard red with freebsd.org red
        palette.setColorAtIndex(HSSFColor.RED.index,
                (byte) 153,  //RGB red (0-255)
                (byte) 0,    //RGB green
                (byte) 0     //RGB blue
        );
        //replacing lime with freebsd.org gold
        palette.setColorAtIndex(HSSFColor.LIME.index, (byte) 255, (byte) 204, (byte) 102);

        //save with the modified palette
        // note that wherever we have previously used RED or LIME, the
        // new colors magically appear
        out = new FileOutputStream("modified_palette.xls");
        wb.write(out);
        out.close();
    }

    /**
     * 测试单元格换行
     *
     * @throws Exception
     */
    @Test
    public void testUsingNewlinesInCells() throws Exception {
        Workbook wb = new XSSFWorkbook();   //or new HSSFWorkbook();
        Sheet sheet = wb.createSheet();

        Row row = sheet.createRow(2);
        Cell cell = row.createCell(2);
        cell.setCellValue("Use \n with word wrap on to create a new line");

        //to enable newlines you need set a cell styles with wrap=true
        CellStyle cs = wb.createCellStyle();
        cs.setWrapText(true);
        cell.setCellStyle(cs);

        //increase row height to accomodate two lines of text
        row.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));

        //adjust column width to fit the content
        sheet.autoSizeColumn((short) 2);

        FileOutputStream fileOut = new FileOutputStream("ooxml-newlines.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * 数字之间的格式转换
     *
     * @throws Exception
     */
    @Test
    public void testDataFormats() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("format sheet");
        CellStyle style;
        DataFormat format = wb.createDataFormat();
        Row row;
        Cell cell;
        short rowNum = 0;
        short colNum = 0;

        row = sheet.createRow(rowNum++);
        cell = row.createCell(colNum);
        cell.setCellValue(11111.25);
        style = wb.createCellStyle();
        //四舍五入
        style.setDataFormat(format.getFormat("0.0"));
        cell.setCellStyle(style);

        row = sheet.createRow(rowNum++);
        cell = row.createCell(colNum);
        cell.setCellValue(11111111111.25);
        style = wb.createCellStyle();
        //货币格式
        style.setDataFormat(format.getFormat("#,##0.0000"));
        cell.setCellStyle(style);

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * 没看出有什么不同！！！
     *
     * @throws Exception
     */
    @Test
    public void testFitSheetToOnePage() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("format sheet");
        PrintSetup ps = sheet.getPrintSetup();

        sheet.setAutobreaks(true);

        ps.setFitHeight((short) 1);
        ps.setFitWidth((short) 1);


        // Create various cells and rows for spreadsheet.

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * 设置打印区域
     *
     * @throws Exception
     */
    @Test
    public void testSetPrintArea() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        //sets the print area for the first sheet
        wb.setPrintArea(0, "$A$1:$C$2");

        //Alternatively:
//        wb.setPrintArea(
//                0, //sheet index
//                0, //start column
//                1, //end column
//                0, //start row
//                0  //end row
//        );

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * 设置页码
     *
     * @throws Exception
     */
    @Test
    public void testSetPageNumbersOnFooter() throws Exception {
        Workbook wb = new HSSFWorkbook(); // or new XSSFWorkbook();
        Sheet sheet = wb.createSheet("format sheet");
        Footer footer = sheet.getFooter();

        footer.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());


        // Create various cells and rows for spreadsheet.

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * poi提供的相关实用函数
     *
     * @throws Exception
     */
    @Test
    public void testConvenienceFunctions() throws Exception {
        Workbook wb = new HSSFWorkbook();  // or new XSSFWorkbook()
        Sheet sheet1 = wb.createSheet("new sheet");

        // Create a merged region
        // 创建一个合并区域
        Row row = sheet1.createRow(1);
        Row row2 = sheet1.createRow(2);
        Cell cell = row.createCell(1);
        cell.setCellValue("This is a test of merging");
        CellRangeAddress region = CellRangeAddress.valueOf("B2:E5");
        sheet1.addMergedRegion(region);

        // Set the border and border colors.
        // 设置边框和边框颜色
        final BorderStyle borderMediumDashed = BorderStyle.MEDIUM_DASHED;
        RegionUtil.setBorderBottom(borderMediumDashed,
                region, sheet1);
        RegionUtil.setBorderTop(borderMediumDashed,
                region, sheet1);
        RegionUtil.setBorderLeft(borderMediumDashed,
                region, sheet1);
        RegionUtil.setBorderRight(borderMediumDashed,
                region, sheet1);
        RegionUtil.setBottomBorderColor(IndexedColors.AQUA.getIndex(), region, sheet1);
        RegionUtil.setTopBorderColor(IndexedColors.AQUA.getIndex(), region, sheet1);
        RegionUtil.setLeftBorderColor(IndexedColors.AQUA.getIndex(), region, sheet1);
        RegionUtil.setRightBorderColor(IndexedColors.AQUA.getIndex(), region, sheet1);

        // Shows some usages of HSSFCellUtil
        CellStyle style = wb.createCellStyle();
        style.setIndention((short) 4);
        CellUtil.createCell(row, 8, "This is the value of the cell", style);
        Cell cell2 = CellUtil.createCell(row2, 8, "This is the value of the cell");
        CellUtil.setAlignment(cell2, HorizontalAlignment.CENTER);

        // Write out the workbook
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * 创建冻结窗格（将某行或某列固定）和分裂窗格
     *
     * @throws Exception
     */
    @Test
    public void testSplitsAndFreezePanes() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");
        Sheet sheet3 = wb.createSheet("third sheet");
        Sheet sheet4 = wb.createSheet("fourth sheet");

        // Freeze just one row
        sheet1.createFreezePane(0, 1, 0, 1);
        // Freeze just one column
        sheet2.createFreezePane(1, 0, 1, 0);
        // Freeze the columns and rows (forget about scrolling position of the lower right quadrant).
        sheet3.createFreezePane(2, 2);
        // Create a split with the lower left side being the active quadrant
        sheet4.createSplitPane(2000, 2000, 0, 0, Sheet.PANE_LOWER_LEFT);

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * 暂时不知道用途！！！
     *
     * @throws Exception
     */
    @Test
    public void testRepeatingRowsAndColumns() throws Exception {
        Workbook wb = new HSSFWorkbook();           // or new XSSFWorkbook();
        Sheet sheet1 = wb.createSheet("Sheet1");
        Sheet sheet2 = wb.createSheet("Sheet2");

        // Set the rows to repeat from row 4 to 5 on the first sheet.
        sheet1.setRepeatingRows(CellRangeAddress.valueOf("4:5"));
        // Set the columns to repeat from column A to C on the second sheet
        sheet2.setRepeatingColumns(CellRangeAddress.valueOf("A:C"));

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * 设置表格的头部和尾部
     *
     * @throws Exception
     */
    @Test
    public void testHeadersAndFooters() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Header header = sheet.getHeader();
        header.setCenter("Center Header");
        header.setLeft("Left Header");
        header.setRight(HSSFHeader.font("Stencil-Normal", "Italic") +
                HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * 各种超链接
     *
     * @throws Exception
     */
    @Test
    public void test1() throws Exception {
        Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();

        //cell style for hyperlinks
        //by default hyperlinks are blue and underlined
        CellStyle hlink_style = wb.createCellStyle();
        Font hlink_font = wb.createFont();
        hlink_font.setUnderline(Font.U_SINGLE);
        hlink_font.setColor(IndexedColors.BLUE.getIndex());
        hlink_style.setFont(hlink_font);

        Cell cell;
        Sheet sheet = wb.createSheet("Hyperlinks");
        //URL
        cell = sheet.createRow(0).createCell((short) 0);
        cell.setCellValue("URL Link");

        Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
        link.setAddress("http://poi.apache.org/");
        cell.setHyperlink(link);
        cell.setCellStyle(hlink_style);

        //link to a file in the current directory
        cell = sheet.createRow(1).createCell((short) 0);
        cell.setCellValue("File Link");
        link = createHelper.createHyperlink(HyperlinkType.FILE);
        link.setAddress("link1.xls");
        cell.setHyperlink(link);
        cell.setCellStyle(hlink_style);

        //e-mail link
        cell = sheet.createRow(2).createCell((short) 0);
        cell.setCellValue("Email Link");
        link = createHelper.createHyperlink(HyperlinkType.FILE);
        //note, if subject contains white spaces, make sure they are url-encoded
        link.setAddress("mailto:poi@apache.org?subject=Hyperlinks");
        cell.setHyperlink(link);
        cell.setCellStyle(hlink_style);

        //link to a place in this workbook

        //create a target sheet and cell
        Sheet sheet2 = wb.createSheet("Target Sheet");
        sheet2.createRow(0).createCell((short) 0).setCellValue("Target Cell");

        cell = sheet.createRow(3).createCell((short) 0);
        cell.setCellValue("Worksheet Link");
        Hyperlink link2 = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        link2.setAddress("'Target Sheet'!A1");
        cell.setHyperlink(link2);
        cell.setCellStyle(hlink_style);

        FileOutputStream out = new FileOutputStream("hyperinks.xlsx");
        wb.write(out);
        out.close();
    }

    @Test
    public void test() throws Exception {

    }

}














