package com.markchen;

import org.apache.commons.io.IOUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author markchen
 * @version 1.0
 * @date 2023/7/14 23:02
 */
public class Demo1 {
    public static void main(String[] args) throws IOException {
        InputStream inputStream = Demo1.class.getClassLoader().getResourceAsStream("template/template.docx");
        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
        List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
        Pattern compile = Pattern.compile("^\\{(.*)\\}$");
        for (XWPFParagraph paragraph : paragraphs) {
            for (XWPFRun run : paragraph.getRuns()) {
                Matcher matcher = compile.matcher(run.getText(0));
                if (matcher.matches()) {
                    System.out.println(matcher.group(1));
                    run.setText("100", 0);
                }
            }
        }

        Project project1 = new Project("Hi2312", "chen", "25", "66%", "1234.56", "1255.6", "Hi2312", "智慧视觉");
        Project project2 = new Project("Hi2312 V100", "chen", "25", "66%", "1234.56", "1255.6", "Hi2312", "智慧视觉");
        Project project3 = new Project("Hi2312 V101", "chen", "25", "66%", "1234.56", "1255.6", "Hi2312", "智慧视觉");
        Project project4 = new Project("Hi2312 V102", "chen", "30", "66%", "1234.56", "1255.6", "Hi2312", "智慧视觉");
        Project project5 = new Project("Hiblue", "chen", "25", "66%", "1234.56", "1255.6", "Hiblue", "智慧视觉");
        Project project6 = new Project("Hiblue V100", "chen", "25", "66%", "1234.56", "1255.6", "Hiblue", "智慧视觉");
        Project project7 = new Project("Hiblue V101", "chen", "25", "66%", "1234.56", "1255.6", "Hiblue", "智慧视觉");
        Project project8 = new Project("Hiblue V102", "chen", "26", "66%", "1234.56", "1255.6", "Hiblue", "智慧视觉");
        List<Project> data = Arrays.asList(project1, project2, project3, project4, project5, project6, project7,
                project8);
/*        List<Project> list1 = Arrays.asList(project1, project2, project3, project4);
        List<Project> list2 = Arrays.asList(project5, project6, project7, project8);
        Product hi2312 = new Product("Hi2312", list1);
        Product hiblue = new Product("Hiblue", list2);
        Domain domain = new Domain("智慧视觉", Arrays.asList(hi2312, hiblue));*/
        // 处理表格
        List<XWPFTable> tables = xwpfDocument.getTables();
        if (Objects.nonNull(tables) && !tables.isEmpty()) {
            XWPFTable xwpfTable = tables.get(0);
            for (int i = 0; i < data.size(); i++) {
                XWPFTableRow row = xwpfTable.createRow();
                row.getCell(0).setText(data.get(i).domainName);
                row.getCell(1).setText(data.get(i).productName);
                row.getCell(2).setText(data.get(i).projectName);
                row.getCell(3).setText(data.get(i).projectManager);
                XWPFTableCell cell = row.getCell(4);
                cell.setText(data.get(i).deviation);
                if (Integer.parseInt(data.get(i).deviation) > 25) {
                    cell.setColor("ff0000");
                }
                row.getCell(5).setText(project1.deviationRate);
                XWPFTableCell xwpfTableCell = row.addNewTableCell();
                xwpfTableCell.setText(project1.budget);
                XWPFTableCell xwpfTableCell1 = row.addNewTableCell();
                xwpfTableCell1.setText(project1.assign);
            }
            merge(xwpfTable, 2, 0, 2);
            XWPFTableCell cell = xwpfTable.getRow(2).getCell(0);
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            merge(xwpfTable,2,1,2);
        }
        xwpfDocument.write(new FileOutputStream(new File("aaa.docx")));
    }

    public static void merge(XWPFTable table, int rowIndex, int colIndex, int tableHeadRows) {
        String firstText = table.getRow(rowIndex).getCell(colIndex).getText();
        int rows = table.getNumberOfRows() - tableHeadRows;
        if (rows <= 1) {
            return;
        }
        int mergeStart = tableHeadRows;
        for (int i = tableHeadRows + 1; i < table.getNumberOfRows(); i++) {
            String currentText = table.getRow(i).getCell(colIndex).getText();
            if (!currentText.equals(firstText) && (i - mergeStart) > 1) {
                mergeCell(table, mergeStart, i - 1, colIndex);
                mergeStart = i;
                firstText = currentText;
            }
            if (i == (table.getNumberOfRows()-1) && (i - mergeStart) > 1) {
                mergeCell(table, mergeStart, i, colIndex);
            }
        }
    }

    public static void mergeCell(XWPFTable table, int beginRowIndex, int endRowIndex, int colIndex) {
        if (beginRowIndex == endRowIndex || beginRowIndex > endRowIndex) {
            return;
        }
        //合并行单元格的第一个单元格
        CTVMerge startMerge = CTVMerge.Factory.newInstance();
        startMerge.setVal(STMerge.RESTART);
        //合并行单元格的第一个单元格之后的单元格
        CTVMerge endMerge = CTVMerge.Factory.newInstance();
        endMerge.setVal(STMerge.CONTINUE);
        XWPFTableCell cell1 = table.getRow(beginRowIndex).getCell(colIndex);
        if (cell1.getCTTc().getTcPr() == null) {
            cell1.getCTTc().addNewTcPr();
        }
        cell1.getCTTc().getTcPr().setVMerge(startMerge);
        for (int i = beginRowIndex + 1; i <= endRowIndex; i++) {
            XWPFTableCell cell = table.getRow(i).getCell(colIndex);
            if (cell.getCTTc().getTcPr() == null) {
                cell.getCTTc().addNewTcPr();
            }
            cell.getCTTc().getTcPr().setVMerge(endMerge);
        }
    }

    static class Domain {
        public String domainName;
        public List<Product> productList;

        public Domain(String domainName, List<Product> productList) {
            this.domainName = domainName;
            this.productList = productList;
        }
    }

    static class Product {
        public String productName;
        public List<Project> projectList;

        public Product(String productName, List<Project> projectList) {
            this.productName = productName;
            this.projectList = projectList;
        }
    }

    static class Project {
        public String projectName;
        public String projectManager;
        public String deviation;
        public String deviationRate;
        public String budget;
        public String assign;
        public String productName;
        public String domainName;

        public Project(String projectName, String projectManager, String deviation, String deviationRate,
                       String budget, String assign,
                       String productName, String domainName) {
            this.projectName = projectName;
            this.projectManager = projectManager;
            this.deviation = deviation;
            this.deviationRate = deviationRate;
            this.budget = budget;
            this.assign = assign;
            this.productName = productName;
            this.domainName = domainName;
        }
    }
}
