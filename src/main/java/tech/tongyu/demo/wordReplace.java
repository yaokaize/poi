package tech.tongyu.demo;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;

public class wordReplace {

    public static void main(String[] args) throws Exception {
        String filePath =  "/Users/yaokaize/poi/src/main/resources/交易确认书.docx";
        String fileDemo =  "/Users/yaokaize/poi/src/main/resources/demo.docx";

        if ("doc".equals(filePath.split("\\.")[1])) {
        } else if ("docx".equals(filePath.split("\\.")[1])) {

            XWPFDocument xwpfDocument = new XWPFDocument(POIXMLDocument.openPackage(filePath));
            // 编辑文本
            Iterator<XWPFParagraph> paragraphsIterator = xwpfDocument.getParagraphsIterator();
            while (paragraphsIterator.hasNext()) {
                XWPFParagraph xwpfParagraph = paragraphsIterator.next();
                List<XWPFRun> runs = xwpfParagraph.getRuns();
                for (XWPFRun run : runs) {
                    String oneParaString = run.getText(run.getTextPosition());
                    if (StringUtils.isBlank(oneParaString)) {
                        continue;
                    }
                    oneParaString = oneParaString.replace("${name}", "中证");
                    oneParaString = oneParaString.replace("${partyName}", "中证");
                    oneParaString = oneParaString.replace("${party}", "中证");
                    run.setText(oneParaString, 0);
                }
            }
            // 编辑表格
            Iterator<XWPFTable> tablesIterator = xwpfDocument.getTablesIterator();
            while (tablesIterator.hasNext()) {
                XWPFTable table = tablesIterator.next();

            }


            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            FileOutputStream out = new FileOutputStream(fileDemo);
            xwpfDocument.write(out);
            out.write(outputStream.toByteArray());
            out.flush();

        } else {
            throw new RuntimeException("不支持此文件上传格式");
        }

    }

}
