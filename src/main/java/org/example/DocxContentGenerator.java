package org.example;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;

public class DocxContentGenerator {

    public static void main(String[] args) {
        try {
            // Открытие существующего документа
            XWPFDocument doc = new XWPFDocument(new FileInputStream("D:/Alarm/alarmD.docx"));

            // Добавление содержания
            addTableOfContents(doc);

            // Сохранение документа с новыми изменениями
            try (FileOutputStream out = new FileOutputStream("D:/Alarm/alarmD_updated.docx")) {
                doc.write(out);
            }

            System.out.println("Документ успешно обновлен.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void addTableOfContents(XWPFDocument doc) {
        XWPFParagraph tocParagraph = doc.createParagraph();
        tocParagraph.setStyle("TOCHeading");

        XWPFRun tocRun = tocParagraph.createRun();
        tocRun.setText("СОДЕРЖАНИЕ");

        // Добавление содержания
        XWPFParagraph p = doc.createParagraph();
        CTP ctP = p.getCTP();
        CTSimpleField toc = ctP.addNewFldSimple();
        toc.setInstr("TOC \\o \"1-3\" \\h \\z \\u");

        XWPFRun r1 = p.createRun();
        r1.setText("Cодержание");
        r1.setBold(true);
        r1.setFontFamily("Times New Roman");
        r1.setFontSize(16);
    }

    private static void addPageNumbers(XWPFDocument doc) {
        XWPFHeaderFooterPolicy policy = doc.createHeaderFooterPolicy();

        CTP ctpFooter = CTP.Factory.newInstance();
        CTR ctrFooter = ctpFooter.addNewR();
        CTText ctTextFooter = ctrFooter.addNewT();
        ctTextFooter.setStringValue("Page ");
        ctrFooter.addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        ctrFooter.addNewInstrText().setStringValue("PAGE \\* MERGEFORMAT");
        ctrFooter.addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        ctrFooter.addNewT().setStringValue("1");
        ctrFooter.addNewFldChar().setFldCharType(STFldCharType.END);
        CTP ctpFooterPageCount = CTP.Factory.newInstance();
        CTR ctrFooterPageCount = ctpFooterPageCount.addNewR();
        ctrFooterPageCount.addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        ctrFooterPageCount.addNewInstrText().setStringValue("NUMPAGES \\* MERGEFORMAT");
        ctrFooterPageCount.addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        ctrFooterPageCount.addNewT().setStringValue("1");
        ctrFooterPageCount.addNewFldChar().setFldCharType(STFldCharType.END);

        XWPFParagraph[] parsFooter = new XWPFParagraph[2];
        parsFooter[0] = new XWPFParagraph(ctpFooter, doc);
        parsFooter[1] = new XWPFParagraph(ctpFooterPageCount, doc);
        parsFooter[0].setAlignment(ParagraphAlignment.RIGHT);
        parsFooter[1].setAlignment(ParagraphAlignment.RIGHT);

        policy.createFooter(STHdrFtr.DEFAULT, parsFooter);
    }
}
