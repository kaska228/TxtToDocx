package org.example;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STOnOff1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.*;

public class TxtToDocxConverter {

    private static final Map<String, String> HIERARCHY_MAP = new HashMap<>() {{
        put("SDH_Alarm_References", "SDH");
        put("OTN_Alarm_References", "OTN");
        put("DWDM_Alarm_References", "DWDM");
        put("PDH_Alarm_References", "PDH");
        put("Agent_Alarm_References", "AGENT");
        put("NMS_Alarm_References", "NMS");
    }};

    private static final Map<String, String> SEVERITY_MAP = new HashMap<>() {{
        put("MAJOR", "Серьезная");
        put("MINOR", "Малая");
        put("CRITICAL", "Критическая");
        put("WARNING", "Предупреждение");
    }};

    private static final Map<String, String> EVENT_TYPE_MAP = new HashMap<>() {{
        put("Communication alarm", "Авария связи");
        put("Quality of service alarm", "Авария качества обслуживания");
        put("Processing error alarm", "Сигнал об ошибке обработки");
        put("Equipment alarm", "Сигнализация оборудования");
        put("Environmental alarm", "Авария окружающей среды");
        put("Integrity alarm", "Сигнал об ошибке целостности");
        put("Operation alarm", "Сигнал об ошибке операции");
        put("Physical resource alarm", "Сигнал об ошибке физического ресурса");
        put("Security alarm", "Сигнал об ошибке безопасности");
        put("Time domain alarm", "Сигнал временной области");
        put("Property change", "Изменение свойства");
        put("Object creation", "Создание объекта");
        put("Object delete", "Удаление объекта");
        put("Relationship change", "Изменение отношений");
        put("State change", "Изменение состояния");
        put("Route change", "Изменение маршрута");
        put("Protection switching", "Переключение защиты");
        put("Over limit", "Превышение лимита");
        put("File transfer status", "Статус передачи файла");
        put("Backup status", "Статус резервного копирования");
        put("Heart beat", "Событие сердцебиения");
        put("Network alarm", "Авария сети");
    }};

    public static void main(String[] args) {
        String sqlFilePath = "D:/Alarm/alarm_data.sql"; // Путь к файлу SQL
        String docxFilePath = "D:/Alarm/alarmD.docx"; // Путь к выходному DOCX файлу
        String templateFilePath = "D:/Alarm/nms1.docx"; // Путь к шаблону DOCX файлу

        try {
            List<Accident> accidents = parseSqlFile(sqlFilePath);
            if (accidents.isEmpty()) {
                System.out.println("Нет данных для записи в DOCX файл.");
                return;
            }
            writeDocxFile(accidents, docxFilePath, templateFilePath);
            System.out.println("SQL файл успешно преобразован в DOCX файл.");

            // Извлечение заголовков из созданного файла
            extractHeadings(docxFilePath);

        } catch (IOException e) {
            System.err.println("Произошла ошибка при преобразовании файла: " + e.getMessage());
        }
    }

    public static List<Accident> parseSqlFile(String sqlFilePath) throws IOException {
        List<Accident> accidents = new ArrayList<>();
        Map<String, String> eventCategoryMap = new HashMap<>();
        boolean readingAccidents = false;
        boolean readingCategories = false;

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(sqlFilePath)))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (line.startsWith("alarmdataentity")) {
                    readingAccidents = true;
                    readingCategories = false;
                    continue;
                }

                if (line.startsWith("alarmeventcategory")) {
                    readingAccidents = false;
                    readingCategories = true;
                    continue;
                }

                if (line.trim().isEmpty() || line.startsWith("\\.")) {
                    continue;
                }

                if (readingCategories) {
                    String[] parts = line.split("\\t");
                    if (parts.length >= 2) {
                        String id = parts[0];
                        String categoryName = parts[1];
                        eventCategoryMap.put(id, categoryName);
                    }
                } else if (readingAccidents) {
                    String[] parts = line.split("\\t");
                    if (parts.length < 10) {
                        continue;
                    }

                    String hierarchy = parts[9].replace("\\N", "-");
                    if (!HIERARCHY_MAP.containsKey(hierarchy)) {
                        System.out.println("Пропуск строки с неверным hierarchy: " + hierarchy);
                        continue;
                    }

                    String severity = SEVERITY_MAP.getOrDefault(parts[1].replace("\\N", "-"), parts[1].replace("\\N", "-"));
                    String categoryId = parts[7].replace("\\N", "-").replaceAll("[\\[\\]\"]", "");
                    String category = eventCategoryMap.getOrDefault(categoryId, parts[7].replace("\\N", "-"));
                    String translatedCategory = translateCategory(category);
                    String eventType = EVENT_TYPE_MAP.getOrDefault(parts[6].replace("\\N", "-"), parts[6].replace("\\N", "-"));
                    String description = parts[4].replace("\\N", "-").replace("\\n", "\n");
                    String operatorAction = parts[5].replace("\\N", "-").replace("\\n", "\n");
                    String nameRus = parts[2].replace("\\N", "-");

                    accidents.add(new Accident(hierarchy, severity, translatedCategory, eventType, description, operatorAction, nameRus));
                }
            }
        }

        return accidents;
    }


    private static String translateCategory(String category) {
        category = category.replace("[", "").replace("]", "").replace("\"", "");
        String[] categories = category.split(",");
        List<String> translatedCategories = new ArrayList<>();
        for (String cat : categories) {
            translatedCategories.add(EVENT_TYPE_MAP.getOrDefault(cat.trim(), cat.trim()));
        }
        return String.join(", ", translatedCategories);
    }

    public static void writeDocxFile(List<Accident> accidents, String docxFilePath, String templateFilePath) throws IOException {
        accidents.sort(Comparator.comparing(Accident::getHierarchy).thenComparing(Accident::getDescription));

        Map<String, List<Accident>> groupedAccidents = new LinkedHashMap<>();
        for (Accident accident : accidents) {
            String section = HIERARCHY_MAP.getOrDefault(accident.getHierarchy(), "Неизвестный раздел");
            groupedAccidents.computeIfAbsent(section, k -> new ArrayList<>()).add(accident);
        }

        XWPFDocument template;
        try (FileInputStream fis = new FileInputStream(templateFilePath)) {
            template = new XWPFDocument(fis);
        }

        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(docxFilePath)) {

            // Копируем стили из шаблона
            XWPFStyles newStyles = document.createStyles();
            newStyles.setStyles(template.getStyle());

            // Добавляем настраиваемые стили заголовков
            addCustomHeadingStyle(document, "Heading1", 1);
            addCustomHeadingStyle(document, "Heading2", 2);

            // Создаем содержание с авариями
            createTableOfContents(document);

            // Добавляем аварии
            addAccidents(document, groupedAccidents);

            // Добавляем номера страниц в верхний колонтитул
            addPageNumbers(document);

            document.write(out);
        } catch (XmlException e) {
            throw new RuntimeException(e);
        }
    }

    private static void createTableOfContents(XWPFDocument document) {
        // Заголовок "СОДЕРЖАНИЕ"
        XWPFParagraph tocTitleParagraph = document.createParagraph();
        tocTitleParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun tocTitleRun = tocTitleParagraph.createRun();
        tocTitleRun.setText("СОДЕРЖАНИЕ");
        tocTitleRun.setFontSize(16);
        tocTitleRun.setFontFamily("Times New Roman");
        tocTitleRun.setBold(false);

        // Создаем TOC
        CTSdtBlock sdtBlock = document.getDocument().getBody().addNewSdt();
        CTSdtPr sdtPr = sdtBlock.addNewSdtPr();
        CTString docPart = sdtPr.addNewDocPartObj().addNewDocPartGallery();
        docPart.setVal("Table of Contents");
        sdtPr.addNewDocPartObj().addNewDocPartUnique().setVal(STOnOff1.ON);

        CTSdtContentBlock sdtContentBlock = sdtBlock.addNewSdtContent();
        XWPFParagraph paragraph = document.createParagraph();
        CTP ctp = paragraph.getCTP();
        sdtContentBlock.set(ctp);
        CTSimpleField tocField = ctp.addNewFldSimple();
        tocField.setInstr("TOC \\o \"1-3\" \\h \\z \\u");

        // Создаем параграф для содержания
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Times New Roman");
        run.setFontSize(12); // Обычный текст, 12pt
        run.setBold(false); // Убираем жирный шрифт для содержания

        // Вставка пустого пробела для обновления TOC
        run.setText(" ");

        // Настраиваем стили TOC
        addTOCStyles(document);
    }

    private static void addTOCStyles(XWPFDocument document) {
        addCustomHeadingStyle(document, "TOCHeading", 1);
        addCustomHeadingStyle(document, "TOC1", 2);
        addCustomHeadingStyle(document, "TOC2", 3);
        addCustomHeadingStyle(document, "TOC3", 4);
    }






    private static void addAccidents(XWPFDocument document, Map<String, List<Accident>> groupedAccidents) {
        int sectionNumber = 1;
        for (Map.Entry<String, List<Accident>> entry : groupedAccidents.entrySet()) {
            String section = entry.getKey();
            List<Accident> sectionAccidents = entry.getValue();

            // Добавляем заголовок раздела с номером
            XWPFParagraph sectionParagraph = document.createParagraph();
            sectionParagraph.setStyle("Heading1");
            sectionParagraph.setPageBreak(true); // Начать с новой страницы
            sectionParagraph.setAlignment(ParagraphAlignment.CENTER); // Выравнивание по центру
            XWPFRun sectionRun = sectionParagraph.createRun();
            sectionRun.setBold(true);
            sectionRun.setFontSize(16);
            sectionRun.setText("АВАРИИ: " + section.toUpperCase());
            sectionRun.setFontFamily("Times New Roman");
            sectionRun.addBreak();

            int accidentNumber = 1;
            for (Accident accident : sectionAccidents) {
                // Добавляем заголовок аварии с номером
                XWPFParagraph nameParagraph = document.createParagraph();
                nameParagraph.setStyle("Heading2");
                nameParagraph.setFirstLineIndent(600); // Отступ 1,5 см

                CTPPr ppr = nameParagraph.getCTP().getPPr();
                if (ppr == null) ppr = nameParagraph.getCTP().addNewPPr();

                CTOnOff keepNext = ppr.isSetKeepNext() ? ppr.getKeepNext() : ppr.addNewKeepNext();
                keepNext.setVal(STOnOff1.ON);

                CTOnOff keepLines = ppr.isSetKeepLines() ? ppr.getKeepLines() : ppr.addNewKeepLines();
                keepLines.setVal(STOnOff1.ON);

                XWPFRun nameRun = nameParagraph.createRun();
                nameRun.setBold(true);
                nameRun.setFontSize(14);
                nameRun.setFontFamily("Times New Roman");
                nameRun.setText(sectionNumber + "." + accidentNumber + " " + accident.getNameRus());
                nameRun.addBreak();

                // Создаем таблицу для деталей аварии
                XWPFTable table = document.createTable();
                table.setWidth("100%");
                setNoBorder(table);

                addTableRow(table, "Серьезность аварии:", accident.getSeverity());
                addTableRow(table, "Категория события:", accident.getCategory());
                addTableRow(table, "Тип события:", accident.getEventType());

                // Описание аварии
                XWPFParagraph descriptionParagraph = document.createParagraph();
                descriptionParagraph.setFirstLineIndent(600); // Отступ 1,5 см
                createFormattedParagraph(descriptionParagraph, "Описание аварии: ", accident.getDescription(), true);

                // Действия оператора
                XWPFParagraph operatorActionParagraph = document.createParagraph();
                operatorActionParagraph.setFirstLineIndent(600); // Отступ 1,5 см
                createFormattedParagraph(operatorActionParagraph, "Действия оператора: ", accident.getOperatorAction(), true);

                accidentNumber++;
            }
            sectionNumber++;
        }
    }









    private static void addTableRow(XWPFTable table, String label, String text) {
        XWPFTableRow row = table.createRow();
        XWPFTableCell cell0 = row.getCell(0);
        XWPFTableCell cell1 = row.getCell(1);

        // Ensure cells are not null and remove default empty paragraph
        if (cell0 == null) {
            cell0 = row.addNewTableCell();
        } else {
            cell0.removeParagraph(0); // Remove the default empty paragraph
        }
        if (cell1 == null) {
            cell1 = row.addNewTableCell();
        } else {
            cell1.removeParagraph(0); // Remove the default empty paragraph
        }

        setCellText(cell0, label, true);
        setCellText(cell1, text, false);

        setNoBorder(cell0);
        setNoBorder(cell1);
    }

    private static void setCellText(XWPFTableCell cell, String text, boolean bold) {
        // Remove existing paragraphs to avoid empty paragraphs
        while (cell.getParagraphs().size() > 0) {
            cell.removeParagraph(0);
        }

        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontFamily("Times New Roman");
        run.setFontSize(12); // Устанавливаем шрифт 12pt
        run.setBold(bold);
    }



    private static void setNoBorder(XWPFTableCell cell) {
        CTTcPr tcPr = cell.getCTTc().addNewTcPr();
        CTTcBorders borders = tcPr.addNewTcBorders();
        borders.addNewTop().setVal(STBorder.NONE);
        borders.addNewBottom().setVal(STBorder.NONE);
        borders.addNewLeft().setVal(STBorder.NONE);
        borders.addNewRight().setVal(STBorder.NONE);
    }

    private static void setNoBorder(XWPFTable table) {
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr == null) tblPr = table.getCTTbl().addNewTblPr();
        CTTblBorders borders = tblPr.addNewTblBorders();
        borders.addNewTop().setVal(STBorder.NONE);
        borders.addNewBottom().setVal(STBorder.NONE);
        borders.addNewLeft().setVal(STBorder.NONE);
        borders.addNewRight().setVal(STBorder.NONE);
        borders.addNewInsideH().setVal(STBorder.NONE);
        borders.addNewInsideV().setVal(STBorder.NONE);
    }

    private static void createFormattedParagraph(XWPFParagraph paragraph, String label, String text, boolean boldLabel) {
        XWPFRun run = paragraph.createRun();
        run.setBold(boldLabel);
        run.setText(label);
        run.setFontFamily("Times New Roman");
        run.setFontSize(12);
        run.addBreak();

        String[] lines = text.split("(?<=\\.)\\s*|\\s+-\\s+");

        for (String line : lines) {
            if (line.startsWith("-")) {
                // Если линия начинается с дефиса, форматируем как элемент списка
                XWPFRun listRun = paragraph.createRun();
                listRun.setText(line.trim());
                listRun.setFontFamily("Times New Roman");
                listRun.setFontSize(12);
                listRun.addBreak();
            } else {
                // Иначе просто добавляем текст
                XWPFRun textRun = paragraph.createRun();
                textRun.setText(line.trim());
                textRun.setFontFamily("Times New Roman");
                textRun.setFontSize(12);
                textRun.addBreak();
            }
        }
    }






    private static void addPageNumbers(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);

        // Создаем верхний колонтитул
        XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph paragraph = header.getParagraphArray(0);
        if (paragraph == null) {
            paragraph = header.createParagraph();
        }

        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Times New Roman");
        run.setFontSize(12);
        run.setText("");
        run.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        run = paragraph.createRun();
        run.setFontFamily("Times New Roman");
        run.setFontSize(12);
        run.getCTR().addNewInstrText().setStringValue("PAGE \\* MERGEFORMAT");
        run = paragraph.createRun();
        run.setFontFamily("Times New Roman");
        run.setFontSize(12);
        run.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        run = paragraph.createRun();
        run.setFontFamily("Times New Roman");
        run.setFontSize(12);
        run.getCTR().addNewT().setStringValue("1");
        run = paragraph.createRun();
        run.setFontFamily("Times New Roman");
        run.setFontSize(12);
        run.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);

        // Добавляем текст "7.ТАИЦ.00018-01 34 02" в центр верхнего колонтитула
        XWPFParagraph footerParagraph = header.createParagraph();
        footerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun footerRun = footerParagraph.createRun();
        footerRun.setFontFamily("Times New Roman");
        footerRun.setFontSize(12);
        footerRun.setText("7.ТАИЦ.00018-01 34 02");
    }

    private static void extractHeadings(String docxFilePath) {
        try (FileInputStream fis = new FileInputStream(docxFilePath);
             XWPFDocument document = new XWPFDocument(OPCPackage.open(fis))) {

            List<XWPFParagraph> paragraphs = document.getParagraphs();
            List<String> headings = new ArrayList<>();

            for (XWPFParagraph paragraph : paragraphs) {
                if (paragraph.getStyleID() != null) {
                    headings.add(paragraph.getText());
                }
            }

            // Output the headings
            for (String heading : headings) {
                System.out.println("Heading: " + heading);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPrGeneral ppr = CTPPrGeneral.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);

    }

    public static class Accident {
        private final String hierarchy;
        private final String severity;
        private final String category;
        private final String eventType;
        private final String description;
        private final String operatorAction;
        private final String nameRus;

        public Accident(String hierarchy, String severity, String category, String eventType, String description, String operatorAction, String nameRus) {
            this.hierarchy = hierarchy;
            this.severity = severity;
            this.category = category;
            this.eventType = eventType;
            this.description = description;
            this.operatorAction = operatorAction;
            this.nameRus = nameRus;
        }

        public String getHierarchy() {
            return hierarchy;
        }

        public String getSeverity() {
            return severity;
        }

        public String getCategory() {
            return category;
        }

        public String getEventType() {
            return eventType;
        }

        public String getDescription() {
            return description;
        }

        public String getOperatorAction() {
            return operatorAction;
        }

        public String getNameRus() {
            return nameRus;
        }
    }
}