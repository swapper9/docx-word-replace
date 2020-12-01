package com.swap;

import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DocxWordsReplace {
    public static void main(String[] args) throws IOException {
        Map<String, String> replaceMap = new HashMap<>();
        replaceMap.put("${user}", "Иванов Иван Иванович");
        replaceMap.put("${phone}", "(800)880-8811");
        XWPFDocument doc = new XWPFDocument(new FileInputStream("/input.docx"));
        XWPFDocument outFile = replaceAllWords(doc, replaceMap);
        outFile.write(new FileOutputStream("/output.docx"));
        outFile.close();
    }

    private static XWPFDocument replaceAllWords(XWPFDocument doc, Map<String, String> replaceMap){
        for (XWPFHeader header : doc.getHeaderList()) {
            replaceAllBodyElements(header.getBodyElements(), replaceMap);
        }
        replaceAllBodyElements(doc.getBodyElements(), replaceMap);
        return doc;
    }

    private static void replaceAllBodyElements(List<IBodyElement> bodyElements, Map<String, String> replaceMap){
        for (IBodyElement bodyElement : bodyElements) {
            if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0) {
                replaceInParagraph((XWPFParagraph) bodyElement, replaceMap);
            }
            if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0) {
                replaceInTable((XWPFTable) bodyElement, replaceMap);
            }
        }
    }

    private static void replaceInTable(XWPFTable table, Map<String, String> replaceMap) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (IBodyElement bodyElement : cell.getBodyElements()) {
                    if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0) {
                        replaceInParagraph((XWPFParagraph) bodyElement, replaceMap);
                    }
                    if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0) {
                        replaceInTable((XWPFTable) bodyElement, replaceMap);
                    }
                }
            }
        }
    }

    private static void replaceInParagraph(XWPFParagraph paragraph, Map<String, String> replaceMap) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (Map.Entry<String, String> replPair : replaceMap.entrySet()) {
            String placeholder = replPair.getKey();
            String replace = replPair.getValue();
            TextSegment segment = paragraph.searchText(placeholder, new PositionInParagraph());
            if ( segment != null ) {
                if ( segment.getBeginRun() == segment.getEndRun() ) {
                    XWPFRun run = runs.get(segment.getBeginRun());
                    String runText = run.getText(run.getTextPosition());
                    String replaced = runText.replace(placeholder, replace);
                    run.setText(replaced, 0);
                } else {
                    StringBuilder b = new StringBuilder();
                    for (int runPos = segment.getBeginRun(); runPos <= segment.getEndRun(); runPos++) {
                        XWPFRun run = runs.get(runPos);
                        b.append(run.getText(run.getTextPosition()));
                    }
                    String connectedRuns = b.toString();
                    String replaced = connectedRuns.replace(placeholder, replace);

                    XWPFRun partOne = runs.get(segment.getBeginRun());
                    partOne.setText(replaced, 0);
                    for (int runPos = segment.getBeginRun()+1; runPos <= segment.getEndRun(); runPos++) {
                        XWPFRun partNext = runs.get(runPos);
                        partNext.setText("", 0);
                    }
                }
            }
        }
    }

}
