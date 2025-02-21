package com.example.technologyprogr1;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

public class WordWriter {
    public static void writeToWord(List<String> data, File outputFile) {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(outputFile)) {
            for (String line : data) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText(line);
            }
            document.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}