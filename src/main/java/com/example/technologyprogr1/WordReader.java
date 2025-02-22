package com.example.technologyprogr1;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class WordReader {
    public static List<String> readWordFile(File file) throws IOException {
        List<String> content = new ArrayList<>();
        try (XWPFDocument document = new XWPFDocument(new FileInputStream(file))) {
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                content.add(paragraph.getText());
            }
        }
        return content;
    }
}
