package com.xbinfo.zipreader.util;

import java.io.*;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class ZipReader {

    public static void main(String[] args) {
        String zipFilePath = "/Users/zhuweiliu/Downloads/计算机基础知识.zip";
        readZipFile(zipFilePath);
    }

    public static void readZipFile(String zipFilePath) {
        File file = new File(zipFilePath);

        if (!file.exists() || !file.isFile()) {
            System.out.println("文件不存在或不是有效的压缩包: " + zipFilePath);
            return;
        }

        try (ZipFile zipFile = new ZipFile(file)) {
            Enumeration<? extends ZipEntry> entries = zipFile.entries();

            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                System.out.println("条目名称: " + entry.getName());

                if (entry.isDirectory()) {
                    System.out.println("这是一个目录");
                    continue;
                }

                String outputFileName = "output_" + entry.getName().replace("/", "_") + ".txt";

                // 判断文件类型
                if (entry.getName().endsWith(".pdf")) {
                    System.out.println("PDF文件内容:");
                    handlePdfFile(zipFile, entry, outputFileName);
                } else if (entry.getName().endsWith(".doc") || entry.getName().endsWith(".docx")) {
                    System.out.println("处理 Word 文件:");
                    handleWordFile(zipFile, entry, outputFileName);
                } else {
                    System.out.println("其他文件类型（如文本）内容:");
                    printFileContent(zipFile, entry, outputFileName);
                }
            }
        } catch (IOException e) {
            System.out.println("读取压缩包出错: " + e.getMessage());
        }
    }

    // 处理PDF文件并输出到文本
    public static void handlePdfFile(ZipFile zipFile, ZipEntry entry, String outputFileName) {
        try (InputStream is = zipFile.getInputStream(entry);
             PDDocument document = PDDocument.load(is)) {
            PDFTextStripper pdfStripper = new PDFTextStripper();
            String pdfText = pdfStripper.getText(document);

            // 输出PDF内容到文本文件
            writeToFile(pdfText, outputFileName);
        } catch (IOException e) {
            System.out.println("读取PDF文件内容出错: " + e.getMessage());
        }
    }

    // 处理 Word 文件并输出到文本
    public static void handleWordFile(ZipFile zipFile, ZipEntry entry, String outputFileName) {
        try (InputStream is = zipFile.getInputStream(entry)) {
            if (isDocxFile(is)) {
                System.out.println("处理 DOCX 文件");
                try (XWPFDocument docx = new XWPFDocument(is)) {
                    XWPFWordExtractor extractor = new XWPFWordExtractor(docx);
                    String text = extractor.getText();
                    writeToFile(text, outputFileName);
                }
            } else if (isDocFile(is)) {
                System.out.println("处理 DOC 文件");
                try (HWPFDocument doc = new HWPFDocument(is)) {
                    WordExtractor extractor = new WordExtractor(doc);
                    String text = extractor.getText();
                    writeToFile(text, outputFileName);
                }
            } else {
                System.out.println("该文件不是有效的 DOC 或 DOCX 文件");
            }
        } catch (IOException e) {
            System.out.println("读取 Word 文件内容出错: " + e.getMessage());
        }
    }

    // 打印压缩包中其他文件的内容并输出到文本
    public static void printFileContent(ZipFile zipFile, ZipEntry entry, String outputFileName) {
        try (InputStream is = zipFile.getInputStream(entry);
             BufferedReader reader = new BufferedReader(new InputStreamReader(is))) {

            StringBuilder content = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                content.append(line).append("\n");
            }

            writeToFile(content.toString(), outputFileName);
        } catch (IOException e) {
            System.out.println("读取文件内容出错: " + e.getMessage());
        }
    }

    // 将内容写入文本文件
    public static void writeToFile(String content, String outputFileName) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputFileName))) {
            writer.write(content);
            System.out.println("文件已成功输出到: " + outputFileName);
        } catch (IOException e) {
            System.out.println("写入文件出错: " + e.getMessage());
        }
    }

    // 检查是否为 DOCX 文件（通过文件头）
    public static boolean isDocxFile(InputStream is) {
        try {
            byte[] signature = new byte[4];
            is.read(signature, 0, 4);
            is.reset();  // 重置流到起始位置
            return signature[0] == 'P' && signature[1] == 'K' && signature[2] == 3 && signature[3] == 4; // DOCX 是一个 ZIP 文件格式，PK 为 ZIP 文件头
        } catch (IOException e) {
            return false;
        }
    }

    // 检查是否为 DOC 文件（通过文件头）
    public static boolean isDocFile(InputStream is) {
        try {
            byte[] signature = new byte[8];
            is.read(signature, 0, 8);
            is.reset();  // 重置流到起始位置
            return signature[0] == (byte) 0xD0 && signature[1] == (byte) 0xCF && signature[2] == (byte) 0x11 && signature[3] == (byte) 0xE0; // DOC 是复合文件格式（Compound File）
        } catch (IOException e) {
            return false;
        }
    }
}
