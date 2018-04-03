package com.cechao.pdf;

import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * @author Kent
 * @date 2018-03-28.
 */
public class PDF {

    public static void main(String[] args) {
        //File file = new File("E:\\workspace\\Maven\\EasyPDF\\src\\main\\resources\\acctoutBuy2017032416514698.pdf");
        getTextFromPdf("E:\\workspace\\Maven\\EasyPDF\\src\\main\\resources\\acctoutBuy2017032416514698.pdf");
    }
    /**
     *
     * @Title: getTextFromPdf
     * @Description: 读取pdf文件内容
     * @param filePath
     * @return: 读出的pdf的内容
     */
    public static String getTextFromPdf(String filePath) {
        String result = null;
        FileInputStream is = null;
        PDDocument document = null;
        try {
            is = new FileInputStream(filePath);
            PDFParser parser = new PDFParser(is);
            parser.parse();
            document = parser.getPDDocument();
            PDFTextStripper stripper = new PDFTextStripper();
            result = stripper.getText(document);
            //System.out.println(result);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (document != null) {
                try {
                    document.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return result;
    }
}
