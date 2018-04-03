package com.cechao.convert;

import com.cechao.pdf.PDF;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;

/**
 * @author Kent
 * @date 2018-03-28.
 */
public class PdfToWord {
    public static void main(String[] args) {
        /**
         * 1.读取pdf
         * 2.写入word
         */
        String pdfPath = "E:\\workspace\\Maven\\EasyPDF\\src\\main\\resources\\resource.pdf";
        //XWPFDocument wordDoc = new XWPFDocument();
        try(FileOutputStream out = new FileOutputStream(new File("E:\\workspace\\Maven\\EasyPDF\\src\\main\\resources\\test.docx"))){
            String pdfContent = PDF.getTextFromPdf(pdfPath);
            //XWPFRun run = wordDoc.createParagraph().createRun();
            //run.setText(pdfContent);
            //wordDoc.write(out);

            XWPFDocument document = new XWPFDocument();
            XWPFParagraph firstParagraph = document.createParagraph();
            XWPFRun run = firstParagraph.createRun();
            run.setText(pdfContent);
            run.setFontSize(16);

            //换行
            XWPFParagraph paragraph1 = document.createParagraph();
            XWPFRun paragraphRun1 = paragraph1.createRun();
            paragraphRun1.setText("\r");

            document.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {

        }


    }
}
