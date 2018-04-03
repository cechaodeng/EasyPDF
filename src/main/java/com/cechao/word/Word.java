package com.cechao.word;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;

import java.io.*;

/**
 * @author Kent
 * @date 2018-03-28.
 */
public class Word {
    public static void main(String[] args) {

        File file = new File("E:\\workspace\\Maven\\EasyPDF\\src\\main\\resources\\test.doc");
        try (HWPFDocument document = new HWPFDocument(new BufferedInputStream(new FileInputStream(file)))) {
            Range range = document.getRange();

            for (int i = 0; i < range.numSections(); i++) {
                Section section = range.getSection(i);
                System.out.println("Paragraph 的数量：" + section.numParagraphs());
                for (int j = 0; j < section.numParagraphs(); j++) {
                    Paragraph paragraph = section.getParagraph(j);
                    for (int k = 0; k < paragraph.numCharacterRuns(); k++) {
                        CharacterRun characterRun = paragraph.getCharacterRun(k);
                        System.out.println(characterRun.text());
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     *
     * @Title: getTextFromWord
     * @Description: 读取word
     * @param filePath
     *            文件路径
     * @return: String 读出的Word的内容
     */
    public static String getTextFromWord(String filePath) {
        String result = null;
        File file = new File(filePath);
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(file);
            @SuppressWarnings("resource")
            WordExtractor wordExtractor = new WordExtractor(fis);
            result = wordExtractor.getText();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return result;
    }

    /**
     * @description 将数据归档到.doc的word文档中。数据续写到原目标文件末尾。
     * @param source
     *            源文件（必须存在！）
     * @param sourChs
     *            读取源文件要用的编码，若传入null，则默认是GBK编码
     * @param target
     *            目标word文档（必须存在！）
     */
    public static void storeDoc(File source, String sourChs, File target) {
        /*
         * 思路： 1.建立字符输入流，读取source中的数据。 2.在目标文件路径下new File：temp.doc
         * 3.将目标文件重命名为temp.doc，并用HWPFDocument类关联（temp.doc）。
         * 3.由temp.doc建立Range对象，写入source中的数据。 4.建立字节输出流，关联target。
         * 5.将range中的数据写入关联target的字节输出流。
         */
        if (!target.exists()) {
            throw new RuntimeException("目标文件不存在！");
        }
        if (sourChs == null) {
            sourChs = "GBK";
        }
        BufferedReader in = null;
        HWPFDocument temp = null;
        BufferedOutputStream out = null;
        String path = target.getParent();
        File tempDoc = new File(path, "temp.doc");
        target.renameTo(tempDoc);
        try {
            in = new BufferedReader(new InputStreamReader(new FileInputStream(source), sourChs));
            temp = new HWPFDocument(new BufferedInputStream(new FileInputStream(tempDoc)));
            out = new BufferedOutputStream(new FileOutputStream(target));
            Range range = temp.getRange();
            String line = null;
            //range.insertAfter(getDate(12));
            range.insertAfter("\r");
            while ((line = in.readLine()) != null) {
                range.insertAfter(line);
                // word中\r是换行符
                range.insertAfter("\r");
            }
            range.insertAfter("\r");
            range.insertAfter("\r");
            temp.write(out);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (temp != null) {
                try {
                    temp.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            tempDoc.deleteOnExit();
        }
    }
}
