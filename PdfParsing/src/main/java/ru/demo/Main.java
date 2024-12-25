package ru.demo;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.cos.COSName;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.pdfbox.text.PDFTextStripper;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {

    public static void main(String[] args) {
        String pdfFilePath = "file.pdf";
        String docxFilePath = "Document.docx";
        try  (PDDocument document = PDDocument.load(new File(pdfFilePath));
             XWPFDocument wordDocument = new XWPFDocument()) {
            PDPageTree pages = document.getPages();
            int imageCounter = 1;

            for (PDPage page : pages) {
                PDFTextStripper pdfStripper = new PDFTextStripper();
                String pageText = pdfStripper.getText(document);
                XWPFParagraph textParagraph = wordDocument.createParagraph();
                XWPFRun textRun = textParagraph.createRun();
                textRun.setText(pageText);
                PDResources resources = page.getResources();
                for (COSName name : resources.getXObjectNames()) {
                    if (resources.getXObject(name) instanceof PDImageXObject) {
                        PDImageXObject image = (PDImageXObject) resources.getXObject(name);
                        BufferedImage bufferedImage = image.getImage();
                        File tempImageFile = new File("temp_image_" + imageCounter + ".png");
                        ImageIO.write(bufferedImage, "PNG", tempImageFile);


                        try (FileInputStream imageStream = new FileInputStream(tempImageFile)) {
                            XWPFParagraph imageParagraph = wordDocument.createParagraph();
                            XWPFRun imageRun = imageParagraph.createRun();
                            imageRun.addPicture(imageStream, XWPFDocument.PICTURE_TYPE_PNG, tempImageFile.getName(),
                                    Units.toEMU(bufferedImage.getWidth()/2), Units.toEMU(bufferedImage.getHeight()/2));
                        } catch (Exception e) {
                            System.err.println("Ошибка при добавлении изображения: " + e.getMessage());
                        }
                        imageCounter++;
                    }
                }
            }

            try (FileOutputStream out = new FileOutputStream(docxFilePath)) {
                wordDocument.write(out);
            }
            System.out.println("Текст и изображения успешно извлечены из PDF и сохранены в Word.");
        } catch (IOException e) {
            System.err.println("Ошибка при обработке: " + e.getMessage());
        }
    }
}