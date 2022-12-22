package com.a3k.pocs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.MalformedURLException;
import java.net.URL;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.BreakType;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.TextRange;
import jakarta.xml.bind.JAXBElement;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.AltChunkType;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.wml.Text;

public class POC {
    public static void main(String[] args) throws Docx4JException {
        //  generateApachePoiWord();
        generateSpireDoc();
        //generatedocx4j();
    }

    private static void generatedocx4j() throws Docx4JException {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
        Style s = mdp.getStyleDefinitionsPart().getDefaultParagraphStyle();
        ObjectFactory factory = Context.getWmlObjectFactory();
        HpsMeasure hpsmeasure = Context.getWmlObjectFactory().createHpsMeasure();
        hpsmeasure.setVal(BigInteger.valueOf(26));
        RPr rpr = s.getRPr();
        if (rpr == null) {
            rpr = factory.createRPr();
            s.setRPr(rpr);
        }
        rpr.setSz(hpsmeasure);
        RFonts rf = rpr.getRFonts();
        if (rf == null) {
            rf = factory.createRFonts();
            rpr.setRFonts(rf);
        }
        // This is where you set your font name.
        rf.setAscii("Times New Roman");
        P p = new P();
        R r = factory.createR();
        p.getContent().add(r);

        // Create object for rPr
        RPr rpr1 = factory.createRPr();
        r.setRPr(rpr1);
        // Create object for b
        rpr1.setB(factory.createBooleanDefaultTrue());

        // Create object for t (wrapped in JAXBElement)
        Text text = factory.createText();
        JAXBElement<Text> textWrapped = factory.createRT(text);
        r.getContent().add( textWrapped);
        text.setValue( "Title:");

        // Create object for second run
        R r2 = factory.createR();
        p.getContent().add( r2);

        // Create object for rPr
        RPr rpr2 = factory.createRPr();
        r2.setRPr(rpr2);

        // Create object for t (wrapped in JAXBElement)
        Text text2 = factory.createText();
        JAXBElement<org.docx4j.wml.Text> textWrapped2 = factory.createRT(text2);
        r2.getContent().add( textWrapped2);
        text2.setValue( " Diana Trujillo: A Lesson in Perseverance");
        text2.setSpace( "preserve");
        mdp.addObject(p);

        try (InputStream input = new URL("https://cdn.achieve3000.com/assets/content/images/KB/w/home/WorldCupWrap_Hero.jpg").openStream()) {
            String altText = null;
            String filenameHint = null;
            int id2 = 1;

            P p1 = newImage(wordMLPackage, input.readAllBytes(), filenameHint, altText, id2);
            wordMLPackage.getMainDocumentPart().addObject(p1);

        } catch (MalformedURLException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }


        // Add the XHTML altChunk
        String xhtml = "<html>" +
                "<p>We normally use an <b>apostrophe</b> followed by <b>the letter <i>s</i></b> to show that something belongs to someone or something. So if we said <i>The book is Rebecca's</i>, it would mean that the book belongs to Rebecca. The new word <i>(Rebecca's)</i> is called a <u><a>POSSESSIVE</a></u>.  " +
                " <p><b>What's the Rule?</b>" +
                " <ul>" +
                " <li>To form a possessive, add an apostrophe and an s to the end of the word." +
                " <p>Unfortunately, there are a lot of exceptions to this rule." +
                " <li>When the word is PLURAL and ENDS WITH THE LETTER <i>S</i>, we do not put another <i>s</i> after the apostrophe.  " +
                " <p><b>Let's look at an example:</b><br>" +
                " <i>This is the students' cafeteria.</i> <br>" +
                " This means the same thing as <i>This cafeteria belongs to the student<u>s</u>.</i>" +
                " <p><li>When the word is SINGULAR and ENDS WITH THE LETTER <i>S</i>, we have a choice: We can either put another s after the apostrophe, or we can leave it out." +
                " <p><b>Let's look at an example:</b><br>" +
                " <i>One of the bus's wheels is loose.<br>" +
                " One of the bus' wheels is loose.</i>" +
                " <p><li>When using a POSSESSIVE PRONOUN (ex. yours, mine, his, hers, theirs, its, ours, whose) we do NOT use an apostrophe." +
                " <p><b>Let's look at an example:</b><br>" +
                " The cat licked its entire body.<br>" +
                " Here, the body does belong to <i>it</i>. However, the word <i>its</i> is a possessive pronoun, so there is NO APOSTROPHE. The word <i>it's</i> is a contraction [link to contractions] that means <i>it is</i>." +
                " </ul>" +
                " <p><b>Your turn</b><br>" +
                " Let's see how well you know apostrophes. Click on the activity button above. " +
                " </p><br><br>" +
                " <p align=center><b><font color=\"#990033\"><span name=dic>" +
                "</html>";
        mdp.addAltChunk(AltChunkType.Xhtml, xhtml.getBytes());

        mdp.convertAltChunks();

        mdp.addParagraphOfText("Paragraph 3");

        wordMLPackage.save(new File("docx4j.docx"));

    }


    public static org.docx4j.wml.P newImage(WordprocessingMLPackage wordMLPackage,
                                            byte[] bytes,
                                            String filenameHint, String altText,
                                            int id2) throws Exception {

        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        Inline inline = imagePart.createImageInline(filenameHint, altText,
                wordMLPackage.getDrawingPropsIdTracker().generateId(), id2, false);

        // Now add the inline in w:p/w:r/w:drawing
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();
        org.docx4j.wml.R run = factory.createR();
        p.getContent().add(run);
        org.docx4j.wml.Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);

        return p;

    }

    private static void generateApachePoiWord() {
        //Blank Document
        XWPFDocument document = new XWPFDocument();

        //Write the Document in file system
        try (FileOutputStream out = new FileOutputStream("poi.docx");
             InputStream input = new URL("https://cdn.achieve3000.com/assets/content/images/KB/w/home/WorldCupWrap_Hero.jpg").openStream()) {
            XWPFParagraph titleParagraph = document.createParagraph();
            XWPFRun runTitleHeader = titleParagraph.createRun();
            runTitleHeader.setText("Title: ");
            runTitleHeader.setFontFamily("Times New Roman");
            runTitleHeader.setFontSize(12);
            runTitleHeader.setBold(true);
            XWPFRun runTitleText = titleParagraph.createRun();
            runTitleText.setText("Diana Trujillo: A Lesson in Perseverance");
            runTitleText.setFontFamily("Times New Roman");
            runTitleText.setFontSize(12);
            runTitleText.addBreak();
            runTitleText.addBreak();
            XWPFParagraph lidParagraph = document.createParagraph();
            XWPFRun lid = lidParagraph.createRun();
            lid.setText("LID: 19978");
            lid.setText("<p>We normally use an <b>apostrophe</b> followed by <b>the letter <i>s</i></b> to show that something belongs to someone or something. So if we said <i>The book is Rebecca's</i>, it would mean that the book belongs to Rebecca. The new word <i>(Rebecca's)</i> is called a <u><a>POSSESSIVE</a></u>.  " +
                    " <p><b>What's the Rule?</b>" +
                    " <ul>" +
                    " <li>To form a possessive, add an apostrophe and an s to the end of the word." +
                    " <p>Unfortunately, there are a lot of exceptions to this rule." +
                    " <li>When the word is PLURAL and ENDS WITH THE LETTER <i>S</i>, we do not put another <i>s</i> after the apostrophe.  " +
                    " <p><b>Let's look at an example:</b><br>" +
                    " <i>This is the students' cafeteria.</i> <br>" +
                    " This means the same thing as <i>This cafeteria belongs to the student<u>s</u>.</i>" +
                    " <p><li>When the word is SINGULAR and ENDS WITH THE LETTER <i>S</i>, we have a choice: We can either put another s after the apostrophe, or we can leave it out." +
                    " <p><b>Let's look at an example:</b><br>" +
                    " <i>One of the bus's wheels is loose.<br>" +
                    " One of the bus' wheels is loose.</i>" +
                    " <p><li>When using a POSSESSIVE PRONOUN (ex. yours, mine, his, hers, theirs, its, ours, whose) we do NOT use an apostrophe." +
                    " <p><b>Let's look at an example:</b><br>" +
                    " The cat licked its entire body.<br>" +
                    " Here, the body does belong to <i>it</i>. However, the word <i>its</i> is a possessive pronoun, so there is NO APOSTROPHE. The word <i>it's</i> is a contraction [link to contractions] that means <i>it is</i>." +
                    " </ul>" +
                    " <p><b>Your turn</b><br>" +
                    " Let's see how well you know apostrophes. Click on the activity button above. " +
                    " </p><br><br>" +
                    " <p align=center><b><font color=\"#990033\"><span name=dic>");
            XWPFRun run = lidParagraph.createRun();
            run.addBreak();
            run.addPicture(input, XWPFDocument.PICTURE_TYPE_JPEG, "some text", Units.toEMU(200), Units.toEMU(200)); // 200x200 pixels
            XWPFParagraph para = document.createParagraph();
            XWPFRun runBody = para.createRun();
            runBody.setText("When Trujillo was born in Cali, Colombia in 1980, her madre and abuela (her mother and grandmother) bestowed on her the only gift they could afford to give—an auspicious name—Lady Diana Trujillo. This noble name carried with it all her family’s hopes and dreams for a better future, although Trujillo will be the first to admit that she identifies with intergalactic warrior princesses more than the British royal she’s named after!");

            document.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            throw new RuntimeException(e);
        }
        System.out.println("poi.docx written successully");
    }


    private static void generateSpireDoc() {
        //Create a Document instance
        Document doc = new Document();
        //Add a section
        Section section = doc.addSection();
        section.addParagraph().appendHTML(
                "<html>" +
                        "<p>We normally use an <b>apostrophe</b> followed by <b>the letter <i>s</i></b> to show that something belongs to someone or something. So if we said <i>The book is Rebecca's</i>, it would mean that the book belongs to Rebecca. The new word <i>(Rebecca's)</i> is called a <u><a>POSSESSIVE</a></u>.  " +
                        " <p><b>What's the Rule?</b>" +
                        " <ul>" +
                        " <li>To form a possessive, add an apostrophe and an s to the end of the word." +
                        " <p>Unfortunately, there are a lot of exceptions to this rule." +
                        " <li>When the word is PLURAL and ENDS WITH THE LETTER <i>S</i>, we do not put another <i>s</i> after the apostrophe.  " +
                        " <p><b>Let's look at an example:</b><br>" +
                        " <i>This is the students' cafeteria.</i> <br>" +
                        " This means the same thing as <i>This cafeteria belongs to the student<u>s</u>.</i>" +
                        " <p><li>When the word is SINGULAR and ENDS WITH THE LETTER <i>S</i>, we have a choice: We can either put another s after the apostrophe, or we can leave it out." +
                        " <p><b>Let's look at an example:</b><br>" +
                        " <i>One of the bus's wheels is loose.<br>" +
                        " One of the bus' wheels is loose.</i>" +
                        " <p><li>When using a POSSESSIVE PRONOUN (ex. yours, mine, his, hers, theirs, its, ours, whose) we do NOT use an apostrophe." +
                        " <p><b>Let's look at an example:</b><br>" +
                        " The cat licked its entire body.<br>" +
                        " Here, the body does belong to <i>it</i>. However, the word <i>its</i> is a possessive pronoun, so there is NO APOSTROPHE. The word <i>it's</i> is a contraction [link to contractions] that means <i>it is</i>." +
                        " </ul>" +
                        " <p><b>Your turn</b><br>" +
                        " Let's see how well you know apostrophes. Click on the activity button above. " +
                        " </p><br><br>" +
                        " <p align=center><b><font color=\"#990033\"><span name=dic>" +
                        "</html>"
        );
        Paragraph para = section.addParagraph();
        //Append text to the paragraph
        TextRange text = para.appendText("Title: ");
        text.getCharacterFormat().setBold(true);
        text.getCharacterFormat().setFontSize(12);
        text.getCharacterFormat().setFontName("Times New Roman");
        para.appendText("Diana Trujillo: A Lesson in Perseverance");
        para.appendBreak(BreakType.Line_Break);
        para.appendBreak(BreakType.Line_Break);
        try (InputStream input = new URL("https://cdn.achieve3000.com/assets/content/images/KB/w/home/WorldCupWrap_Hero.jpg").openStream()) {
            para.appendPicture(input);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
//Save the result document
        doc.saveToFile("Spire.docx", FileFormat.Docx);
    }
}
