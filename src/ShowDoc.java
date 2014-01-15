import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import java.io.*;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;


import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.w3c.dom.Document;


public class ShowDoc extends HttpServlet {

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        String data = request.getParameter("path");
        byte[] isoChars = data.getBytes("ISO-8859-1");
        String newData = new String(isoChars, "UTF-8");

        String picPath = request.getParameter("pic");
        byte[] picChars = picPath.getBytes("ISO-8859-1");
        final String newPicPath = new String(picChars, "UTF-8");

        final HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(newData));
        WordToHtmlConverter wordToHtmlConverter = null;
        try {
            wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
            wordToHtmlConverter.setPicturesManager(new PicturesManager() {
                public String savePicture(byte[] content,
                                          PictureType pictureType, String suggestedName,
                                          float widthInches, float heightInches) {
                    System.out.println(suggestedName);
                    String path = newPicPath + suggestedName;
                    try {
                        FileOutputStream out = new FileOutputStream(new File(path));
                        try {
                            out.write(content);
                        } catch (java.io.IOException e) {
                            e.printStackTrace();
                        }
                    } catch (java.io.FileNotFoundException e) {
                        e.printStackTrace();
                    }
                    return "word_pic?pic_name=" + suggestedName;

                }
            });
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        }

        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(out);

        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = null;
        try {

            serializer = tf.newTransformer();
        } catch (TransformerConfigurationException e) {
            e.printStackTrace();
        }
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        try {
            serializer.transform(domSource, streamResult);
        } catch (TransformerException e) {
            e.printStackTrace();
        }
        out.close();
        String result = new String(out.toByteArray());
        response.reset();
        response.setContentType("text/html;charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename=word");
        response.getWriter().println(result);
    }


    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

    }

}




