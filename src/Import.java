import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import java.io.*;

public class Import extends HttpServlet {
    public Import() {
        super();
    }

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

        String data = request.getParameter("body");
        byte[] isoChars = data.getBytes("ISO-8859-1");
        String newData = new String(isoChars, "UTF-8");
        System.out.print("file_path:");
        System.out.println(newData);
//        读取excel
        InputStream is = new FileInputStream(newData);
        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(is));

        ExcelExtractor extractor = new ExcelExtractor(wb);
        extractor.setIncludeSheetNames(false);
        extractor.setFormulasNotResults(false);
        extractor.setIncludeCellComments(true);

        String text = extractor.getText();
        System.out.println(text);
        byte[] outChars = text.getBytes("UTF-8");
        String output = new String(outChars, "ISO-8859-1");
//        传回数据
        response.setContentType("text/html; charset=ISO-8859-1");
        response.reset();
        PrintWriter writer = response.getWriter();
        writer.println(output);
        writer.close();

    }

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

    }

}
