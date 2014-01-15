import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.io.FileInputStream;
import java.io.InputStream;

import java.io.*;

public class Test extends HttpServlet {
    public Test() {
        super();
    }

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

        String text = "数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据数据";
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
