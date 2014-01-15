import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;


public class Export extends HttpServlet {
    public Export() {
        super();
    }

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

        String data = request.getParameter("body");
        byte[] isoChars = data.getBytes("ISO-8859-1");
        String newData = new String(isoChars, "UTF-8");
        System.out.println(newData);


// 创建Excel的工作书册 Workbook,对应到一个excel文档
        HSSFWorkbook wb = new HSSFWorkbook();
//        FileOutputStream fileOut = new FileOutputStream("Export.xls");

// 处理数据//////////////////////////////////////////////
//        String newData = newData.replaceAll("\\[","").replaceAll("\\]\\]","").replaceAll("\"","").replaceAll(", ",",");
        String[] dataArray = newData.split("\\*!");
        String[] sheetName = dataArray[1].split("\\|!");
        String[] sheets = dataArray[2].split("`!");
        System.out.print("excel_name:");
        System.out.println(dataArray[0]);
        for (int k = 0; k < sheets.length; k++) {
            System.out.print("sheet_name" + k + ":");
            System.out.println(sheetName[k]);
            System.out.print("sheet_content" + k + ":");
            System.out.println(sheets[k]);
            HSSFSheet sheet = wb.createSheet(sheetName[k]);
            String[] result = sheets[k].split("~!");
            for (int i = 0; i < result.length; i++) {
                String[] each_row = result[i].split("\\|!");
                HSSFRow row = sheet.createRow(i);
                row.setHeight((short) 500);
                for (int j = 0; j < each_row.length; j++) {
                    HSSFCell cell = row.createCell(j);
                    cell.setCellValue(each_row[j]);
                }
            }
        }

        response.reset();
        response.setContentType("application/vnd.ms-Export;charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + dataArray[0]);
        wb.write(response.getOutputStream());
        response.getOutputStream().flush();
        response.getOutputStream().close();
    }

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

    }

}
