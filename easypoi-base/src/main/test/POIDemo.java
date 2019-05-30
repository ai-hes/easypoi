import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.export.ExcelExportServer;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIDemo {
    public static void main(String[] args) throws IOException {



        //for (int i = 0; i < 4; i++) {
            new Thread("thread" + "0"){
                @Override
                public void run() {
                    super.run();
                    Random random = new Random();
                    List<Person> list = new ArrayList<>(10);
                    for (int i = 0; i < 100; i++) {
                        Person person = new Person();
                        person.setName("name" + random.nextInt(1000));
                        person.setAge(i + 18);
                        person.setAddress("杭州 余杭 西湖" + i);
                        person.setPhone("139512100" + i);
                        list.add(person);
                    }

                    ExportParams exportParams = new ExportParams(
                        "用户统计" + random.nextInt(1000), null, "班级1");
                    exportParams.setType(ExcelType.XSSF);

                    ExcelExportServer excelExportServer = new ExcelExportServer();
                    Workbook workbook = new XSSFWorkbook();
                    excelExportServer.createSheet(workbook, exportParams, Person.class, list);

                    try(FileOutputStream outputStream = new FileOutputStream(
                        "C:\\Users\\wb-ah558847\\Desktop\\2019.05.23-空间设备批量修改需求\\" +  Thread.currentThread().getName() + ".xlsx");) {
                        workbook.write(outputStream);
                        outputStream.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }

                }
            }.start();
        //}


    }

    private static CellStyle getCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(Font.COLOR_RED);
        font.setFontHeightInPoints((short)12);
        style.setFont(font);
        return style;
    }
}
