package source;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sevkiyat {

    public Statement stmtver() throws SQLException {
        String dbURL = "jdbc:sqlserver://localhost\\sqlexpress;user=sevkiyat;password=1234";                 //veritabanı bağlantısı yapılıyor

        Connection conn = DriverManager.getConnection(dbURL);

        Statement stmt = conn.createStatement();

        return stmt;
    }

    public void add(String tmp) throws SQLException {
        try {

            Sevkiyat s = new Sevkiyat();

            s.stmtver().executeUpdate(tmp);
            System.out.println("Eklendi");
        } catch (SQLException e) {
            System.out.println(e);
        }

    }

    public void anamethod(String konum) throws IOException, SQLException {
        System.out.println("Sql Kodu Oluştuluyor");
        FileInputStream fis = null;
        Row row;
        String ek = "";
        int h = 0;
        int c=0;
        Sevkiyat sevkiyat=new Sevkiyat();

        try {
            fis = new FileInputStream(new File(konum));
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet spreadsheet = workbook.getSheetAt(0);
            Iterator< Row> rowIterator = spreadsheet.iterator();
            while (rowIterator.hasNext()) {
                row = (XSSFRow) rowIterator.next();
                Iterator< Cell> cellIterator = row.cellIterator();
if (h == 0) {
    System.out.println("Geç");
                    }else{
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    if (c == 1) {
                    String z="";
                    String f="";
                                z += cell.getNumericCellValue();
                                f="'"+z.substring(0,2)+":"+z.substring(2,4)+":"+z.substring(4,6)+"',";
                                ek+=f;
                                c++;
                               
                    } else {
                        c++;
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_STRING:
                                ek += "'" + cell.getStringCellValue() + "',";
                               
                                break;

                            case Cell.CELL_TYPE_NUMERIC:
                           
                         
                                ek += cell.getNumericCellValue()+ ",";
                              
                                break;
                        }

                    }
                }}
               
              

             
          h++;  }      String tmp = "insert into Sevkiyat (plakano,cıkıs_zamanı,M_adı,siparis_no,Bobin_no) values(";
                    ek = ek.substring(0, ek.length() - 1);
                    tmp += ek + ")";
           
                    sevkiyat.add(tmp);

            fis.close();
        } catch (FileNotFoundException ex) {

        } finally {
            System.out.println();
        }
    }

  

    public void basla() throws IOException, SQLException {
        String yol = "c:\\excel_test";
        Sevkiyat s1 = new Sevkiyat();
        File klasor = new File(yol);
        if (klasor.exists()) {
            String[] dosyalar = klasor.list();
            for (String dosyalar1 : dosyalar) {
                String geçici = yol + "\\" + dosyalar1;
                
                if (dosyalar1.substring(0,1).equals("9")) {
                    
                    System.out.println("Atla");
                }
                
                else{
                    
                    s1.anamethod(geçici);
                    File f2 = new File("c:\\excel_test\\"+dosyalar1);
                    File f1 = new File("c:\\excel_test\\"+"9"+dosyalar1);
                    boolean b = f2.renameTo(f1);
                }
                
            }
        }
    }
}
