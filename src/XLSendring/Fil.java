package XLSendring;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Fil {
	private static String path = "C:\\Users\\Matia\\Documents\\Horde-Exc\\";

	public static void main(String[] args)
			throws FileNotFoundException, IOException, EncryptedDocumentException, InvalidFormatException {
		try (InputStream inp = new FileInputStream(path + "TransaskjonslisteTest.xls")) {
			Workbook wb = WorkbookFactory.create(inp);
			System.out.println("test");
			Sheet sheet = wb.getSheetAt(0);

			for (int i = 1; i < 2010; i++) {
				Row row = sheet.getRow(i);
				Cell cell = row.getCell(5);
				Cell Forklaring = row.getCell(1);
				String f = Forklaring.getStringCellValue();
				Cell b = row.getCell(3);
				
				double belop = 0.0;
				if(b != null) {
					 belop = b.getNumericCellValue();
				}
				
				f = f.toUpperCase();
//				Husleie eller Fast overføringering mellom egne konti
				if(f.contains("KONTOREGULERING")) {
					if(cell == null) {
						cell = row.createCell(5);
					}
					cell.setCellValue(2);
				}
//				Faste utgifter
				if(
						f.contains("TELIA") ||
						f.contains("YA Bank") ||
						f.contains("RESURS BANK") ||
						f.contains("SKYSS") && belop > 400 || 
						f.contains("NETFLIX") ||
						f.contains("SPOTIFY") || 
				        f.contains("ALEKTUM") ||
				        f.contains("CONECTO")||
				        f.contains("LINDORFF") || 
				        f.contains("VIAPLAY") ||
				        f.contains("SISSELS") && belop > 250 ||
				        f.contains("TRENING") && belop > 250 || 
				        f.contains("SATS") && belop > 150 || 
				        f.contains("SERGEL") ||
				        f.contains("KOMPLETT BANK") || 
				        f.contains("SANTANDER" ) ||
				        f.contains("DEMAND") 	
				        &&
				        belop > 0
				        )
				        {
					if(cell == null) {
						cell = row.createCell(5);
					}
					cell.setCellValue(1);
				} else if ( f.contains("INNKREVINGSSENTRAL")) {
					cell.setCellValue("");
				}
				
				
			}
//				if(Forklaring.getStringCellValue().contains("Kontoregulering")) {
//				if(cell == null) {
//					cell = row.createCell(5);
//				}
//				cell.setCellValue("2");
//			}
//           cell.setCellType(CellType.STRING);
				
			
			
			try (OutputStream fileOut = new FileOutputStream("TransaskjonslisteTest.xls")) {
				wb.write(fileOut);
			}
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}
