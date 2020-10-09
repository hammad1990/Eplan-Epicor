import java.io.File;
import java.io.FileOutputStream;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Epicor_reservation {

	static ArrayList<String> code2 = new ArrayList<>();
	static ArrayList<String> description2 = new ArrayList<>();
	static ArrayList<String> qty2 = new ArrayList<>();
	PreparedStatement ps=null;
	ResultSet rs1=null;
	//Statement st1=null;
	int count=1;
	public String  path2 = null;
	Connection con = null;

	public Epicor_reservation() {

		Eplan_reservation bb=new Eplan_reservation();
		/*
		 self.cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                              "Server=SRV-ERPDB;"
                              "Database=ERP10LIVE;"
                              "uid=Rep;pwd=")



'EXEC _xgetEBM_Reservation ?',(self.projectNumber,))
		 */



		try {

			////////start establish connecion with SQL server///////
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		String url = "jdbc:sqlserver://SRV-ERPDB;databaseName=ERP10LIVE";
		con = DriverManager.getConnection(url, "Rep", "");
		System.out.println("connected");
		 JOptionPane.showMessageDialog(bb.frame, "connected to SQL server");
	        ////////end establish connecion with SQL server///////

		 /////// start importing all reservarion from first item////////
		 String query="{call EXEC _xgetEBM_Reservation(?)}";
		 CallableStatement stmt1=con.prepareCall(query);
		 stmt1.setInt(1,Integer.parseInt(bb.txtitem1.getText()));

		 rs1=stmt1.executeQuery(query);

		 while(rs1.next()) {
               //  rs1.getString(4);
                 code2.add(rs1.getString(4));
		 }
             	/////////WRITING TO EXCEL /////////////////

						 JFileChooser file1=new JFileChooser();
								file1.setDialogTitle("save the exported file");
								file1.setCurrentDirectory(new File(System.getProperty("user.home")));
								int result=file1.showSaveDialog(bb.frame);
								if (result==JFileChooser.APPROVE_OPTION) {
								    File selectedfile1 = file1.getSelectedFile();
									 path2 = selectedfile1.getAbsolutePath();

									}
								XSSFWorkbook workbook1=new XSSFWorkbook();

								FileOutputStream out1=new FileOutputStream(new File(path2 +".xlsx"));
								XSSFSheet sheet1=workbook1.createSheet("result");
								Row headerrow=sheet1.createRow(0);
					            Cell cellh0 = headerrow.createCell(0);
					           cellh0.setCellValue("code");
					           Cell cellh1 = headerrow.createCell(1);
					           cellh1.setCellValue("description");
					           Cell cellh2 = headerrow.createCell(2);
					           cellh2.setCellValue("qty");
					           for (int i1 = 0; i1<code2.size(); i1++) {
									Row row1=sheet1.createRow(count);
							           Cell cell0 = row1.createCell(0);
							          cell0.setCellValue(code2.get(i1));
							         Cell cell1 = row1.createCell(1);
							          cell1.setCellValue(description2.get(i1));
							        Cell cell2 = row1.createCell(2);
							          cell2.setCellValue(qty2.get(i1));

								    count++;
									}

					          workbook1.write(out1);
					          	out1.close();


						/////////END WRITING TO EXCEL /////////////////


		 stmt1.close();

		 /////// end importing all reservarion from this item////////

		}
		catch (Exception e) {
			e.printStackTrace();
			 JOptionPane.showMessageDialog(bb.frame, "failed to connect with SQL server");
		}



	}

}
