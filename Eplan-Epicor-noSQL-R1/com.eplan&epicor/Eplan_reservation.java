import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class Eplan_reservation extends JFrame {

	 JFrame frame;
	public File selectedfile;
	public static String  path2 = null;
	FileInputStream  file2;
	HSSFWorkbook wb;
	HSSFSheet sheet;
	int count=1;
	static ArrayList<String> code1 = new ArrayList<>();
	static ArrayList<String> description1 = new ArrayList<>();
	static ArrayList<String> qty1 = new ArrayList<>();
	private JTextField txtorderno;
	public JTextField txtitem1;
	private JTextField txtmodel;
	private JTextField txtcapacity;
	public JTextField txtitem2;

	public Eplan_reservation() {
		getContentPane().setLayout(null);


		txtorderno = new JTextField();
		txtorderno.setBounds(112, 25, 86, 20);
		getContentPane().add(txtorderno);
		txtorderno.setColumns(10);

		JLabel lblOrderNo = new JLabel("Order No.");
		lblOrderNo.setBounds(10, 28, 75, 14);
		getContentPane().add(lblOrderNo);

		txtitem1 = new JTextField();
		txtitem1.setColumns(10);
		txtitem1.setBounds(112, 63, 86, 20);
		getContentPane().add(txtitem1);

		JLabel lblItem = new JLabel("item");
		lblItem.setBounds(10, 66, 75, 14);
		getContentPane().add(lblItem);

		txtmodel = new JTextField();
		txtmodel.setColumns(10);
		txtmodel.setBounds(112, 130, 86, 20);
		getContentPane().add(txtmodel);

		JLabel lblUnitModel = new JLabel("Unit Model");
		lblUnitModel.setBounds(10, 133, 75, 14);
		getContentPane().add(lblUnitModel);

		txtcapacity = new JTextField();
		txtcapacity.setColumns(10);
		txtcapacity.setBounds(112, 172, 86, 20);
		getContentPane().add(txtcapacity);

		JLabel lblCapacity = new JLabel("capacity");
		lblCapacity.setBounds(10, 175, 75, 14);
		getContentPane().add(lblCapacity);

		JButton btnStart = new JButton("start");
		btnStart.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent arg0) {

				String eplan_sheet_path;
				eplan_sheet_path="Z:\\MOHD RAFIQ\\project\\Reservation sheets";//Z:\MOHD RAFIQ\project\Reservation sheets
				File dir = new File(eplan_sheet_path);//D:\\HTML programs\\java script test\\JS
				FilenameFilter filter = new FilenameFilter() {
					@Override
					public boolean accept(File dir, String name) {

						return name.contains(txtorderno.getText()) && name.contains(txtmodel.getText())
								&& name.contains(txtcapacity.getText()) && name.contains(txtitem1.getText());

					}
				};
                  if (txtcapacity.getText().equalsIgnoreCase("")||txtitem1.getText().equalsIgnoreCase("")||
					txtmodel.getText().equalsIgnoreCase("")||txtorderno.getText().equalsIgnoreCase(""))

					{
                	  JOptionPane.showMessageDialog(frame, "enter missing data");
				}
             else {
				String[] children = dir.list(filter);
				if (children.length==0) {
					 JOptionPane.showMessageDialog(frame, "NO EPLAN FILE FOUND");
				} else {
					for (int i = 0; i < children.length; i++) {
						selectedfile = dir.getAbsoluteFile();
						String filename = selectedfile.getAbsolutePath();

						try {

							file2 = new FileInputStream(new File(filename + "\\" + children[i]));
							wb = new HSSFWorkbook(file2);
							sheet = wb.getSheetAt(0);

							Iterator rowIterator = sheet.rowIterator();
							Iterator rowIterator1 = sheet.rowIterator();
							Iterator rowIterator2 = sheet.rowIterator();
							rowIterator.next();// skip fist row in eplan sheet
							rowIterator.next();// skip second row in eplan sheet

							//////// getting code////////////
							while (rowIterator.hasNext()) {

								Row nextRow = (Row) rowIterator.next();
								Iterator cellIterator = nextRow.cellIterator();

								while (cellIterator.hasNext()) {
									Cell nextCell = (Cell) cellIterator.next();

									int columnIndex = nextCell.getColumnIndex();

									switch (columnIndex) {
									case 1:
										code1.add(nextCell.getStringCellValue());
										//System.out.println(code1);
										break;

									}

								}
							}
							// end getting code////////////
							///// getting description////////////
							rowIterator1.next();
							rowIterator1.next();
							while (rowIterator1.hasNext()) {

								Row nextRow = (Row) rowIterator1.next();
								Iterator cellIterator = nextRow.cellIterator();

								while (cellIterator.hasNext()) {
									Cell nextCell = (Cell) cellIterator.next();

									int columnIndex = nextCell.getColumnIndex();

									switch (columnIndex) {

									case 3:
										description1.add(nextCell.getStringCellValue());
										//System.out.println(description1);

									}

								}
							}
							/// end getting description/////////////////

							/// getting qty/////////////////

							rowIterator2.next();
							rowIterator2.next();
							while (rowIterator2.hasNext()) {

								Row nextRow = (Row) rowIterator2.next();
								Iterator cellIterator = nextRow.cellIterator();

								while (cellIterator.hasNext()) {
									Cell nextCell = (Cell) cellIterator.next();

									int columnIndex = nextCell.getColumnIndex();

									switch (columnIndex) {

									case 10:

										qty1.add(nextCell.getStringCellValue());
										//System.out.println(qty1);
									}

								}
							}
					/////////// end getting qty/////////////////
						/////////WRITING TO EXCEL /////////////////

					/*		 JFileChooser file1=new JFileChooser();
								file1.setDialogTitle("save the exported file");
								file1.setCurrentDirectory(new File(System.getProperty("user.home")));
								int result=file1.showSaveDialog(frame);
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
					           for (int i1 = 0; i1<code1.size(); i1++) {
									Row row1=sheet1.createRow(count);
							           Cell cell0 = row1.createCell(0);
							          cell0.setCellValue(code1.get(i1));
							         Cell cell1 = row1.createCell(1);
							          cell1.setCellValue(description1.get(i1));
							        Cell cell2 = row1.createCell(2);
							          cell2.setCellValue(qty1.get(i1));



								    count++;
									}






					          workbook1.write(out1);
					          	out1.close();


						/////////END WRITING TO EXCEL /////////////////



					 			*/

							JOptionPane.showMessageDialog(frame, "Eplan sheet imported ok");//
							setVisible(false);

							Epicor_reservation er =new Epicor_reservation();


						} catch (IOException e1) {

							 e1.printStackTrace();
						}
					}
					}
				}

			}
		});
		btnStart.setBounds(208, 200, 89, 23);
		getContentPane().add(btnStart);

		txtitem2 = new JTextField();
		txtitem2.setColumns(10);
		txtitem2.setBounds(112, 99, 86, 20);
		getContentPane().add(txtitem2);

		JLabel label = new JLabel("item");
		label.setBounds(10, 102, 75, 14);
		getContentPane().add(label);
		setVisible(true);
		setSize(493, 272);
		setLocation(700, 200);
	}
}
