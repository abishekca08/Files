package jp.co;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;

import javafx.application.Application;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.embed.swing.SwingFXUtils;
import javafx.scene.Group;
import javafx.scene.Scene;
import javafx.scene.SnapshotParameters;
import javafx.scene.chart.CategoryAxis;
import javafx.scene.chart.AreaChart;
import javafx.scene.chart.BarChart;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.PieChart;
import javafx.scene.chart.PieChart.Data;
import javafx.scene.chart.XYChart;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("restriction")
public class JavaFxChartFromExcel extends Application {
	static XSSFRow row;
	FileInputStream fis;
	int chrt_type;
	@SuppressWarnings("rawtypes")
	List keyList = new ArrayList<String>();
	@SuppressWarnings("rawtypes")
	Map buginfo = new TreeMap();
	XSSFWorkbook workBook;
	@SuppressWarnings("unchecked")
	public void readFromExcel() throws IOException {
		fis = new FileInputStream(new File("D:\\Files\\CJK\\Task\\BugSheet.xlsx"));
		workBook = new XSSFWorkbook(fis);
		XSSFSheet spreadsheet = workBook.getSheetAt(0);
		Iterator < Row > rowIterator = spreadsheet.iterator();
		rowIterator.next();
		while (rowIterator.hasNext()) 
		{
			row = (XSSFRow) rowIterator.next();
			Iterator < Cell > cellIterator = row.cellIterator();
			String ticketid = "", bugcount = "";
			while (cellIterator.hasNext()) 
			{
				Cell cell = cellIterator.next();
				if(ticketid == ""){
					ticketid = cell.getStringCellValue();
				}else{
					bugcount = cell.getStringCellValue();
				}
			}  
			buginfo.put(ticketid,bugcount);
			keyList.add(ticketid);
		}
		System.out.println("Values from excel");
		System.out.println(buginfo);
		System.out.println("\n");
		fis.close();
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void start(Stage stage) throws IOException{
		readFromExcel();
		Set set = buginfo.entrySet();
		Iterator i = set.iterator();
		Scene scene;
		System.out.println("(1)Area Chart");
		System.out.println("(2)Bar Chart");
		System.out.println("(3)Pie Chart");
		System.out.println();
		System.out.println("Enter chart type :");
		Scanner sc=new Scanner(System.in);
		chrt_type=sc.nextInt();
		Group root;
		CategoryAxis xAxis;
		NumberAxis yAxis;
		switch(chrt_type){
		case 1:
			xAxis = new CategoryAxis();
			xAxis.setCategories(FXCollections.<String> observableArrayList(keyList));
			xAxis.setLabel("Ticket ID");
			yAxis = new NumberAxis(0, 1000, 50);
			yAxis.setLabel("Bug Count");
			AreaChart linechart = new AreaChart(xAxis, yAxis);
			XYChart.Series<String, Number> series = new XYChart.Series<String, Number>();
			series.setName("No of Bugs");
			while (i.hasNext()) {
				Map.Entry me = (Map.Entry) i.next();
				series.getData().add(
						new XYChart.Data((String) me.getKey(), Integer.parseInt((String) me.getValue())));
			}
			linechart.getData().add(series);
			root = new Group(linechart);
			scene = new Scene(root, 600, 400);
			linechart.setAnimated(false);
			stage.setTitle("Line Chart");
			stage.setScene(scene);
			//stage.show();
			saveAsPng(linechart,"D:\\Files\\CJK\\Task\\chart.png");
			break;
		case 2:
			xAxis = new CategoryAxis();
			xAxis.setCategories(FXCollections.<String> observableArrayList(keyList));
			xAxis.setLabel("Ticket ID");
			yAxis = new NumberAxis(0, 1000, 50);
			yAxis.setLabel("Bug Count");
			BarChart barChart = new BarChart(xAxis, yAxis);
			XYChart.Series<String, Number> series1 = new XYChart.Series<String, Number>();
			series1.setName("No of Bugs");
			while (i.hasNext()) {
				Map.Entry me = (Map.Entry) i.next();
				series1.getData().add(
						new XYChart.Data((String) me.getKey(), Integer.parseInt((String) me.getValue())));
			}
			barChart.getData().add(series1);
			root = new Group(barChart);
			scene = new Scene(root, 600, 400);
			barChart.setAnimated(false);
			stage.setTitle("Line Chart");
			stage.setScene(scene);
			//scene.getStylesheets().add("chart.css");
			//stage.show();
			saveAsPngBar(barChart,"D:\\Files\\CJK\\Task\\chart.png");
			break;
		case 3:
			stage.setTitle("Bug Count");
			stage.setWidth(500);
			stage.setHeight(500);
			PieChart pieChart = new PieChart();
			PieChart.Data slice = null;
			while (i.hasNext()) {
				Map.Entry me = (Map.Entry) i.next();
				slice = new PieChart.Data((String) me.getKey(), Integer.parseInt((String) me.getValue()));
				pieChart.getData().add(slice);
			}
			pieChart.setClockwise(true); 
			pieChart.setLabelLineLength(50); 
			pieChart.setLabelsVisible(true); 
			pieChart.setStartAngle(180);     
			root = new Group(pieChart); 
			scene = new Scene(root, 600, 400);
			stage.setScene(scene);
			//stage.show();
			saveAsPngPie(pieChart,"D:\\Files\\CJK\\Task\\chart.png");
			break;
		default: System.out.println("Please Enter the chart type correctly");
				System.exit(0);
		}
		

	}
	@SuppressWarnings("rawtypes")
	public void saveAsPng(AreaChart linechart,String path) throws IOException{
		WritableImage image =linechart.snapshot(new SnapshotParameters(),null);
		File file= new File(path);
		try {
			ImageIO.write(SwingFXUtils.fromFXImage(image,null),"png",file);		
		} catch (IOException e) {
			e.printStackTrace();
		}
		writeToFile(3,2);
	}

	public void saveAsPngPie(PieChart linechart,String path) throws IOException{
		WritableImage image =linechart.snapshot(new SnapshotParameters(),null);
		File file= new File(path);
		try {
			ImageIO.write(SwingFXUtils.fromFXImage(image,null),"png",file);		
		} catch (IOException e) {
			e.printStackTrace();
		}
		writeToFile(13,2);
	}
	public void saveAsPngBar(BarChart linechart,String path) throws IOException{
		WritableImage image =linechart.snapshot(new SnapshotParameters(),null);
		File file= new File(path);
		try {
			ImageIO.write(SwingFXUtils.fromFXImage(image,null),"png",file);		
		} catch (IOException e) {
			e.printStackTrace();
		}
		writeToFile(3,25);
	}
	public void writeToFile(int col,int row) throws IOException{
		fis = new FileInputStream(new File("D:\\Files\\CJK\\Task\\BugSheet.xlsx"));
		workBook = new XSSFWorkbook(fis);
		InputStream is = null;
		File delfile;
		is = new FileInputStream("D:\\Files\\CJK\\Task\\chart.png");
		byte[] bytes = IOUtils.toByteArray(is);
		int pictureIdx = workBook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
		is.close();
		CreationHelper helper = workBook.getCreationHelper();
		Sheet sheet = workBook.getSheetAt(0);
		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setCol1(col);
		anchor.setRow1(row);
		Picture pict = drawing.createPicture(anchor, pictureIdx);
		pict.resize();
		try {
			delfile = new File("D:\\Files\\CJK\\Task\\chart.png");
			delfile.delete();
		} catch(Exception ec) {
			ec.printStackTrace();
		}
		String file = "D:\\Files\\CJK\\Task\\BugSheet.xlsx";
		FileOutputStream fileOut = new FileOutputStream(file);
		workBook.write(fileOut);
		fileOut.close();
	}

	public static void main(String args[]) throws IOException {
		launch(args);
	}
}
