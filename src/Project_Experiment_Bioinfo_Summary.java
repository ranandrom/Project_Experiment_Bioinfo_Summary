import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Hashtable;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.jcraft.jsch.ChannelExec;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.Session;

import ch.ethz.ssh2.Connection;
import ch.ethz.ssh2.SCPClient;

public class Project_Experiment_Bioinfo_Summary {
	public static void main(String[] args) throws InterruptedException
    {
		System.out.println();
		Calendar now_star = Calendar.getInstance();
		SimpleDateFormat formatter_star = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		System.out.println("程序开始时间: "+now_star.getTime());
		System.out.println("程序开始时间: "+formatter_star.format(now_star.getTime()));
		System.out.println("===============================================");
		//System.out.println();
		System.out.println("Project_Experiment_Bioinfo_Summary.1.0.0");
		System.out.println("***********************************************");
		System.out.println();
		
		SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");//定义日期格式
		String day = formatter.format(now_star.getTime());//获取当天日期
		
		//String Output_dir_name = "\\\\wdmycloud\\anchordx_cloud\\杨莹莹\\项目实验生信汇总表\\" + day;
		//String Bioinfo_Path = "\\\\wdmycloud\\anchordx_cloud\\杨莹莹\\项目-生信-汇总表";
		//String Experiment_File = "\\\\wdmycloud\\anchordx_cloud\\杨莹莹\\项目进展汇总表";
		//String Bioinfo_File = Path + "\\" + FilePath;
		
		String Output_dir_name = "/wdmycloud/anchordx_cloud/杨莹莹/项目实验生信汇总表/" + day;//文件输出路径
		String Bioinfo_Path = "/wdmycloud/anchordx_cloud/杨莹莹/项目-生信-汇总表";//生信文件所在路径
		String Experiment_File = "/wdmycloud/anchordx_cloud/杨莹莹/项目进展汇总表";//实验文件所在路径
		String Newest_FileDir = null;//生信最新目录名
		String Bioinfo_File = null;//生信带路径文件名
		
		int args_len = args.length;//输入参数长度
		int log = 0;
		for(int len = 0; len < args_len; len++){
			if( args[len].equals("-B") || args[len].equals("-b") ){
				Bioinfo_Path = args[len+1];
				log = 1;
			}else if(args[len].equals("-O") || args[len].equals("-o")){
				Output_dir_name = args[len+1] + "/" + day;
			}else if(args[len].equals("-E") || args[len].equals("-e")){
				Experiment_File = args[len+1];
			}
		}
		
		my_mkdir(Output_dir_name);//创建输出目录
		
		if(log == 0){
			Newest_FileDir = Get_New_FilePAth(Bioinfo_Path);//获取最新生信文件所在目录
		}

		File Exp_file = new File(Experiment_File);
		ArrayList<String> Experiment_list = new ArrayList<String>();
		Search_Experiment_File(Exp_file, Experiment_list);//获取需要合并的实验文件列表
		
		if(log == 0){
			Bioinfo_File = Bioinfo_Path + "/" + Newest_FileDir;
		}else{
			Bioinfo_File = Bioinfo_Path;
		}
		File Bio_file = new File(Bioinfo_File);
		ArrayList<String> Bioinfo_list = new ArrayList<String>();
		Search_Bioinfo_File(Bio_file, Bioinfo_list);//获取最新生信文件列表
		
		String Bioinfo_Plasma_FilePath = null;
		String Bioinfo_Tissue_FilePath = null;
		String OutputFilePath = null;
		
		ArrayList<String> Data_Head = new ArrayList<String>();//头列表（以 '\t' 合并，前面带该表格的表格名的 String）的列表。
		ArrayList<ArrayList<String>> Experiment_Data_list = new ArrayList<ArrayList<String>>();//实验文件数据列表
		ArrayList<String> Bioinfo_Plasma_Data_list = new ArrayList<String>();//生信血浆文件数据列表
		ArrayList<String> Bioinfo_Tissue_Data_list = new ArrayList<String>();//生信组织文件数据列表
		
		//循环合并每个实验表
		for(int i = 0; i  < Experiment_list.size(); i++){
			//System.out.println(Experiment_list.get(i));
			String Experiment_str[] = Experiment_list.get(i).split("\t");
			
			File Experiment_file = new File(Experiment_str[0]);
			Data_Head.clear();
			Data_Head = readHead(Experiment_file);
			//OutputFilePath = dir_name + "\\" + Experiment_str[1] + "_项目实验生信汇总表_" + day + ".xlsx";
			OutputFilePath = Output_dir_name + "/" + Experiment_str[1] + "_项目实验生信汇总表_" + day + ".xlsx";
			File OutPutfile = new File(OutputFilePath);
			CreateXlsx(OutPutfile, Data_Head);//创建合并文件，表的个数与 Experiment_file 里的表的个数相同。
			
			Experiment_Data_list.clear();
			Experiment_Data_list = read_Experiment_Xlsx(Experiment_file);//读实验表数据
			
			for(int j = 0;j  < Bioinfo_list.size(); j++){
				//System.out.println(Bioinfo_list.get(j));
				String Bioinfo_str[] = Bioinfo_list.get(j).split("\t");
				if(Experiment_str[1].equals(Bioinfo_str[2])){					
					if(Bioinfo_str[1].equals("Plasma")){
						Bioinfo_Plasma_FilePath = Bioinfo_str[0];
						File Bioinfo_Plasma_file = new File(Bioinfo_Plasma_FilePath);
						Bioinfo_Plasma_Data_list.clear();
						Bioinfo_Plasma_Data_list = read_Bioinfo_Xlsx(Bioinfo_Plasma_file);//读血浆表数据
					}else if(Bioinfo_str[1].equals("Tissue")){
						Bioinfo_Tissue_FilePath = Bioinfo_str[0];
						File Bioinfo_Tissue_file = new File(Bioinfo_Tissue_FilePath);
						Bioinfo_Tissue_Data_list.clear();
						Bioinfo_Tissue_Data_list = read_Bioinfo_Xlsx(Bioinfo_Tissue_file);//读组织表数据
					}
				}
			}
			Merge_File(OutPutfile, Experiment_Data_list, Bioinfo_Plasma_Data_list, Bioinfo_Tissue_Data_list);//合并文件
		}
		
		Calendar now_end = Calendar.getInstance();
		SimpleDateFormat formatter_end = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		System.out.println();
		System.out.println("==============================================");
		System.out.println("程序结束时间: "+now_end.getTime());
		System.out.println("程序结束时间: "+formatter_end.format(now_end.getTime()));
		System.out.println();
    }
	
	//查找生信文件
	public static void Search_Bioinfo_File(File des_file, ArrayList<String> list)
    {	
		try{
			for (File pathname : des_file.listFiles())
			{
				if (pathname.isFile()) //如果是文件，则判断是否需要记录
				{
					//获取文件的绝对路径
					String Folder = pathname.getParent();					
					//获取文件名（basename）
					String FileName = pathname.getName();
					
					if( !(FileName.startsWith("~$")) && !(FileName.startsWith("."))  ){
						String FileName_str[] = FileName.split("_");
						//String str = Folder + "\\" + FileName + "\t" + FileName_str[0]+ "\t" + FileName_str[1];
						String str = Folder + "/" + FileName + "\t" + FileName_str[0]+ "\t" + FileName_str[1];
						list.add(str);
					}
					continue;
				}else{
					//如果是目录，则递归
					Search_Experiment_File(pathname, list);
				}
			}
		}catch(Exception e){
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
	
	//查找实验文件
	public static void Search_Experiment_File(File des_file, ArrayList<String> list)
    {	
		try{
			for (File pathname : des_file.listFiles())
			{
				if (pathname.isFile()) //如果是文件，则判断是否需要记录
				{
					//获取文件的绝对路径
					String Folder = pathname.getParent();					
					//获取文件名（basename）
					String FileName = pathname.getName();
					
					if( !(FileName.startsWith("~$")) && !(FileName.startsWith("."))  ){
						String FileName_str[] = FileName.split("-");
						if(FileName_str[1].contains("项目进展")){
							//String str = Folder + "\\" + FileName + "\t" + FileName_str[0];
							String str = Folder + "/" + FileName + "\t" + FileName_str[0];
							list.add(str);
						}
					}
					continue;
				}else{
					//如果是目录，则递归
					Search_Experiment_File(pathname, list);
				}
			}
		}catch(Exception e){
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
	
	//获取最新文件所在目录
	public static String Get_New_FilePAth(String Path)
	{
		File file = new File(Path);
		int daynum = 0;
		for (File dir : file.listFiles()){
			if (dir.isDirectory()) { //如果是目录
				String dir_name = dir.getName(); //目录名
				if(daynum < Integer.valueOf(dir_name)){
					daynum = Integer.valueOf(dir_name);
				}else{
					continue;
				}
			}else{
				continue;
			}
		}
		return String.valueOf(daynum);
	}
	
	//合并文件
	public static void Merge_File(File Outputfile, ArrayList<ArrayList<String>> Experiment_Data_list, ArrayList<String> Bioinfo_Plasma_Data_list, ArrayList<String> Bioinfo_Tissue_Data_list)
	{
		ArrayList<String> Data_list = new ArrayList<String>();
		ArrayList<String> Bioinfo_Data = new ArrayList<String>();
		ArrayList<String> Bioinfo_remove = new ArrayList<String>();
		ArrayList<String> Bioinfo_Data_list = new ArrayList<String>();
		
		try{
			FileInputStream is = new FileInputStream(Outputfile);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			//XSSFSheet sheet = workbook.getSheetAt(0);	//获取第1个工作薄
			XSSFSheet sheet = null;
			int Sheet_Num = workbook.getNumberOfSheets();//获取工作薄个数
			//System.out.println(Sheet_Num);
			
			for(int numSheet = 0; numSheet < Sheet_Num; numSheet++ ){
				sheet = workbook.getSheetAt(numSheet);	//获取工作薄
				//String Sheet_Name = sheet.getSheetName();//获取当前工作薄名字
				
				Bioinfo_Data_list.clear();
				for(int j = 0; j < Experiment_Data_list.get(numSheet).size(); j++){
					String Experiment[] = Experiment_Data_list.get(numSheet).get(j).split("\t");
					if( !(Experiment[1].equals("null")) ){
						String Pre_lib_name[] = Experiment[1].split("-");
						if(Pre_lib_name.length > 1){
							if( Pre_lib_name[Pre_lib_name.length-1].startsWith("PS") ){
								Bioinfo_Data_list = Bioinfo_Plasma_Data_list;
								break;
							}else if( Pre_lib_name[Pre_lib_name.length-1].startsWith("BC") ){
								break;
							}else if( Pre_lib_name[Pre_lib_name.length-1].startsWith("F") ){
								Bioinfo_Data_list = Bioinfo_Tissue_Data_list;
								break;
							}else{
								continue;
							}
						}
					}
				}
				
				String null_data = null;
				String Experiment[] = Experiment_Data_list.get(numSheet).get(0).split("\t");
				for(int j = 0; j < Experiment.length; j++){
					if(j == 0){
						null_data = "null";
					}else{
						null_data += "\t" + "null";
					}
				}
				int rownum = 1;
				Bioinfo_Data.clear();
				Bioinfo_remove.clear();
				//写数据
				for(int j = 0; j < Experiment_Data_list.get(numSheet).size(); j++){

					String data = null;
					int log = 0;
					String str_Experiment[] = Experiment_Data_list.get(numSheet).get(j).split("\t");
					Data_list.clear();
					for(int x = 0; x < Bioinfo_Data_list.size(); x++){
						String str_Bioinfo[] = Bioinfo_Data_list.get(x).split("\t");
						if(str_Experiment[1].equals(str_Bioinfo[1])){
							if(log == 0){
								data = Experiment_Data_list.get(numSheet).get(j) + "\t" + Bioinfo_Data_list.get(x);
							}else{
								data = null_data + "\t" + Bioinfo_Data_list.get(x);
							}
							Data_list.add(data);
							Bioinfo_Data.add(Bioinfo_Data_list.get(x));
							log++;
						}else{
							continue;
						}
					}
					if(Data_list.size() != 0){
						for(int x = 0; x < Data_list.size(); x++){
							XSSFRow row = sheet.createRow((short) rownum++);
							String str_row[] = Data_list.get(x).split("\t");
							for(int i = 0; i < str_row.length; i++ ){
								// 在索引0的位置创建单元格（左上端）
								XSSFCell cell = row.createCell(i);
								if(str_row[i].equals("null")){
									cell.setCellValue("");
								}else{
									cell.setCellValue(str_row[i]);
								}
							}
						}
					}else{
						XSSFRow row = sheet.createRow((short) rownum++);
						for(int i = 0; i < str_Experiment.length; i++ ){
							// 在索引0的位置创建单元格（左上端）
							XSSFCell cell = row.createCell(i);
							if(str_Experiment[i].equals("null")){
								cell.setCellValue("");
							}else{
								cell.setCellValue(str_Experiment[i]);
							}
						}
					}
				}
				for(int x = 0; x < Bioinfo_Data_list.size(); x++){
					if( !(Bioinfo_Data.contains(Bioinfo_Data_list.get(x))) ){
						Bioinfo_remove.add(Bioinfo_Data_list.get(x));
					}
				}
				for(int x = 0; x < Bioinfo_remove.size(); x++){
					XSSFRow row = sheet.createRow((short) rownum++);
					String removedata = null_data + "\t" + Bioinfo_remove.get(x);
					String str_row[] = removedata.split("\t");
					for(int i = 0; i < str_row.length; i++ ){
						// 在索引0的位置创建单元格（左上端）
						XSSFCell cell = row.createCell(i);
						if(str_row[i].equals("null")){
							cell.setCellValue("");
						}else{
							cell.setCellValue(str_row[i]);
						}
					}
				}

				// 新建一输出文件流
				FileOutputStream fOut = new FileOutputStream(Outputfile);
				// 把相应的Excel 工作簿存盘
				workbook.write(fOut);
				fOut.flush();
				// 操作结束，关闭文件
				fOut.close();
			}
			is.close();
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//读生信文件
	public static ArrayList<String> read_Bioinfo_Xlsx(File file)
	{
		ArrayList<String> Data_list = new ArrayList<String>();
		try{
			InputStream is = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(is);
			XSSFSheet sheet = wb.getSheetAt(0);	//获取第三个工作薄
			//XSSFSheet sheet = null;
			//int Sheet_Num = wb.getNumberOfSheets();//获取工作薄个数
			//System.out.println(Sheet_Num);
			//String Sheet_Name = sheet.getSheetName();//获取当前工作薄名字
			String data = null;
			// 获取当前工作薄的每一行
			for (int i = sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
		
				XSSFRow xssfrow = sheet.getRow(i);
					
				// 获取当前工作薄的每一列
				for (int j = xssfrow.getFirstCellNum(); j < xssfrow.getLastCellNum(); j++) {
					XSSFCell xssfcell = xssfrow.getCell(j);
						
					if(j == 0){
						xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
						data = xssfcell.getStringCellValue().trim();
					}else{
						xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
						data += "\t" + xssfcell.getStringCellValue().trim();
					}
				}
				Data_list.add(data);
			}
			is.close();
			wb.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return Data_list;
	}
	
	//写数据
	public static void WriteExcelData (File file, ArrayList<ArrayList<String>> Data_list)
	{
		try{
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			//XSSFSheet sheet = workbook.getSheetAt(0);	//获取第1个工作薄
			XSSFSheet sheet = null;
			int Sheet_Num = workbook.getNumberOfSheets();//获取工作薄个数
			//System.out.println(Sheet_Num);
			
			for(int numSheet = 0; numSheet < Sheet_Num; numSheet++ ){
				sheet = workbook.getSheetAt(numSheet);	//获取工作薄
				//String Sheet_Name = sheet.getSheetName();//获取当前工作薄名字
				//写数据
				for(int j = 0; j < Data_list.get(numSheet).size(); j++){
					XSSFRow row = sheet.createRow((short) j+1);
					String str_row[] = Data_list.get(numSheet).get(j).split("\t");
					for(int i = 0; i < str_row.length; i++ ){
						// 在索引0的位置创建单元格（左上端）
						XSSFCell cell = row.createCell(i);
						if(str_row[i].equals("null")){
							cell.setCellValue("");
						}else{
							cell.setCellValue(str_row[i]);
						}
					}
				}

				// 新建一输出文件流
				FileOutputStream fOut = new FileOutputStream(file);
				// 把相应的Excel 工作簿存盘
				workbook.write(fOut);
				fOut.flush();
				// 操作结束，关闭文件
				fOut.close();
			}
			is.close();
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//读实验文件
	public static ArrayList<ArrayList<String>> read_Experiment_Xlsx(File file)
	{	
		ArrayList< ArrayList<String> > Data_list = new ArrayList< ArrayList<String> >();
		try {
		InputStream is = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(is);
		//XSSFSheet sheet = wb.getSheetAt(2);	//获取第三个工作薄
		XSSFSheet sheet = null;
		int Sheet_Num = wb.getNumberOfSheets();//获取工作薄个数
		//System.out.println(Sheet_Num);
		
		String data = null;
		int celllog = 0;// 读取的最后一列。
		for(int numSheet = 0; numSheet < Sheet_Num; numSheet++ ){
			ArrayList<String> datalist = new ArrayList<String>();
			sheet = wb.getSheetAt(numSheet);	//获取工作薄
			//String Sheet_Name = sheet.getSheetName();//获取当前工作薄名字
			//System.out.println(Sheet_Name);
			
			XSSFRow xssfrow0 = sheet.getRow(0);
			// 获取当前工作薄 "TNM stage" 列前的每一列。
			for (int j = xssfrow0.getFirstCellNum(); j < xssfrow0.getLastCellNum(); j++) {
				XSSFCell xssfcell = xssfrow0.getCell(j);
				int xcellty = xssfcell.getCellType();
				String head = null;
				if( xcellty == 0 ){
					if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
						Date date = xssfcell.getDateCellValue();
						SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
						head = dateFormat.format(date);
					}else{
						xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
						HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
						String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
						head = cellFormatted;
					}
				}else{
					xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
					head = xssfcell.getStringCellValue().trim();
				}
				if( head.equals("TNM stage") ){
					celllog = j; 
					break;
				}
			}
			int nullrowlog = 0;
			// 获取当前工作薄的每一行
			for (int i = sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
				int nulllog = 0;
				XSSFRow xssfrow = sheet.getRow(i);
				//System.out.println(i);
				
				if ( xssfrow == null || (CheckRowNull(xssfrow) == 0) ) {
					for(int j = 0; j < celllog; j++){
						if(j == 0){
							data = "null";
						}else{
							data += "\t" + "null";
						}
					}
					nulllog++;
					nullrowlog++;
				}else{
					// 获取当前工作薄的每一列
					for (int j = 0; j <= celllog; j++) {
						XSSFCell xssfcell = xssfrow.getCell(j);
						
						if(j == 0 ){
							if( xssfcell == null  || (("") == String.valueOf(xssfcell)) || xssfcell.toString().equals("") || xssfcell.getCellType() == HSSFCell.CELL_TYPE_BLANK ){
								data = "null";
							}else{
								int xcellty = xssfcell.getCellType();
								if( xcellty == 0 ){
									if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
										Date date = xssfcell.getDateCellValue();
										SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
										data = dateFormat.format(date);	//以日期格式获取数据
									}else{
										xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
										HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
										String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
										data = cellFormatted;
									}
								}else{
									xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
									data = xssfcell.getStringCellValue().trim();
								}
								nulllog++;
							}
							//System.out.println(data);
						}else{
							if( xssfcell == null  || (("") == String.valueOf(xssfcell)) || xssfcell.toString().equals("") || xssfcell.getCellType() == HSSFCell.CELL_TYPE_BLANK ){
								data += "\t" +  "null";
							}else{
								int xcellty = xssfcell.getCellType();
								if( xcellty == 0 ){
									if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
										Date date = xssfcell.getDateCellValue();
										SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
										data += "\t" + dateFormat.format(date);
									}else{
										xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
										HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
										String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
										data += "\t" + cellFormatted;
									}
								}else{
									xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
									data += "\t" + xssfcell.getStringCellValue().trim();
								}
								nulllog++;
							}							
						}
					}
				}
				if(nulllog != 0){
					datalist.add(data);
					//System.out.println(data);
				}
				if(nullrowlog > 5){
					for(int n = 0; n <= 5; n++){
						datalist.remove(datalist.size()-1);
					}
					break;
				}
			}
			Data_list.add(datalist);
		}
			is.close();
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return Data_list;
	}
	
	//判断行为空,如果为空，则返回0
	public static int CheckRowNull(XSSFRow xssfRow)
	{
		int num = 0;
		// 获取当前工作薄的每一列
		for (int j = xssfRow.getFirstCellNum(); j < xssfRow.getLastCellNum(); j++) {
			XSSFCell xssfcell = xssfRow.getCell(j);
			//String cellValue = String.valueOf(xssfcell);
			if( xssfcell == null  || (("") == String.valueOf(xssfcell)) || xssfcell.toString().equals("") || xssfcell.getCellType() == HSSFCell.CELL_TYPE_BLANK ){
				continue;
			}else{
				num++;
			}
		}
		return num;
	}
	
	//读xlsx格式文件，返回标头列表（以 '\t' 合并，前面带该表格的表格名的 String）的列表。
	public static ArrayList<String> readHead(File file)
	{		
		ArrayList<String> Data_list = new ArrayList<String>();
		try {
			InputStream is = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(is);
			XSSFSheet sheet = null;
			int Sheet_Num = wb.getNumberOfSheets();//获取工作薄个数
			//System.out.println(Sheet_Num);
			
			String data = null;
			for(int numSheet = 0; numSheet < Sheet_Num; numSheet++ ){
				sheet = wb.getSheetAt(numSheet);	//获取工作薄
				String Sheet_Name = sheet.getSheetName();//获取当前工作薄名字
				//System.out.println(Sheet_Name);
				XSSFRow xssfrow = sheet.getRow(0);
					
				// 获取当前工作薄 "TNM stage" 列前的每一列。
				for (int j = xssfrow.getFirstCellNum(); j < xssfrow.getLastCellNum(); j++) {
					XSSFCell xssfcell = xssfrow.getCell(j);
					if(j == 0){
						//data = xssfcell.getStringCellValue().trim();
						int xcellty = xssfcell.getCellType();
						if( xcellty == 0 ){
							if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
								Date date = xssfcell.getDateCellValue();
								SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
								data = dateFormat.format(date);
							}else{
								xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
								HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
								String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
								data = cellFormatted;
							}
						}else{
							xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
							data = xssfcell.getStringCellValue().trim();
						}
					}else{
						//data += "\t" + xssfcell.getStringCellValue().trim();
						int xcellty = xssfcell.getCellType();
						if( xcellty == 0 ){
							if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
								Date date = xssfcell.getDateCellValue();
								SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
								data += "\t" + dateFormat.format(date);
							}else{
								xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
								HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
								String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
								data += "\t" + cellFormatted;
							}
						}else{
							xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
							data += "\t" + xssfcell.getStringCellValue().trim();
						}
					}
					int xcellty = xssfcell.getCellType();
					String head = null;
					if( xcellty == 0 ){
						if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
							Date date = xssfcell.getDateCellValue();
							SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
							head = dateFormat.format(date);
						}else{
							xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
							HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
							String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
							head = cellFormatted;
						}
					}else{
						xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
						head = xssfcell.getStringCellValue().trim();
					}
					if(head.equals("TNM stage") ){
						break;
					}
				}
				data = Sheet_Name +  "\t" + data;
				//System.out.println(data);
				Data_list.add(data);
			}
			is.close();
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return Data_list;
	}
	
	//新建合并文件
	public static void CreateXlsx(File file, ArrayList<String> Data_list)
	{
		try{
			XSSFWorkbook workbook = new XSSFWorkbook();
			for(int j = 0; j < Data_list.size(); j++){
				String str_head[] = Data_list.get(j).split("\t");
				workbook.createSheet(str_head[0]);
			}
			for(int j = 0; j < Data_list.size(); j++){
				String str_h[] = Data_list.get(j).split("\t");
				// 创建Excel的工作sheet,对应到一个excel文档的tab  
				XSSFSheet sheet = workbook.getSheet(str_h[0]);	//获取工作薄;
				// 在索引0的位置创建行（最顶端的行）
				XSSFRow row0 = sheet.createRow((short) 0);
				// 在索引0的位置创建单元格（左上端）
				//XSSFCell cell = row.createCell((short) 0);
				
				String head_row0 = "Sample ID"+"\t"+
				"Pre-lib name"+"\t"+
						"Identification name"+"\t"+
				"Sequencing info"+"\t"+
						"Sequencing file name"+"\t"+
				"Mapping%"+"\t"+
						"Total PF reads"+"\t"+
				"Mean_insert_size"+"\t"+
						"Median_insert_size"+"\t"+
				"On target%"+"\t"+
						"Pre-dedup mean bait coverage"+"\t"+
				"Deduped mean bait coverage"+"\t"+
						"Deduped mean target coverage"+"\t"+
				"% target bases > 30X"+"\t"+
						"Uniformity (0.2X mean)"+"\t"+
				"C methylated in CHG context"+"\t"+
						"C methylated in CHH context"+"\t"+
				"C methylated in CpG context"+"\t"+
						"QC result"+"\t"+
				"Date of QC"+"\t"+
						"Path to sorted.deduped.bam"+"\t"+
				"Date of path update"+"\t"+
						"Bait set"+"\t"+
				"Sample QC"+"\t"+
						"Failed QC Detail"+"\t"+
				"Warning QC Detail"+"\t"+
						"Check"+"\t"+
				"Note1"+"\t"+
						"Note2"+"\t"+
				"Note3";
				
				String head_row1 = "样本编号"+"\t"
						+"预文库样本名"+"\t"
						+ "上机"+"\t"
						+"e.g. with S01 prefix before a Pre-lib name - this is to separate multiple sequencing files derived from the same pre-library";
				
				//1、创建字体，设置其为红色：
				XSSFFont font = workbook.createFont();
				font.setColor(HSSFFont.COLOR_RED);
				font.setFontHeightInPoints((short)10);
				font.setFontName("Palatino Linotype");
				//font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				//2、创建格式
				XSSFCellStyle cellStyle= workbook.createCellStyle();
				cellStyle.setFont(font);
				cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
				cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体，设置其为粗体，背景蓝色：
				XSSFFont font1 = workbook.createFont();
				//font1.setColor(HSSFFont.COLOR_RED);
				font1.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font1.setFontHeightInPoints((short)10);
				font1.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle1= workbook.createCellStyle();
				cellStyle1.setFont(font1);
				cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle1.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
				cellStyle1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体，设置其为红色、粗体，背景绿色：
				XSSFFont font2 = workbook.createFont();
				font2.setColor(HSSFFont.COLOR_RED);
				font2.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font2.setFontHeightInPoints((short)10);
				font2.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle2= workbook.createCellStyle();
				cellStyle2.setFont(font2);
				cellStyle2.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle2.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle2.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
				cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体大小为10，背景蓝色：
				XSSFFont font3 = workbook.createFont();
				font3.setFontHeightInPoints((short)10);
				font3.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle3= workbook.createCellStyle();
				cellStyle3.setFont(font3);
				cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle3.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
				cellStyle3.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体大小为10，背景黄色：
				XSSFFont font4 = workbook.createFont();
				font4.setFontHeightInPoints((short)10);
				font4.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle4= workbook.createCellStyle();
				cellStyle4.setFont(font4);
				cellStyle4.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle4.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle4.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle4.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle4.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
				cellStyle4.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体，设置其为粗体，背景黄色：
				XSSFFont font5 = workbook.createFont();
				//font1.setColor(HSSFFont.COLOR_RED);
				font5.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font5.setFontHeightInPoints((short)10);
				font5.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle5= workbook.createCellStyle();
				cellStyle5.setFont(font5);
				cellStyle5.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle5.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle5.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle5.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle5.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
				cellStyle5.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				
				//1、创建字体，设置其为粗体，背景橘色：
				XSSFFont font6 = workbook.createFont();
				//font6.setColor(HSSFFont.COLOR_RED);
				font6.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font6.setFontHeightInPoints((short)10);
				font6.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle6 = workbook.createCellStyle();
				cellStyle6.setFont(font6);
				cellStyle6.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle6.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle6.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle6.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle6.setFillForegroundColor(HSSFColor.TAN.index);
				cellStyle6.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				
				//1、创建字体，设置其为粗体，红字，背景橘色：
				XSSFFont font7 = workbook.createFont();
				font7.setColor(HSSFFont.COLOR_RED);
				font7.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font7.setFontHeightInPoints((short)10);
				font7.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle7 = workbook.createCellStyle();
				cellStyle7.setFont(font7);
				cellStyle7.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle7.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle7.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle7.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle7.setFillForegroundColor(HSSFColor.TAN.index);
				cellStyle7.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
								 
				String str_head = Data_list.get(j) + "\t" + head_row0;
				String str_Data[] = Data_list.get(j).split("\t");
				String str_head_row0[] = str_head.split("\t");
				// 在单元格中输入一些内容
				for(int i = 1; i < str_head_row0.length; i++ ){
					// 在索引0的位置创建单元格（左上端）
					XSSFCell cell = row0.createCell(i-1);
					if( i < 4 ){// 实验表格的 "Sample ID" ～ "Sequencing info"：红字橘底
						cell.setCellValue(str_head_row0[i]);
						cell.setCellStyle(cellStyle7);
					}else if( i >= str_Data.length && i < str_Data.length+3 ){// 生信表格的 "Sample ID" ～ "Sequencing info"：红字绿底。
						cell.setCellValue(str_head_row0[i]);
						cell.setCellStyle(cellStyle2);
					}else if( i == str_head_row0.length-10 || i == str_head_row0.length-9 ){// "Path to sorted.deduped.bam"、"Date of path update"：黑字黄底。
						cell.setCellStyle(cellStyle5);
						cell.setCellValue(str_head_row0[i]);
					}else if(i > str_Data.length){// 剩下的生信表格的列：黑字蓝底。
						cell.setCellStyle(cellStyle1);
						cell.setCellValue(str_head_row0[i]);
					}else{// 剩下的部分（实验表格的列）：黑子橘底。
						cell.setCellStyle(cellStyle6);
						cell.setCellValue(str_head_row0[i]);
					}
				}
				/*XSSFRow row1 = sheet.createRow((short) 1);
				String str_head_row1[] = head_row1.split("\t");
				for(int i = 0; i < str_head_row0.length; i++ ){
					// 在索引0的位置创建单元格（左上端）
					XSSFCell cell = row1.createCell(i);
					if( i < str_head_row1.length ){
						if (i < 3 ){
							cell.setCellStyle(cellStyle);
							cell.setCellValue(str_head_row1[i]);
						}else{
							cell.setCellStyle(cellStyle3);
							cell.setCellValue(str_head_row1[i]);
						}
					}else if( i == str_head_row0.length-3 || i == str_head_row0.length-2 ){
						cell.setCellStyle(cellStyle4);
					}else{
						cell.setCellStyle(cellStyle3);
					}
				}*/
				// 新建一输出文件流
				FileOutputStream fOut = new FileOutputStream(file);
				// 把相应的Excel 工作簿存盘
				workbook.write(fOut);
				fOut.flush();
				// 操作结束，关闭文件
				fOut.close();
				//System.out.println("文件生成...");
				//is.close();
			}
			workbook.close();
		}catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//创建目录
	public static void my_mkdir(String dir_name)
	{
		File file = new File( dir_name );
		
		//如果文件不存在，则创建
		if(!file.exists() && !file.isDirectory()){
			//System.out.println("//目录不存在");
			file.mkdirs();
		}
	}
	
	//用SSh查找文件
	public static void SSh_Find_File(String command, ArrayList<String> filelist)
	{
		String user = "zhirong_lu";
		String pass = "zhirong_lu";
		String host = "192.192.192.220";
		int port = 22;
		try {
			//String command = "mkdir " + PutPath;
			JSch jsch = new JSch();
			Session session = jsch.getSession(user, host, port);
			Hashtable<String, String> config = new Hashtable<String, String>();
            config.put("StrictHostKeyChecking", "no");
			//session.setConfig("StrictHostKeyChecking","no");
			session.setConfig(config);
			session.setPassword(pass);
			session.connect();
			/*ChannelExec channelExec1 = (ChannelExec)session.openChannel("exec");
			channelExec1.setCommand(command1);
			channelExec1.connect();*/
			ChannelExec channelExec = (ChannelExec)session.openChannel("exec");
			InputStream in = channelExec.getInputStream();
			channelExec.setCommand(command);
			channelExec.connect();
		
			channelExec.setInputStream(null);
            BufferedReader input = new BufferedReader(new InputStreamReader(channelExec.getInputStream()));
			//channelExec.setErrStream(System.err);
			channelExec.connect();
			//接收远程服务器执行命令的结果
            String line;
            while ((line = input.readLine()) != null) {  
                //System.out.println(line);
                
                File file = new File(line);
                String file_name = file.getName();
                //System.out.println("file_name: "+file_name);
	            if( !(file_name.startsWith("~$")) && !(file_name.startsWith("._")) ){
	            	filelist.add(line);
	                //System.out.println("SSh: "+new String(line.getBytes("gbk"),"gbk"));
	            }
            }  
            input.close(); 
			//String out = IOUtils.toString(in, "UTF-8");
			channelExec.disconnect();
			session.disconnect();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}	

	//用SSh上传文件
	public static void SSh_UDload_File(String filename, String PutPath, int putupordownload)
	{
		String user = "zhirong_lu";
		String pass = "zhirong_lu";
		String host = "192.192.192.220";
		//int port = 22;
		try {
			Connection con = new Connection(host);
			con.connect();
			boolean isAuthed = con.authenticateWithPassword(user, pass); 
			//System.out.println("文件已上传："+filename);
			SCPClient scpClient = con.createSCPClient();
			if( putupordownload == 0 ){
				scpClient.put(filename, PutPath);//从本地复制文件到远程目录
			}else{
				scpClient.get(new String(filename.getBytes("gbk"),"gbk"), PutPath);//从远程获取文件
			}
			
			//scpClient.get("/wdmycloud/anchordx_cloud/杨莹莹/项目-生信-汇总表/项目-生信-进展更新-模板.xlsx","./66");//从远程获取文件
			con.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
