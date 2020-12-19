import java.io.*;
import java.util.*;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SpreadSheet {
	
	public SpreadSheet() {
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		XSSFSheet sheet = workbook.createSheet("2k ScoreSheet");
		
		Map<String, Object[]> data = new TreeMap<String, Object[]>(); 
        data.put("1", new Object[]{ "Name", "Wins", "Losses","win %" }); 
        data.put("2", new Object[]{"Michael" , 0, 0,0}); 
        data.put("3", new Object[]{ "Jimmy", 0, 0,0 }); 
        data.put("4", new Object[]{ "Rowan", 0, 0, 0}); 
        data.put("5", new Object[]{ "Connor", 0, 0,0 }); 
  
        
        Set<String> keyset = data.keySet(); 
        int rownum = 0; 
        for (String key : keyset) { 
            
            Row row = sheet.createRow(rownum++); 
            Object[] objArr = data.get(key); 
            int cellnum = 0; 
            for (Object obj : objArr) { 
                
                Cell cell = row.createCell(cellnum++); 
                if (obj instanceof String) 
                    cell.setCellValue((String)obj); 
                else if (obj instanceof Integer) 
                    cell.setCellValue((Integer)obj); 
            } 
        }
   
        
        try { 
            
            FileOutputStream out = new FileOutputStream(new File("Fedigan2KScoreSheet.xlsx")); 
            workbook.write(out); 
            out.close(); 
            System.out.println("Fedigan2KScoreSheet.xlsx written successfully on disk."); 
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } 
	}
 public static void showScoreBoard() throws IOException {
		
		
		String excelFilePath = "Fedigan2KScoreSheet.xlsx";
		try {
			FileInputStream in = new FileInputStream(new File(excelFilePath));
			XSSFWorkbook scoresheet = new XSSFWorkbook(in);
			
			
			XSSFSheet newSheet = scoresheet.getSheetAt(0);
			Iterator<Row> iterator = newSheet.iterator();
			
			while(iterator.hasNext()) {
				Row r=iterator.next();
				Iterator<Cell> iteratorCol= r.cellIterator();
				while(iteratorCol.hasNext()) {
					Cell c = iteratorCol.next();
					org.apache.poi.ss.usermodel.CellType t= c.getCellType();
					switch (c.getCellType()) {
						case NUMERIC:
							
							System.out.print(c.getNumericCellValue()+"\t");
							break;
						case STRING:
							System.out.print(c.getStringCellValue()+"\t");
							break;
						
					
				}
					
				}
				
				System.out.println();
			}
			scoresheet.close();
			in.close();
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
 public static void enterScore() throws IOException {
	System.out.println("enter game score, line by line");
	Scanner scan = new Scanner(System.in);
///intakes the name or each player and the score he/she had
	String winName=scan.next();
	int winScore=Integer.parseInt(scan.next());
	String loseName=scan.next();
	int loseScore=Integer.parseInt(scan.next());
/// ensure that the winner name + score is assigned/reference to correct variable; swaps if necessary
	if(loseScore>winScore) {
		
		int temp=loseScore;
		loseScore=winScore;
		winScore=temp;
		
		String tem=loseName;
		loseName=winName;
		winName=tem;
	} 
	
//To overwrite everything	
	String excelFilePath = "Fedigan2KScoreSheet.xlsx";
	try {
	///DECLARATIONS OF ALL THE VARIABLES
		FileInputStream in = new FileInputStream(new File(excelFilePath));
		XSSFWorkbook scoresheet = new XSSFWorkbook(in);
		XSSFSheet newSheet = scoresheet.getSheetAt(0);
		
		Iterator<Row> iterator = newSheet.iterator();
		
		while(iterator.hasNext()) {
			Row r = iterator.next();
			if(r.getCell(0).getStringCellValue().equalsIgnoreCase(winName)) {
				Cell wins= r.getCell(1);
				wins.setCellValue(wins.getNumericCellValue()+1);
			}
			else if(r.getCell(0).getStringCellValue().equalsIgnoreCase(loseName)) {
				Cell loser= r.getCell(2);
				loser.setCellValue(loser.getNumericCellValue()+1);
			}
		}
		in.close();
		
		FormulaEvaluator evaluator = scoresheet.getCreationHelper().createFormulaEvaluator();
		for(int i=1;i<5;i++) {
			
	        String formula;
	        if(newSheet.getRow(i).getCell(1).getNumericCellValue()+newSheet.getRow(i).getCell(1).getNumericCellValue()!=0){
	        	formula ="(ROUND(B"+(i+1)+"/SUM(B"+(i+1)+"+C"+(i+1)+"),2))";
	        	newSheet.getRow(i).getCell(3).setCellFormula(formula);
	        	
		        evaluator.evaluateInCell(newSheet.getRow(i).getCell(3));
	        }
	        else {
	        	newSheet.getRow(i).getCell(3).setCellValue(0);
	        }
	       
	       
	        }
		FileOutputStream out = new FileOutputStream(new File(excelFilePath)); 
		scoresheet.write(out);
		out.close();
		scoresheet.close();
	
		
		
	} catch (FileNotFoundException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} 
	
	

	
	
}

public static void addPlayer() throws IOException {
	
	
	try {
		System.out.println("Enter a player's name");
		
		Scanner scan = new Scanner(System.in);
		String name = scan.nextLine();
		String excelFilePath = "Fedigan2KScoreSheet.xlsx";
		
		FileInputStream in = new FileInputStream("Fedigan2KScoreSheet.xlsx");
		XSSFWorkbook scoresheet= new XSSFWorkbook(in);
		XSSFSheet sheet =scoresheet.getSheetAt(0);
		
		
		Object[] arrObj= {name,0,0,0};
		int rowNumber= sheet.getLastRowNum();
		int count=0;
		Row row=sheet.createRow(rowNumber+1);
		
		for(Object obj:arrObj) {
			Cell cell = row.createCell(count);
			if (obj instanceof String) 
                cell.setCellValue((String)obj); 
            else if (obj instanceof Integer) 
                cell.setCellValue((Integer)obj); 
			count++;
		}
		
		
		FileOutputStream out = new FileOutputStream(new File(excelFilePath)); 
		scoresheet.write(out);
		out.close();
		scoresheet.close();
		System.out.println(name+"appended sucessfully to file");
		
	}
	catch(FileNotFoundException x) {
		x.printStackTrace();
	}
	
}
 public static void displayRankings() throws IOException {
	
	 Map<String,Double> map = new TreeMap<String,Double>();
	 FileInputStream in = new FileInputStream(new File("Fedigan2KScoreSheet.xlsx"));
	 XSSFWorkbook scoresheet = new XSSFWorkbook(in);
	 XSSFSheet newSheet = scoresheet.getSheetAt(0);
	 Object[][] list= new Object[newSheet.getLastRowNum()][2];
	
	
	
	
	try {
		
		for(int i =1;i<=newSheet.getLastRowNum();i++) {
			
			Cell c = newSheet.getRow(i).getCell(0);
			String name = c.getStringCellValue();
			
			c=newSheet.getRow(i).getCell(3);
			Double i1=  c.getNumericCellValue();
			
			//map.put(name,i1);
			Object[] obj = {name, i1};
			list[i-1]=obj;
			
		}
		scoresheet.close();
		in.close();
		
	} catch (FileNotFoundException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} 
	
	
	// One by one move boundary of unsorted subarray  
    for (int i = 0; i < list.length-1; i++)  
    {  
        // Find the minimum element in unsorted array  
        int min_idx=i;
        for (int j = i+1; j < list.length; j++)  {
        	double x= (double)list[i][1];
        	double y = (double)list[j][1];
        	if(x<y) {
        		Object[] temp = list[i];
        		list[i]=list[j];
        		list[j]=temp;
        		
        	}
        	
        }
       
        	
    }  
    
    for(int x=0;x<list.length;x++) {
    	System.out.println((x+1)+". "+ list[x][0]);
    }
 
	
	
	
} 
	

	//Stream<Map.Entry<String ,Double>> sorted =map.entrySet().stream().sorted((Map.Entry.comparingByValue()));
	//sorted.forEach(System.out.println());
	
	
	
	
	
	
	
}


