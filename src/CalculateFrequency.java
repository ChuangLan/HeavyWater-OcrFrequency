import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import org.apache.commons.lang3.StringUtils;

public class CalculateFrequency {
	
	static ArrayList<String> filePathList = new ArrayList<String>();
	static ArrayList<String> fileNameList = new ArrayList<String>();
	static HashMap<String, Integer> netFreq = new HashMap<String, Integer>();
	static ArrayList<HashMap<String, Integer>> wordFreqList = new ArrayList<HashMap<String, Integer>>();
	static int maxCol = 250;
	static int validWordNumber = 0;
	static int noiseNumber = 0;
	static int contentNumber = 0;
	static ArrayList<Integer> noiseRemoveList = new ArrayList<Integer>();
	static ArrayList<Integer> contentRemoveList = new ArrayList<Integer>();
	static ArrayList<Integer> validWordNumberList = new ArrayList<Integer>();
	static ArrayList<String> stopWordsList = new ArrayList<String>();
	
	static public void setFilePath(String path){
		//Get all xml files in the path
		File folder = new File(path);
		File[] listOfFiles = folder.listFiles();
	    for (int i = 0; i < listOfFiles.length; i++) {
	      if (listOfFiles[i].isFile()) {
	    	  String name = listOfFiles[i].getName();
	    	  if(name.endsWith(".xml")) {
	    		  filePathList.add(path + "\\" + name);
	    		  fileNameList.add(name);
	    	  }	    		  
	      }
	    }
	}
	
	 public static void calculateNetFreq(){
		if(filePathList == null ||filePathList.isEmpty())
			{
				System.out.println("filePathList is empty!");
				return;
			}
		for(String path: filePathList){
			//System.out.println(path);
			parseXml(path);
		}
	}

	private static void parseXml(String path) {
		// TODO Auto-generated method stub
		File  xmlFile = new File(path); 
		if(!xmlFile.exists()) {
			System.out.println("File doesn't exist!");
			//System.out.println(path);
			return;
		}
        try{
	        DocumentBuilderFactory  builderFactory =  DocumentBuilderFactory.newInstance(); 
	        DocumentBuilder builder;
			builder = builderFactory.newDocumentBuilder();          
	        Document doc;
			doc = builder.parse(xmlFile);            
	        doc.getDocumentElement().normalize();
	        //Calculate the freq and store it in a small list, then store it into the big hashmap
	        processDoc(doc);
 	   }catch(Exception  e){ 
	       e.printStackTrace();  	         
	   }         
	}

	private static void processDoc(Document doc) {
		// TODO Auto-generated method stub
		//HashMap<String, Integer> OCRValueList = new HashMap<String, Integer>();
		HashMap<String, Integer> wordFreq = new HashMap<String, Integer>();
		
		NodeList  nList = doc.getElementsByTagName("Span"); 
		
		int xmlValidWords = 0;
		int xmlNoiseRemoved = 0;
		int xmlContentRemoved = 0;
        for(int  i = 0 ; i<nList.getLength();i++){ 
        	Node  node = nList.item(i);
        	NodeList children = node.getChildNodes();
        	Element ele = (Element)children;
        	//text
        	NodeList children2 = ele.getElementsByTagName("Value");
        	String value = children2.item(0).getTextContent();
        	String text = normalize(value);
        	//System.out.println(i+value);
        	/**If text is "i", it is a noise word**/
        	if(!text.equals("") && text.length()!= 1){
        		/**Content words identification**/
        		if(isContent(text)){
        			contentNumber++;
        			xmlContentRemoved++;
        		}
        		else{
        			validWordNumber++;
        			xmlValidWords++;
        			if(wordFreq.containsKey(text)) {
            			wordFreq.put(text, wordFreq.get(text)+1);
                	}
            		else 
            			wordFreq.put(text, 1);
        		}       		        		
        	}
        	else {
        		//Count the number of removed words.
        		noiseNumber++;
        		xmlNoiseRemoved++;
        	}
        }
        wordFreqList.add(wordFreq);
        validWordNumberList.add(xmlValidWords);
        noiseRemoveList.add(xmlNoiseRemoved);
        contentRemoveList.add(xmlContentRemoved);
        
        /**NetFreq Initialization**/
        for(String word: wordFreq.keySet()) {
        	//System.out.println(word);
        	if(netFreq.containsKey(word))
        		netFreq.put(word, netFreq.get(word) + 1);
        	else 
        		netFreq.put(word, 1);
        }     
		
	}

	private static boolean isContent(String text) {
		// Auto-generated method stub		
		return stopWordsList.contains(text);
	}

	private static String normalize(String value) {
		String result = value;
		Pattern regex = Pattern.compile("[a-zA-Z]");
		/**Trim leading**/
		for(int i = 0; i < result.length(); i++) {
			char c = result.charAt(i);
			Matcher matcher = regex.matcher(Character.toString(c));
			if (matcher.find()){
			    result = result.substring(i);
			    break;
			}
			if(i == result.length()-1){
				result = "";
			}
		}
		/**Trim trailing**/
		for(int i = result.length()-1; i >= 0; i--) {
			char c = result.charAt(i);
			Matcher matcher = regex.matcher(Character.toString(c));
			if (matcher.find()){
			    result = result.substring(0, i+1);
			    break;
			}
		}		
		/**Remove the noise with threshold 50%**/
		result = removeNoise(result, 0.5);
		result = result.toLowerCase();
		
		return result;
	}
	private static String removeNoise(String str, double ratio) {
		// TODO Auto-generated method stub
		Pattern regex = Pattern.compile("[a-zA-Z]");
		int n = 0;
		for(char c: str.toCharArray()){
			Matcher matcher = regex.matcher(Character.toString(c));
			if (matcher.find()){
				n++;
			}
		}
		if(n < ratio * str.length())
			return "";
		return str;
	}

	private static void displayResult() {
		// TODO Auto-generated method stub
		System.out.println("filePathList number: " + Integer.toString(filePathList.size()));
		if(netFreq == null || netFreq.isEmpty()){
			System.out.println("netFreq is Empty!");
			return;
		}
		System.out.println("Key pairs below");
		System.out.println("Lines: " + Integer.toString(netFreq.size()));
		Iterator it = netFreq.entrySet().iterator();
	    while (it.hasNext()) {
	        Map.Entry pair = (Map.Entry)it.next();
	        System.out.println(pair.getKey() + ":\t" + pair.getValue());
	        it.remove(); // avoids a ConcurrentModificationException
	    }		
	}
	
	public static void createExcel(OutputStream os, int pageNumber) throws WriteException,IOException{
        WritableWorkbook workbook = Workbook.createWorkbook(os);
        WritableSheet sheet = workbook.createSheet("First Sheet",0);
    	jxl.write.NumberFormat nf = new jxl.write.NumberFormat("###"); 
        jxl.write.WritableCellFormat wcfN = new jxl.write.WritableCellFormat(nf);
    	jxl.write.NumberFormat nf1 = new jxl.write.NumberFormat("##%"); 
        jxl.write.WritableCellFormat wcfN1 = new jxl.write.WritableCellFormat(nf1);
        int row = 1;
        ArrayList<Label> Labels = new ArrayList<Label>();
    	//Col 0 and Col 1
        for(String word:netFreq.keySet()){
        	Label wordCell = new Label(0,row,word);
            jxl.write.Number freqCell = new jxl.write.Number(1, row, netFreq.get(word), wcfN); 
            sheet.addCell(freqCell);
        	sheet.addCell(wordCell);
        	Labels.add(wordCell);
        	row++;       	
        	
        	//System.out.println(wordCell.getContents());
        } 
        int finalRow = row;
    	//Initiate localNetFreq
    	int[] localNetFreq = new int[netFreq.size()];
    	for(int i = 0; i < netFreq.size(); i++){
    		localNetFreq[i] = 0;
    	}
        //After Col 3
        int col = 3;
        for(int i = maxCol * pageNumber; i < maxCol * (pageNumber+1) && i < fileNameList.size(); i++){
        	Label fileLabel = new Label(col,0,fileNameList.get(i));
        	sheet.addCell(fileLabel);
        	HashMap<String, Integer> fileMap = wordFreqList.get(i); 
        	int j = 0;//The iteration of netFreqList
    		for(Label label:Labels){
    			String word = label.getContents();
    			//System.out.println("Word: " );
    			if(fileMap.containsKey(word)){
    				int count = fileMap.get(word);
    	            jxl.write.Number newLabel = new jxl.write.Number(col,label.getRow(),count, wcfN); 
    				sheet.addCell(newLabel);
    				//local net freq ++
    				localNetFreq[j]++;
    			}
    			else {
    	            jxl.write.Number newLabel = new jxl.write.Number(col,label.getRow(), 0, wcfN); 
    				sheet.addCell(newLabel);
    			}
    			j++;
    		}
        	col++; 	
        }
        int finalCol = col;
        /**Net Freq for each xls**/
        for(int i = 0; i < netFreq.size(); i++){
            jxl.write.Number newLabel = new jxl.write.Number(2,i+1, localNetFreq[i], wcfN); 
			sheet.addCell(newLabel);
        }
        
        Label title1 = new Label(0,0,"Word List");
    	sheet.addCell(title1);
    	Label title2 = new Label(1,0,"Total Net Frequency among " + fileNameList.size() + " OCR files");
    	sheet.addCell(title2);
    	Label title3 = new Label(2,0,"Local Net Frequency among " + (finalCol - 3) + " OCR files");
    	sheet.addCell(title3);
        
		/**Add the count of noise words number **/
        Label title4 = new Label(1,finalRow,"Total Noise Number among " + fileNameList.size() + " OCR files");
    	sheet.addCell(title4);
    	Label title5 = new Label(2,finalRow,"Local Noise Number among " + (finalCol - 3) + " OCR files");
    	sheet.addCell(title5);
        Label wordRemovedLabel = new Label(0,finalRow+1,"Noise Number: ");
    	sheet.addCell(wordRemovedLabel);
    	//total noise number 
        jxl.write.Number wordRemovedNumber = new jxl.write.Number(1,finalRow+1, noiseNumber, wcfN); 
		sheet.addCell(wordRemovedNumber);
		//xml noise number
    	int localWordRemoved = 0;
    	for(int i = 0; i < finalCol - 3; i++){
    		int xmlNoiseNumber = noiseRemoveList.get(i + maxCol * pageNumber);
        	localWordRemoved += xmlNoiseNumber;
            jxl.write.Number xmlNoiseLabel = new jxl.write.Number(i+3,finalRow+1, xmlNoiseNumber, wcfN); 
			sheet.addCell(xmlNoiseLabel);
    	}
    	//local noise number
        jxl.write.Number localNoiseLabel = new jxl.write.Number(2,finalRow+1, localWordRemoved, wcfN); 
		sheet.addCell(localNoiseLabel);		
		
		/**Add the count of content words number **/
        Label title6 = new Label(1,finalRow + 2,"Total Content Number among " + fileNameList.size() + " OCR files");
    	sheet.addCell(title6);
    	Label title7 = new Label(2,finalRow + 2,"Local Content Number among " + (finalCol - 3) + " OCR files");
    	sheet.addCell(title7);
        Label contentTitleLabel = new Label(0,finalRow+3,"Content Number: ");
    	sheet.addCell(contentTitleLabel); 
    	//total content number
        jxl.write.Number contentNumberLabel = new jxl.write.Number(1,finalRow+3, contentNumber, wcfN); 
		sheet.addCell(contentNumberLabel);
		//xml content number
		int localContentRemoved = 0;
    	for(int i = 0; i < finalCol - 3; i++){
    		int xmlContentNumber = contentRemoveList.get(i + maxCol * pageNumber);
    		localContentRemoved += xmlContentNumber;
            jxl.write.Number xmlContentLabel = new jxl.write.Number(i+3,finalRow+3, xmlContentNumber, wcfN); 
			sheet.addCell(xmlContentLabel);
    	}
		//local content number
        jxl.write.Number localContentLabel = new jxl.write.Number(2,finalRow+3, localContentRemoved, wcfN); 
		sheet.addCell(localContentLabel);
		
		/**Add the quality function **/
		Label title8 = new Label(1,finalRow + 4,"Total Quality Ratio among " + fileNameList.size() + " OCR files");
    	sheet.addCell(title8);
    	Label title9 = new Label(2,finalRow + 4,"Local Quality Ratio among " + (finalCol - 3) + " OCR files");
    	sheet.addCell(title9);
        Label qualityTitleLabel = new Label(0,finalRow+5,"Quality Ratio (#noise/(#dictionary+#content)): ");
    	sheet.addCell(qualityTitleLabel); 
    	//total quality ratio
    	double totalRatio = noiseNumber / (double)(validWordNumber + contentNumber);
        jxl.write.Number qualityNumberLabel = new jxl.write.Number(1,finalRow+5, totalRatio, wcfN1); 
		sheet.addCell(qualityNumberLabel);
		//xml quality ratio
		int localValidNumber = 0;
    	for(int i = 0; i < finalCol - 3; i++){
    		int xmlNoiseNumber = noiseRemoveList.get(i + maxCol * pageNumber);
    		int xmlContentNumber = contentRemoveList.get(i + maxCol * pageNumber);
    		int xmlValidNumber = validWordNumberList.get(i + maxCol * pageNumber);
    		localValidNumber += xmlValidNumber;
    		double xmlQualityRatio = xmlNoiseNumber/(double)(xmlContentNumber + xmlValidNumber);
            jxl.write.Number xmlQualityLabel = new jxl.write.Number(i+3,finalRow+5, xmlQualityRatio, wcfN1); 
			sheet.addCell(xmlQualityLabel);
    	}
		//local quality ratio
    	double localQualityRatio = localWordRemoved/(double)(localValidNumber + localContentRemoved);
        jxl.write.Number localQualityLabel = new jxl.write.Number(2,finalRow+5, localQualityRatio, wcfN1); 
		sheet.addCell(localQualityLabel);
		
		/**Stream Operation**/
        workbook.write();
        workbook.close();
        os.flush();
        os.close();
    }
	
	private static void setStopWordList(String path) throws IOException {
		// TODO Auto-generated method stub
		// Open the file
		FileInputStream fstream = new FileInputStream(path);
		BufferedReader br = new BufferedReader(new InputStreamReader(fstream));

		String strLine;

		//Read File Line By Line
		while ((strLine = br.readLine()) != null)   {
		  // Print the content on the console
			stopWordsList.add(strLine);
			//System.out.println (strLine);
		}

		//Close the input stream
		br.close();
	}

	public static void main(String[] args) throws WriteException, IOException {
		// TODO Auto-generated method stub
		// Input can use args list. 
		String contentWordListPath = "src/stopwordslist.txt";
		String path = "F:\\Work\\Work_15_16_2\\HeavyWater\\input\\xml_samples";
		//String path = "F:\\Work\\Work_15_16_2\\HeavyWater\\input\\test_samples";
		//String path = "F:\\Work\\Work_15_16_2\\HeavyWater\\april-codeathon-master\\input-files";	
		setStopWordList(contentWordListPath);
		setFilePath(path);
		calculateNetFreq();
		//displayResult();
		
        //Caculate the maxCol 2M cells is the maximum
        maxCol = (int) ((2E6-6) / (netFreq.size() + 6)) - 3;
        
		int num = fileNameList.size()/maxCol;
		if(fileNameList.size()%maxCol != 0) num++;
		for(int i = 0; i < num; i++){
			OutputStream os = new FileOutputStream(path + "\\" + "result_" + Integer.toString(i) + ".xls"); 
			createExcel(os, i);
		}		
	}
	
}
