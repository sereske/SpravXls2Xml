package ru.spravxls2xml;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Main {
	
	private static final String PERSON_PREF = "SE_PERSON_";
	private static final String OFFICE_PREF = "SE_OFFICE_";
	private static final String DOLGNOST_PREF = "SE_DOLGNOST_";
	private static final String PODR_PREF = "SE_PODR_";
	
	private static List<Person> people = new ArrayList<>();
	private static List<Podr> departments = new ArrayList<>();
	private static List<Dolgnost> positions = new ArrayList<>();
	private static List<Office> offices = Arrays.asList(
			new Office(OFFICE_PREF + 0, "677001, г. Якутск, пер. Энергетиков, д. 2, Исполнительная дирекция", 1000), 
			new Office(OFFICE_PREF + 1, "677004, г. Якутск, ул. Беринга, д. 42, Производственный центр", 100)
			);
	
	public static void main(String[] args) {
		try {
			readFromExcel("C:\\sprav.xlsx");
			writeIntoXml();
		} catch (IOException exc) {
			exc.printStackTrace();
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		} catch (TransformerException e) {
			e.printStackTrace();
		}
	}
	
	public static void readFromExcel(String fileName) throws IOException {
		XSSFWorkbook book = new XSSFWorkbook(new FileInputStream(fileName));
		XSSFSheet sheet = book.getSheet("Лист1");
		Iterator<Row> iterator = sheet.iterator();
		
		iterator.next();
		int rowNumber = 1;
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();
			processRow(currentRow, rowNumber);
			rowNumber++;
		}
		processImages(sheet);
		     
        book.close();
	}
	
	public static void processImages(XSSFSheet sheet) {
		List<XSSFShape> shapes = sheet.getDrawingPatriarch().getShapes();
		for (XSSFShape shape : shapes) {
		    if (shape instanceof XSSFPicture) {
		        XSSFPicture picture = (XSSFPicture) shape;
		        XSSFClientAnchor anchor = (XSSFClientAnchor) picture.getAnchor();
		        // Ensure to use only relevant pictures
		        if (anchor.getCol1() == 0) {
		            // Use the row from the anchor
		            XSSFRow pictureRow = sheet.getRow(anchor.getRow1());
		            if (pictureRow != null) {
		            	int rowNum = pictureRow.getRowNum();
		            	Person person = people.get(rowNum - 1);
		            	byte[] picData = picture.getPictureData().getData();
		            	String encodedImage = Base64.encodeBase64String(picData);
		            	person.setPhoto(encodedImage);
		            	people.set(rowNum - 1, person);
		            }
		        }
		    }
		}
	}
	
	public static void writeIntoXml() throws ParserConfigurationException, TransformerException {
		DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
		
		Document doc = docBuilder.newDocument();
		Element organizationElement = doc.createElement("Organization");
		doc.appendChild(organizationElement);
		
		Element nameElement = doc.createElement("Name");
		nameElement.appendChild(doc.createTextNode("АО \"Сахаэнерго\""));
		organizationElement.appendChild(nameElement);
		
		Element innElement = doc.createElement("INN");
		innElement.appendChild(doc.createTextNode("1435117944"));
		organizationElement.appendChild(innElement);
		
		Element kppElement = doc.createElement("KPP");
		kppElement.appendChild(doc.createTextNode("144950001"));
		organizationElement.appendChild(kppElement);
		
		Element xmlUpdateDateElement = doc.createElement("XMLUpdateDate");
		xmlUpdateDateElement.appendChild(doc.createTextNode(LocalDate.now().format(DateTimeFormatter.ISO_LOCAL_DATE)));
		organizationElement.appendChild(xmlUpdateDateElement);
		
		Element xmlResponsiblePersId = doc.createElement("XMLResponsiblePers_ID");
		xmlResponsiblePersId.appendChild(doc.createTextNode("111"));
		organizationElement.appendChild(xmlResponsiblePersId);
		
		Element orgStructure = doc.createElement("OrgStructure");
		organizationElement.appendChild(orgStructure);
		
		for (Podr podr : departments) {
			Element podrElement = doc.createElement("Podr");
			
			Element podrIdElement = doc.createElement("ID");
			podrIdElement.appendChild(doc.createTextNode(podr.id));
			
			Element parentPodrIdElement = doc.createElement("ParentPodr_ID");
			parentPodrIdElement.appendChild(doc.createTextNode(podr.parentPodrId));
			
			Element bossPersIdElement = doc.createElement("BossPers_ID");
			bossPersIdElement.appendChild(doc.createTextNode(podr.bossPersId));
			
			Element podrNameElement = doc.createElement("Name");
			podrNameElement.appendChild(doc.createTextNode(podr.name));
			
			Element podrRankElement = doc.createElement("Rank");
			podrRankElement.appendChild(doc.createTextNode(String.valueOf(podr.rank)));
			
			podrElement.appendChild(podrIdElement);
			podrElement.appendChild(parentPodrIdElement);
			podrElement.appendChild(bossPersIdElement);
			podrElement.appendChild(podrNameElement);
			podrElement.appendChild(podrRankElement);
			
			orgStructure.appendChild(podrElement);
		}
		
		Element positionsElement = doc.createElement("Positions");
		organizationElement.appendChild(positionsElement);
		
		for (Dolgnost dolgnost : positions) {
			Element dolgnostElement = doc.createElement("Dolgnost");
			
			Element dolgnostIdElement = doc.createElement("ID");
			dolgnostIdElement.appendChild(doc.createTextNode(dolgnost.id));
			
			Element dolgnostNameElement = doc.createElement("Name");
			dolgnostNameElement.appendChild(doc.createTextNode(dolgnost.name));
			
			Element dolgnostRankElement = doc.createElement("Rank");
			dolgnostRankElement.appendChild(doc.createTextNode(String.valueOf(dolgnost.rank)));
			
			
			dolgnostElement.appendChild(dolgnostIdElement);
			dolgnostElement.appendChild(dolgnostNameElement);
			dolgnostElement.appendChild(dolgnostRankElement);
			
			positionsElement.appendChild(dolgnostElement);
		}
		
		Element officesElement = doc.createElement("Offices");
		organizationElement.appendChild(officesElement);
		
		for (Office office : offices) {
			Element addressElement = doc.createElement("Address");
			
			Element addressIdElement = doc.createElement("ID");
			addressIdElement.appendChild(doc.createTextNode(office.id));
			
			Element addressNameElement = doc.createElement("Name");
			addressNameElement.appendChild(doc.createTextNode(office.name));
			
			Element rankElement = doc.createElement("Rank");
			rankElement.appendChild(doc.createTextNode(String.valueOf(office.rank)));
			
			addressElement.appendChild(addressIdElement);
			addressElement.appendChild(addressNameElement);
			addressElement.appendChild(rankElement);
			
			officesElement.appendChild(addressElement);
		}
		
		Element personalsElement = doc.createElement("Personals");
		organizationElement.appendChild(personalsElement);
		
		for (Person person : people) {
			Element persElement = doc.createElement("Pers");
			personalsElement.appendChild(persElement);
			
			Element persIdElement = doc.createElement("ID");
			persIdElement.appendChild(doc.createTextNode(person.getId()));
			persElement.appendChild(persIdElement);
			
			Element podrIdElement = doc.createElement("Podr_ID");
			podrIdElement.appendChild(doc.createTextNode(person.getPodrId()));
			persElement.appendChild(podrIdElement);
			
			Element dolgnostIdElement = doc.createElement("Dolgnost_ID");
			dolgnostIdElement.appendChild(doc.createTextNode(person.getDolgnostId()));
			persElement.appendChild(dolgnostIdElement);
			
			Element addressIdElement = doc.createElement("Address_ID");
			addressIdElement.appendChild(doc.createTextNode(person.getAddressId()));
			persElement.appendChild(addressIdElement);
			
			Element bossPersIdElement = doc.createElement("BossPers_ID");
			bossPersIdElement.appendChild(doc.createTextNode(person.getBossPersId()));
			persElement.appendChild(bossPersIdElement);
			
			Element roomNumberElement = doc.createElement("RoomNumber");
			roomNumberElement.appendChild(doc.createTextNode(person.getRoomNumber()));
			persElement.appendChild(roomNumberElement);
			
			Element persFioElement = doc.createElement("FIO");
			persFioElement.appendChild(doc.createTextNode(person.getFio()));
			persElement.appendChild(persFioElement);
			
			Element telVnutrElement = doc.createElement("TelVnutr");
			telVnutrElement.appendChild(doc.createTextNode(person.getTelVnutr()));
			persElement.appendChild(telVnutrElement);
			
			Element telGorElement = doc.createElement("TelGor");
			telGorElement.appendChild(doc.createTextNode(person.getTelGor()));
			persElement.appendChild(telGorElement);
			
			Element telSotElement = doc.createElement("TelSot");
			telSotElement.appendChild(doc.createTextNode(person.getTelSot()));
			persElement.appendChild(telSotElement);
			
			Element emailElement = doc.createElement("Email");
			emailElement.appendChild(doc.createTextNode(person.getEmail()));
			persElement.appendChild(emailElement);
			
			Element birthDateElement = doc.createElement("BirthDate");
			String birthDateString = person.getBirthDate() == null ? "" : person.getBirthDate().toString();
			birthDateElement.appendChild(doc.createTextNode(birthDateString));
			persElement.appendChild(birthDateElement);
			
			Element vacationStartDateElement = doc.createElement("VocationStartDate");
			String vacationStartDateString = person.getVacationStartDate() == null ? "" : person.getVacationStartDate().toString();
			vacationStartDateElement.appendChild(doc.createTextNode(vacationStartDateString));
			persElement.appendChild(vacationStartDateElement);
			
			Element vacationEndDateElement = doc.createElement("VocationEndDate");
			String vacationEndDateString = person.getVacationEndDate() == null ? "" : person.getVacationEndDate().toString();
			vacationEndDateElement.appendChild(doc.createTextNode(vacationEndDateString));
			persElement.appendChild(vacationEndDateElement);
			
			Element missionStartDateElement = doc.createElement("MissionStartDate");
			String missionStartDateString = person.getMissionStartdDate() == null ? "" : person.getMissionStartdDate().toString();
			missionStartDateElement.appendChild(doc.createTextNode(missionStartDateString));
			persElement.appendChild(missionStartDateElement);
			
			Element missionEndDateElement = doc.createElement("MissionEndDate");
			String missionEndDateString = person.getMissionEndDate() == null ? "" : person.getMissionEndDate().toString();
			missionEndDateElement.appendChild(doc.createTextNode(missionEndDateString));
			persElement.appendChild(missionEndDateElement);
			
			Element photoElement = doc.createElement("Photo");
			String encodedImage = person.getEncodedImage() == null ? "" : person.getEncodedImage();
			photoElement.appendChild(doc.createTextNode(encodedImage));
			persElement.appendChild(photoElement);
		}
		
		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		DOMSource source = new DOMSource(doc);
		StreamResult result = new StreamResult(new File("C:\\file.xml"));
		
		// Output to console for testing
		// StreamResult result = new StreamResult(System.out);

		transformer.transform(source, result);

		System.out.println("File saved!");
	}
	
	public static void processRow(Row row, int rowNumber) {
		Cell photoCell = row.getCell(0);
		Cell fioCell = row.getCell(1);
		Cell positionCell = row.getCell(2);
		Cell deptCell = row.getCell(3);
		Cell roomCell = row.getCell(4);
		Cell kvantCell = row.getCell(5);
		Cell ipCell = row.getCell(6);
		Cell gtsCell = row.getCell(7);
		Cell cellPhonetCell = row.getCell(8);
		Cell emailCell = row.getCell(9);
		
		String personId = PERSON_PREF + rowNumber;
		String podrId = "";
		String dolgnostId = "";
		String addressId = "";
		String bossPersId = "";
		String roomNumber = getCellValue(roomCell);
		String fio = fioCell.getStringCellValue();
		String telVnutr = getCellValue(ipCell);
		String telGor = getCellValue(kvantCell).equals("")?  "" : "+7411249" + getCellValue(kvantCell);
		//					: "+74112" + (gtsCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC ? gtsCell.getNumericCellValue() : gtsCell.getStringCellValue());
		String telSot = getCellValue(cellPhonetCell);
		String email = emailCell.getStringCellValue();
		Date birthDate = null;
		Date vacationStartDate = null;
		Date vacationEndDate = null;
		Date missionStartdDate = null;
		Date missionEndDate = null;
		String encodedString = null;
		
		Person person = new Person(personId, podrId, dolgnostId, addressId, bossPersId, 
				roomNumber, fio, telVnutr, telGor, telSot, email,
				birthDate, vacationStartDate, vacationEndDate, missionStartdDate, missionEndDate, encodedString);
		
		String pPodrId = null;//PODR_PREF + departments.size();
		String pPodrName = getCellValue(deptCell);
		int pRank = 100;
		String pParentPodrId = "";
		String pBossPersId = "";
		Podr podr = new Podr(pPodrId, pPodrName, pRank, pParentPodrId, pBossPersId);
		if (!departments.contains(podr)) {
			podr.id = PODR_PREF + departments.size();
			departments.add(podr);
			person.setPodrId(podr.id);
		} else {
			person.setPodrId(departments.get(departments.indexOf(podr)).id);
		}
		
		String dDolgnostId = null;//DOLGNOST_PREF + positions.size();
		String dDognostName = positionCell.getStringCellValue();
		int dRank = 200;
		Dolgnost dolgnost = new Dolgnost(dDolgnostId, dDognostName, dRank);
		if (!positions.contains(dolgnost)) {
			dolgnost.id = DOLGNOST_PREF + positions.size();
			positions.add(dolgnost);
			person.setDolgnostId(dolgnost.id);
		} else {
			person.setDolgnostId(positions.get(positions.indexOf(dolgnost)).id);
		}
		
		people.add(person);
	}
	
	private static String getCellValue(Cell cell) {
		if (cell == null) {
			return "";
		}
		return cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC ? String.valueOf(cell.getNumericCellValue()) : cell.getStringCellValue();
	}
	
	private static class Podr extends BaseEntity {
		private String parentPodrId;
		private String bossPersId;
		
		public Podr(String id, String name, int rank, String parentPodrId, String bossPersId) {
			super(id, name, rank);
			this.parentPodrId = parentPodrId;
			this.bossPersId = bossPersId;
		}

		@Override
		public int hashCode() {
			final int prime = 31;
			int result = super.hashCode();
			result = prime * result + ((bossPersId == null) ? 0 : bossPersId.hashCode());
			result = prime * result + ((parentPodrId == null) ? 0 : parentPodrId.hashCode());
			return result;
		}

		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (!super.equals(obj))
				return false;
			if (getClass() != obj.getClass())
				return false;
			Podr other = (Podr) obj;
			if (bossPersId == null) {
				if (other.bossPersId != null)
					return false;
			} else if (!bossPersId.equals(other.bossPersId))
				return false;
			if (parentPodrId == null) {
				if (other.parentPodrId != null)
					return false;
			} else if (!parentPodrId.equals(other.parentPodrId))
				return false;
			return true;
		}
				
	}
	
	private static class Dolgnost extends BaseEntity {

		public Dolgnost(String id, String name, int rank) {
			super(id, name, rank);
		}
		
	}
	
	private static class Office extends BaseEntity {

		public Office(String id, String name, int rank) {
			super(id, name, rank);
		}
		
	}
}
