package ru.spravxls2xml;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
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
	private static List<Office> offices = new ArrayList<>();
	
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
		Workbook book = new XSSFWorkbook(new FileInputStream(fileName));
		Sheet sheet = book.getSheet("Лист1");
		Iterator<Row> iterator = sheet.iterator();
		
		iterator.next();
		int rowNumber = 1;
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();
			processRow(currentRow, rowNumber);
			rowNumber++;
		}
		     
        book.close();
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
			podrIdElement.setNodeValue(podr.id);
			
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
			
			positionsElement.appendChild(dolgnostElement);
		}
		
		Element officesElement = doc.createElement("Offices");
		organizationElement.appendChild(officesElement);
		
		Element personalsElement = doc.createElement("Personals");
		organizationElement.appendChild(personalsElement);
		
		for (Person person : people) {
			Element persElement = doc.createElement("Pers");
			personalsElement.appendChild(persElement);
			
			Element persIdElement = doc.createElement("ID");
			persIdElement.appendChild(doc.createTextNode(person.id));
			persElement.appendChild(persIdElement);
			
			Element persFioElement = doc.createElement("FIO");
			persFioElement.appendChild(doc.createTextNode(person.fio));
			persElement.appendChild(persFioElement);
		}
		
		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
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
		String telGor = "+7411249" + getCellValue(kvantCell);
		//					: "+74112" + (gtsCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC ? gtsCell.getNumericCellValue() : gtsCell.getStringCellValue());
		String telSot = getCellValue(cellPhonetCell);
		String email = emailCell.getStringCellValue();
		Date birthDate = null;
		Date vacationStartDate = null;
		Date vacationEndDate = null;
		Date missionStartdDate = null;
		Date missionEndDate = null;
		Base64 photo = null;
		
		Person person = new Person(personId, podrId, dolgnostId, addressId, bossPersId, 
				roomNumber, fio, telVnutr, telGor, telSot, email,
				birthDate, vacationStartDate, vacationEndDate, missionStartdDate, missionEndDate, photo);
		
		String pPodrId = PODR_PREF + departments.size();
		String pPodrName = getCellValue(deptCell);
		int pRank = 100;
		String pParentPodrId = "";
		String pBossPersId = "";
		Podr podr = new Podr(pPodrId, pPodrName, pRank, pParentPodrId, pBossPersId);
		if (!departments.contains(podr)) {
			departments.add(podr);
			person.setPodrId(podr.id);
		} else {
			person.setPodrId(departments.get(departments.indexOf(podr)).id);
		}
		
		String dDolgnostId = DOLGNOST_PREF + positions.size();
		String dDognostName = positionCell.getStringCellValue();
		int dRank = 200;
		Dolgnost dolgnost = new Dolgnost(dDolgnostId, dDognostName, dRank);
		if (!positions.contains(dolgnost)) {
			positions.add(dolgnost);
			person.setPodrId(dolgnost.id);
		} else {
			person.setPodrId(positions.get(positions.indexOf(dolgnost)).id);
		}
		
		people.add(person);
	}
	
	private static String getCellValue(Cell cell) {
		if (cell == null) {
			return "";
		}
		return cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC ? String.valueOf(cell.getNumericCellValue()) : cell.getStringCellValue();
	}
	
	private static class Person {
		private String id;
		private String podrId;
		private String dolgnostId;
		private String addressId;
		private String bossPersId;
		private String roomNumber;
		private String fio;
		private String telVnutr;
		private String telGor;
		private String telSot;
		private String email;
		private Date birthDate;
		private Date vacationStartDate;
		private Date vacationEndDate;
		private Date missionStartdDate;
		private Date missionEndDate;
		private Base64 photo;
		
		public Person(String id, String podrId, String dolgnostId, String addressId, String bossPersId,
				String roomNumber, String fio, String telVnutr, String telGor, String telSot, String email,
				Date birthDate, Date vacationStartDate, Date vacationEndDate, Date missionStartdDate,
				Date missionEndDate, Base64 photo) {
			super();
			this.id = id;
			this.podrId = podrId;
			this.dolgnostId = dolgnostId;
			this.addressId = addressId;
			this.bossPersId = bossPersId;
			this.roomNumber = roomNumber;
			this.fio = fio;
			this.telVnutr = telVnutr;
			this.telGor = telGor;
			this.telSot = telSot;
			this.email = email;
			this.birthDate = birthDate;
			this.vacationStartDate = vacationStartDate;
			this.vacationEndDate = vacationEndDate;
			this.missionStartdDate = missionStartdDate;
			this.missionEndDate = missionEndDate;
			this.photo = photo;
		}

		public String getId() {
			return id;
		}

		public void setId(String id) {
			this.id = id;
		}

		public String getPodrId() {
			return podrId;
		}

		public void setPodrId(String podrId) {
			this.podrId = podrId;
		}

		public String getDolgnostId() {
			return dolgnostId;
		}

		public void setDolgnostId(String dolgnostId) {
			this.dolgnostId = dolgnostId;
		}

		public String getAddressId() {
			return addressId;
		}

		public void setAddressId(String addressId) {
			this.addressId = addressId;
		}

		public String getBossPersId() {
			return bossPersId;
		}

		public void setBossPersId(String bossPersId) {
			this.bossPersId = bossPersId;
		}

		public String getRoomNumber() {
			return roomNumber;
		}

		public void setRoomNumber(String roomNumber) {
			this.roomNumber = roomNumber;
		}

		public String getFio() {
			return fio;
		}

		public void setFio(String fio) {
			this.fio = fio;
		}

		public String getTelVnutr() {
			return telVnutr;
		}

		public void setTelVnutr(String telVnutr) {
			this.telVnutr = telVnutr;
		}

		public String getTelGor() {
			return telGor;
		}

		public void setTelGor(String telGor) {
			this.telGor = telGor;
		}

		public String getTelSot() {
			return telSot;
		}

		public void setTelSot(String telSot) {
			this.telSot = telSot;
		}

		public String getEmail() {
			return email;
		}

		public void setEmail(String email) {
			this.email = email;
		}

		public Date getBirthDate() {
			return birthDate;
		}

		public void setBirthDate(Date birthDate) {
			this.birthDate = birthDate;
		}

		public Date getVacationStartDate() {
			return vacationStartDate;
		}

		public void setVacationStartDate(Date vacationStartDate) {
			vacationStartDate = vacationStartDate;
		}

		public Date getVacationEndDate() {
			return vacationEndDate;
		}

		public void setVacationEndDate(Date vacationEndDate) {
			vacationEndDate = vacationEndDate;
		}

		public Date getMissionStartdDate() {
			return missionStartdDate;
		}

		public void setMissionStartdDate(Date missionStartdDate) {
			missionStartdDate = missionStartdDate;
		}

		public Date getMissionEndDate() {
			return missionEndDate;
		}

		public void setMissionEndDate(Date missionEndDate) {
			missionEndDate = missionEndDate;
		}

		public Base64 getPhoto() {
			return photo;
		}

		public void setPhoto(Base64 photo) {
			this.photo = photo;
		}
		
		
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
	
	protected abstract static class BaseEntity {
		protected String id;
		protected String name;
		protected int rank;
		
		public BaseEntity(String id, String name, int rank) {
			super();
			this.id = id;
			this.name = name;
			this.rank = rank;
		}

		@Override
		public int hashCode() {
			final int prime = 31;
			int result = 1;
			result = prime * result + ((id == null) ? 0 : id.hashCode());
			result = prime * result + ((name == null) ? 0 : name.hashCode());
			result = prime * result + rank;
			return result;
		}

		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (obj == null)
				return false;
			if (getClass() != obj.getClass())
				return false;
			BaseEntity other = (BaseEntity) obj;
			if (id == null) {
				if (other.id != null)
					return false;
			} else if (!id.equals(other.id))
				return false;
			if (name == null) {
				if (other.name != null)
					return false;
			} else if (!name.equals(other.name))
				return false;
			if (rank != other.rank)
				return false;
			return true;
		}
	}
}
