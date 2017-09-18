package ru.spravxls2xml;

import java.util.Date;

import org.apache.commons.codec.binary.Base64;

public class Person {
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
	private String encodedImage;
	
	public Person(String id, String podrId, String dolgnostId, String addressId, String bossPersId,
			String roomNumber, String fio, String telVnutr, String telGor, String telSot, String email,
			Date birthDate, Date vacationStartDate, Date vacationEndDate, Date missionStartdDate,
			Date missionEndDate, String encodedImage) {
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
		this.encodedImage = encodedImage;
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

	public String getEncodedImage() {
		return encodedImage;
	}

	public void setPhoto(String encodedImage) {
		this.encodedImage = encodedImage;
	}
}
