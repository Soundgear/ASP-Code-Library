<%
class clsProperty
	' Private variables
		private c_AID
		private aptRS
		private amenRS
		private c_FloorplanOrder
		private c_SpecialsShowDate
		private flag_BulletinDone
		private flag_WelcomeTextDone
		private flag_GuestCardListDone
		private haserrors

	' Public (Private) variables
		public HasSpecial
		public MFirmID
		public Special
		public SpecialsLink
		public SpecialFontFamily
		public SpecialFontColor
		public SpecialTitleFontColor
		public SpecialFontSize
		public SpecialTitleFontSize
		public SpecialOuterBgColor
		public SpecialInnerBgColor
		public Name
		public Manager
		public Address
		public Address2
		public City
		public State
		public Zip
		public Phone
		public PhoneUnformatted
		public Fax
		public FaxUnformatted
		public Email
		public EmailCC
		public URL
		public Mapsco
		public MondayHours
		public TuesdayHours
		public WednesdayHours
		public ThursdayHours
		public FridayHours
		public SaturdayHours
		public SundayHours
		public c_WelcomeText
		public PetPolicy
		public c_BulletinBoard
		public HasEBrochure
		public HasVMoveIn
		public HasOnlineApp
		public Directions
		public Latitude
		public Longitude
		public CommunityFeatures
		public c_MITSAmen
		
	' Public Arrays
		public Amenities
		public UnitAmenities
		public CommunityAmenities
		public StimActAmenities
		public InvActAmenities
		public InsActAmenities
		public SocActAmenities
		public Floorplans		
		public VTours
		
	' Define specific properties
		public property let setAID(i_AID)
			c_AID = i_AID
			flag_BulletinDone = false
			flag_WelcomeTextDone = false
			flag_GuestCardListDone = false
			GetInfo()
		end property
		public property let FloorplanOrder(order)
			c_FloorplanOrder = order
			GetInfo()
		end property
		public property get AID()
			AID = c_AID
		end property
		public property get FloorplanOrder()
			FloorplanOrder = c_FloorplanOrder
		end property
		public property get WelcomeText()
			GetWelcomeText()
			WelcomeText = c_WelcomeText
		end property
		public property get BulletinBoard()
			GetBulletin()
			BulletinBoard = c_BulletinBoard
		end property
		public property let MITSAmen (i_value)
			c_MITSAmen = i_value
			GetAmenities()
		end property
		public property let SpecialsShowDate (i_value)
			c_SpecialsShowDate = i_value
			SetSpecialsLink()
		end property
		public property get ApplySpecialSetting ()
			SetSpecialsLink()
		end property
		

	public sub Class_Initialize()
		c_SpecialsShowDate = true
		c_FloorplanOrder = "WEIGHT"
		c_MITSAmen = true
		haserrors = false		
	end sub
	
	private function GetInfo()
		if c_AID = "" OR isnumeric(c_Aid)=False then
			classError("The .setAID property is required and has not been set.")
			response.End()
		end if
		' Define all the properties of this property after AID is set
			sql = "SELECT * FROM APTS WHERE ApartmentID = " & CLng(c_AID)
			set aptRS = AptConnRead(sql)
			
			sqlAmen = "SELECT * FROM AptAmenities WHERE AID = " & CLng(c_AID)
			set amenRS = AptConnRead(sqlAmen)

			
		' Get Direct Variables
			set LibraryRS = aptRS
			set LibraryRS2 = amenRS
			
			Name = ""
			Manager = ""
			Address = ""
			Address2 = ""
			City = ""
			State = ""
			Zip = ""
			Phone = ""
			Fax = ""
			Email = ""
			EmailCC = ""
			URL = ""
			Mapsco = ""
			MondayHours = ""
			TuesdayHours = ""
			WednesdayHours = ""
			ThursdayHours = ""
			FridayHours = ""
			SaturdayHours = ""
			SundayHours = ""
			PetPolicy = ""
			WelcomeTextAPTSTable = ""
			HasEBrochure = false
			HasVMoveIn = false
			HasOnlineApp = false
			Directions = ""
			HasSpecial = false
			Special = ""
			SpecialsLink = ""
			SpecialFontFamily = "Times New Roman, Times, serif"
			SpecialFontColor = "000000"
			SpecialTitleFontColor = "000000"
			SpecialFontSize = "13px"
			SpecialTitleFontSize="16px"
			SpecialOuterBgColor = "FFFFFF"
			SpecialInnerBgColor = "FFFFFF"
			Latitude = ""
			Longitude = ""
			CommunityFeatures = ""
			MFirmID = 0
			if not LibraryRS.eof then
				Name = HTMLDecode(LibraryRS("APARTMENT"))
				MFirmID = LibraryRS("PMFID")
				Manager = LibraryRS("MANAGER")
				Address = LibraryRS("ADDRESS1")
				Address2 = LibraryRS("ADDRESS2")
				HasSpecial = LibraryRS("special")
				HasSpecialCheck=LibraryRS("SpecialDateActive")
				HasSpecialCheckDate=LibraryRS("SpecialDate")
				If HasSpecial Then
					If HasSpecialCheck Then
						If HasSpecialCheckDate<Date() Then
							HasSpecial = false
						End If
					End If
				End If
				City = LibraryRS("CITY")
				State = LibraryRS("STATE")
				Zip = LibraryRS("ZIP")
				PhoneUnformatted = LibraryRS("PHONE1")
				Phone = FormatPhoneNumber(PhoneUnformatted)
				FaxUnformatted = LibraryRS("FAX")
				Fax = FormatPhoneNumber(FaxUnformatted)
				Email = LibraryRS("Email")
				EmailCC = LibraryRS("EmailCC")
				URL = LibraryRS("URL")
				Mapsco = LibraryRS("MAPSCO")
				Latitude = LibraryRS("Latitude")
				Longitude = LibraryRS("Longitude")
				MondayHours = LibraryRS("MondayHours")
				TuesdayHours = LibraryRS("TuesdayHours")
				WednesdayHours = LibraryRS("WednesdayHours")
				ThursdayHours = LibraryRS("ThursdayHours")
				FridayHours = LibraryRS("FridayHours")
				SaturdayHours = LibraryRS("SaturdayHours")
				SundayHours = LibraryRS("SundayHours")
				HasEBrochure = LibraryRS("eBrochure")
				HasVMoveIn = LibraryRS("VMoveIn")
				HasOnlineApp = LibraryRS("OnlineApp")
				Special = HTMLDecode(LibraryRS("specialmemo"))
				CommunityFeatures = LibraryRS("CommunityFeatures")
				Directions = HTMLDecode(LibraryRS("Directions"))
				PetPolicy = HTMLDecode(LibraryRS("PetPolicy2"))
				SetSpecialsLink()
			end if

		GetAmenities()
		GetFloorplans()
		GetVTours()
	end function

	private function SetSpecialsLink ()
		SpecialsLink = "javascript:;"" onclick=""window.open('/library/v4/helpers/specials_coupon/coupon.asp?AID=" & CLng(c_AID) & "&amp;showdate="&c_SpecialsShowDate&"&amp;SpecialFontFamily="&SpecialFontFamily&"&amp;SpecialFontSize="&SpecialFontSize&"&amp;SpecialTitleFontSize="&SpecialTitleFontSize&"&amp;SpecialFontColor="&SpecialFontColor&"&amp;SpecialTitleFontColor="&SpecialTitleFontColor&"&amp;SpecialOuterBgColor="&SpecialOuterBgColor&"&amp;SpecialInnerBgColor="&SpecialInnerBgColor&"','siteplan4','resizable=yes,scrollbars=yes,status=no,width=340,height=260')"
	end function

	public function GetField (inpField)
		if c_AID = "" then
			classerror("The .setAID property is required and has not been set.")
			response.End()
		end if
		sql = "SELECT " & inpField & " FROM APTS WHERE ApartmentID = " & CLng(c_AID)
		set LibraryRS = AptConnRead(sql)
		GetField = ""
		if not LibraryRS.eof then GetField = LibraryRS(inpField)
	end function
	
	private function GetBulletin ()
		if not flag_BulletinDone then
			sql = "SELECT Comments FROM Bulletin WHERE AID = " & CLng(c_AID)
			set LibraryRS = AptConnRead(sql)
			c_BulletinBoard = HTMLDecode(LibraryRS("Comments"))
			flag_BulletinDone = true
		end if
	end function

	private function GetWelcomeText ()
		if not flag_WelcomeTextDone then
			sql = "SELECT WelcomeText FROM WelcomeText WHERE AID = " & CLng(c_AID)
			set LibraryRS = AptConnRead(sql)
			if not LibraryRS.eof then
				c_WelcomeText = HTMLDecode(LibraryRS("WelcomeText"))
			else
				c_WelcomeText = HTMLDecode(GetField("WelcomeText"))
			end if
			flag_WelcomeTextDone = true
		end if
	end function

	private function GetAmenities ()
			set LibraryRS = aptRS
			set LibraryRS2 = amenRS
			Amenities = array()
			UnitAmenities = array()
			CommunityAmenities = array()
			StimActAmenities = array()
			InvActAmenities = array()
			InsActAmenities = array()
			SocActAmenities = array()
			If c_MITSAmen=true Then
				if LibraryRS("billspaid") then Amenities = addToAry(Amenities, "All Utilities Included")
				if LibraryRS("corphouse") then Amenities = addToAry(Amenities, "Corporate Housing Available")
				if LibraryRS("shortlease") then Amenities = addToAry(Amenities, "Short Term Leases Available")
				if LibraryRS("AC") then Amenities = addToAry(Amenities, "Air Conditioning")
				if LibraryRS("alarm") then Amenities = addToAry(Amenities, "Alarm Included")
				if LibraryRS("AlarmAvail") then Amenities = addToAry(Amenities, "Alarm Available")
				if LibraryRS("BALCONY") then Amenities = addToAry(Amenities, "Balcony/Patio")
				if LibraryRS("BUSINESSLINEACCESS") then Amenities = addToAry(Amenities, "Business Center")
				if LibraryRS("CABLE") then Amenities = addToAry(Amenities, "Cable Available")
				if LibraryRS("CABLEINCLUDED") then Amenities = addToAry(Amenities, "Cable Included")
				if LibraryRS("SATALITE") then Amenities = addToAry(Amenities, "Satellite Included")
				if LibraryRS("SatAvailable") then Amenities = addToAry(Amenities, "Satellite Available")
				if LibraryRS("CEILINGFANS") then Amenities = addToAry(Amenities, "Ceiling Fans")
				if LibraryRS("CONTROLLEDACCESS") then Amenities = addToAry(Amenities, "Controlled Access")
				if LibraryRS("COURTYARD") then Amenities = addToAry(Amenities, "Courtyard")
				if LibraryRS("COVEREDPARKING") then Amenities = addToAry(Amenities, "Covered Parking")
				if LibraryRS("DISABILITY") then Amenities = addToAry(Amenities, "Handicap Accessible")
				if LibraryRS("DISHWASHER") then Amenities = addToAry(Amenities, "Dishwasher")
				if LibraryRS("DISPOSAL") then Amenities = addToAry(Amenities, "Disposal")
				if LibraryRS("FIREPLACE") then Amenities = addToAry(Amenities, "Fireplace")
				if LibraryRS("STORAGE") then Amenities = addToAry(Amenities, "Extra Storage")
				if LibraryRS("ISP") then Amenities = addToAry(Amenities, "High Speed Internet Access")
				if LibraryRS("DSL") then Amenities = addToAry(Amenities, "DSL Available")
				if LibraryRS("FITNESSCENTER") then Amenities = addToAry(Amenities, "Fitness Center")
				if LibraryRS("FURNISHED") then Amenities = addToAry(Amenities, "Furnished Apartment Available")
				if LibraryRS("GARAGES") then Amenities = addToAry(Amenities, "Garages Available")
				if LibraryRS("AttachedGarage") then Amenities = addToAry(Amenities, "Attached Garages Available")
				if LibraryRS("HIGHRISE") then Amenities = addToAry(Amenities, "High Rise")
				if LibraryRS("HouseKeeping") then Amenities = addToAry(Amenities, "Complimentary Housekeeping Available")
				if LibraryRS("ICEMAKERS") then Amenities = addToAry(Amenities, "Ice Maker")
				if LibraryRS("LAUNDRYROOM") then Amenities = addToAry(Amenities, "Laundry Room")
				if LibraryRS("MICROWAVE") then Amenities = addToAry(Amenities, "Microwave Included")
				if LibraryRS("MilitaryDisc") then Amenities = addToAry(Amenities, "Military Discount")
				if LibraryRS("PlayGround2") then Amenities = addToAry(Amenities, "Playground")
				if LibraryRS("Preferred") then Amenities = addToAry(Amenities, "Preferred Employee Program Available")
				if LibraryRS("BUSROUTE") then Amenities = addToAry(Amenities, "Public Transportation Available")
				if LibraryRS("PLAYGROUND") then Amenities = addToAry(Amenities, "Recreational Areas")
				if LibraryRS("SAUNA") then Amenities = addToAry(Amenities, "Sauna")
				if LibraryRS("SCENICOUTLOOK") then Amenities = addToAry(Amenities, "Scenic View")
				if LibraryRS("SPA") then Amenities = addToAry(Amenities, "Spa")
				if LibraryRS("TENNISCOURT") then Amenities = addToAry(Amenities, "Sports Courts")
				if LibraryRS("POOL") then Amenities = addToAry(Amenities, "Swimming Pool(s)")
				if LibraryRS("Trash") then Amenities = addToAry(Amenities, "Trash Pickup Included")
				if LibraryRS("VAULTEDCEILINGS") then Amenities = addToAry(Amenities, "Vaulted Ceilings")
				if LibraryRS("VIP") then Amenities = addToAry(Amenities, "VIP Discount Program")
				if LibraryRS("WASHDRYINC") then Amenities = addToAry(Amenities, "Washer/Dryer Included")
				if LibraryRS("WASHDRYHOOK") then Amenities = addToAry(Amenities, "Washer/Dryer Connections")
				if LibraryRS("WaterSewer") then Amenities = addToAry(Amenities, "Water/Sewer Included")
				if LibraryRS("WIFI") then Amenities = addToAry(Amenities, "WiFi Available")
				if LibraryRS("Age55") then Amenities = addToAry(Amenities, "55 & Older Age Restriction")
				if LibraryRS("Age62") then Amenities = addToAry(Amenities, "62 & Older Age Restriction")
			End If
					
			' Now grab the custom amenities, if any, that were added
			sql = "SELECT * FROM Amenities WHERE AID = " & CLng(AID) & " ORDER BY OrderPosition, RecID"
			set LibraryRS = AptConnRead(sql)
			do while not LibraryRS.eof
				thisamenityLIBRARY = LibraryRS("Amenity")
				Amenities = addToAry(Amenities, thisamenityLIBRARY)
				if LibraryRS("Category") = "Apartment Feature" then
					UnitAmenities = addToAry(UnitAmenities, thisamenityLIBRARY)
				elseif LibraryRS("Category") = "Stimulating Activities" then
					StimActAmenities = addToAry(StimActAmenities, thisamenityLIBRARY)
				elseif LibraryRS("Category") = "Invigorating Activities" then
					InvActAmenities = addToAry(InvActAmenities, thisamenityLIBRARY)
				elseif LibraryRS("Category") = "Inspiring Activities" then
					InsActAmenities = addToAry(InsActAmenities, thisamenityLIBRARY)
				elseif LibraryRS("Category") = "Social Activities" then
					SocActAmenities = addToAry(SocActAmenities, thisamenityLIBRARY)
				else
					CommunityAmenities = addToAry(CommunityAmenities, thisamenityLIBRARY)
				end if 
				LibraryRS.movenext
			loop
			
			GetUnitAmenities()
			GetCommunityAmenities()
	end function
	
	private function GetUnitAmenities()
		set LibraryRS = aptRS
		set LibraryRS2 = amenRS
		If c_MITSAmen=true Then
			if LibraryRS("AC") then UnitAmenities = addToAry(UnitAmenities, "Air Conditioning")
			if LibraryRS("alarm") then UnitAmenities = addToAry(UnitAmenities, "Alarm Included")
			if LibraryRS("AlarmAvail") then UnitAmenities = addToAry(UnitAmenities, "Alarm Available")
			if LibraryRS("BALCONY") then UnitAmenities = addToAry(UnitAmenities, "Balcony/Patio")
			if LibraryRS("CABLE") then UnitAmenities = addToAry(UnitAmenities, "Cable Available")
			if LibraryRS("CABLEINCLUDED") then UnitAmenities = addToAry(UnitAmenities, "Cable Included")
			if LibraryRS("SATALITE") then UnitAmenities = addToAry(UnitAmenities, "Satellite Included")
			if LibraryRS("SatAvailable") then UnitAmenities = addToAry(UnitAmenities, "Satellite Available")
			if LibraryRS("CEILINGFANS") then UnitAmenities = addToAry(UnitAmenities, "Ceiling Fans")
			if LibraryRS("COVEREDPARKING") then UnitAmenities = addToAry(UnitAmenities, "Covered Parking")
			if LibraryRS("DISABILITY") then UnitAmenities = addToAry(UnitAmenities, "Handicap Accessible")
			if LibraryRS("DISHWASHER") then UnitAmenities = addToAry(UnitAmenities, "Dishwasher")
			if LibraryRS("DISPOSAL") then UnitAmenities = addToAry(UnitAmenities, "Disposal")
			if LibraryRS("FIREPLACE") then UnitAmenities = addToAry(UnitAmenities, "Fireplace")
			if LibraryRS("STORAGE") then UnitAmenities = addToAry(UnitAmenities, "Extra Storage")
			if LibraryRS("ISP") then UnitAmenities = addToAry(UnitAmenities, "High Speed Internet Access")
			if LibraryRS("DSL") then UnitAmenities = addToAry(UnitAmenities, "DSL Available")
			if LibraryRS("FURNISHED") then UnitAmenities = addToAry(UnitAmenities, "Furnished Apartment Available")
			if LibraryRS("GARAGES") then UnitAmenities = addToAry(UnitAmenities, "Garages Available")
			if LibraryRS("AttachedGarage") then UnitAmenities = addToAry(UnitAmenities, "Attached Garages Available")
			if LibraryRS("ICEMAKERS") then UnitAmenities = addToAry(UnitAmenities, "Ice Maker")
			if LibraryRS("MICROWAVE") then UnitAmenities = addToAry(UnitAmenities, "Microwave Included")
			if LibraryRS("SCENICOUTLOOK") then UnitAmenities = addToAry(UnitAmenities, "Scenic View")
			if LibraryRS("VAULTEDCEILINGS") then UnitAmenities = addToAry(UnitAmenities, "Vaulted Ceilings")
			if LibraryRS("WASHDRYINC") then UnitAmenities = addToAry(UnitAmenities, "Washer/Dryer Included")
			if LibraryRS("WASHDRYHOOK") then UnitAmenities = addToAry(UnitAmenities, "Washer/Dryer Connections")
			if LibraryRS("WaterSewer") then UnitAmenities = addToAry(UnitAmenities, "Water/Sewer Included")
			if LibraryRS("WIFI") then UnitAmenities = addToAry(UnitAmenities, "WiFi Available")	
			If not amenRs.EOF Then
				if LibraryRS2("Carport") then UnitAmenities = addToAry(UnitAmenities, "Carport")
				if LibraryRS2("Gate") then UnitAmenities = addToAry(UnitAmenities, "Gate")
				if LibraryRS2("Courtyard") then UnitAmenities = addToAry(UnitAmenities, "Courtyard")
				if LibraryRS2("Dryer") then UnitAmenities = addToAry(UnitAmenities, "Dryer")
				if LibraryRS2("Handrails") then UnitAmenities = addToAry(UnitAmenities, "Handrails")
				if LibraryRS2("Heat") then UnitAmenities = addToAry(UnitAmenities, "Heat")
				if LibraryRS2("IndividualClimateControl") then UnitAmenities = addToAry(UnitAmenities, "Individual Climate Control")
				if LibraryRS2("LargeClosets") then UnitAmenities = addToAry(UnitAmenities, "Large Closets")
				if LibraryRS2("Patio") then UnitAmenities = addToAry(UnitAmenities, "Patio")
				if LibraryRS2("PrivateBalcony") then UnitAmenities = addToAry(UnitAmenities, "Private Balcony")
				if LibraryRS2("PrivatePatio") then UnitAmenities = addToAry(UnitAmenities, "Private Patio")
				if LibraryRS2("Range") then UnitAmenities = addToAry(UnitAmenities, "Range")
				if LibraryRS2("Refrigerator") then UnitAmenities = addToAry(UnitAmenities, "Refrigerator")
				if LibraryRS2("Skylight") then UnitAmenities = addToAry(UnitAmenities, "Skylight")
				if LibraryRS2("WindowCoverings") then UnitAmenities = addToAry(UnitAmenities, "Window Coverings")
			End If
		End If
	end function
	
	private function GetCommunityAmenities()
		set LibraryRS = aptRS
		set LibraryRS2 = amenRS
		If c_MITSAmen=true Then
			if LibraryRS("billspaid") then	CommunityAmenities = addToAry(CommunityAmenities, "All Utilities Included")
			if LibraryRS("corphouse") then CommunityAmenities = addToAry(CommunityAmenities, "Corporate Housing Available")
			if LibraryRS("shortlease") then CommunityAmenities = addToAry(CommunityAmenities, "Short Term Leases Available")
			if LibraryRS("BUSINESSLINEACCESS") then CommunityAmenities = addToAry(CommunityAmenities, "Business Center")
			if LibraryRS("CONTROLLEDACCESS") then CommunityAmenities = addToAry(CommunityAmenities, "Controlled Access")
			if LibraryRS("COURTYARD") then CommunityAmenities = addToAry(CommunityAmenities, "Courtyard")
			if LibraryRS("FITNESSCENTER") then CommunityAmenities = addToAry(CommunityAmenities, "Fitness Center")
			if LibraryRS("HIGHRISE") then CommunityAmenities = addToAry(CommunityAmenities, "High Rise")
			if LibraryRS("HouseKeeping") then CommunityAmenities = addToAry(CommunityAmenities, "Complimentary Housekeeping Available")
			if LibraryRS("LAUNDRYROOM") then CommunityAmenities = addToAry(CommunityAmenities, "Laundry Room")
			if LibraryRS("MilitaryDisc") then CommunityAmenities = addToAry(CommunityAmenities, "Military Discount")
			if LibraryRS("PlayGround2") then CommunityAmenities = addToAry(CommunityAmenities, "Playground")
			if LibraryRS("Preferred") then CommunityAmenities = addToAry(CommunityAmenities, "Preferred Employee Program Available")
			if LibraryRS("BUSROUTE") then CommunityAmenities = addToAry(CommunityAmenities, "Public Transportation Available")
			if LibraryRS("PLAYGROUND") then CommunityAmenities = addToAry(CommunityAmenities, "Recreational Areas")
			if LibraryRS("SAUNA") then CommunityAmenities = addToAry(CommunityAmenities, "Sauna")
			if LibraryRS("SPA") then CommunityAmenities = addToAry(CommunityAmenities, "Spa")
			if LibraryRS("TENNISCOURT") then CommunityAmenities = addToAry(CommunityAmenities, "Sports Courts")
			if LibraryRS("POOL") then CommunityAmenities = addToAry(CommunityAmenities, "Swimming Pool(s)")
			if LibraryRS("Trash") then CommunityAmenities = addToAry(CommunityAmenities, "Trash Pickup Included")
			if LibraryRS("VIP") then CommunityAmenities = addToAry(CommunityAmenities, "VIP Discount Program")
			if LibraryRS("Age55") then CommunityAmenities = addToAry(CommunityAmenities, "55 & Older Age Restriction")
			if LibraryRS("Age62") then CommunityAmenities = addToAry(CommunityAmenities, "62 & Older Age Restriction")
			If not amenRs.EOF Then
				if LibraryRS2("Availability24Hours") then CommunityAmenities = addToAry(CommunityAmenities, "24 Hour Availability")
				if LibraryRS2("basketballCourt") then CommunityAmenities = addToAry(CommunityAmenities, "Basketball Court")
				if LibraryRS2("ChildCare") then CommunityAmenities = addToAry(CommunityAmenities, "Child Care")
				if LibraryRS2("ClubDiscount") then CommunityAmenities = addToAry(CommunityAmenities, "Club Discount")
				if LibraryRS2("ClubHouse") then CommunityAmenities = addToAry(CommunityAmenities, "Clubhouse")
				if LibraryRS2("Concierge") then CommunityAmenities = addToAry(CommunityAmenities, "Concierge Service")
				if LibraryRS2("DoorAttendant") then CommunityAmenities = addToAry(CommunityAmenities, "Door Attendant")
				if LibraryRS2("Elevator") then CommunityAmenities = addToAry(CommunityAmenities, "Elevator")
				if LibraryRS2("FreeWeights") then CommunityAmenities = addToAry(CommunityAmenities, "Free Weights")
				if LibraryRS2("GroupExcercise") then CommunityAmenities = addToAry(CommunityAmenities, "GroupExcercise")
				if LibraryRS2("GuestRoom") then CommunityAmenities = addToAry(CommunityAmenities, "Guest Room")
				if LibraryRS2("HouseSitting") then CommunityAmenities = addToAry(CommunityAmenities, "House Sitting")
				if LibraryRS2("Library") then CommunityAmenities = addToAry(CommunityAmenities, "Library")
				if LibraryRS2("MealService") then CommunityAmenities = addToAry(CommunityAmenities, "Meal Service")
				if LibraryRS2("NightPatrol") then CommunityAmenities = addToAry(CommunityAmenities, "Night Patrol")
				if LibraryRS2("OnSiteMaintenance") then CommunityAmenities = addToAry(CommunityAmenities, "On Site Maintenance")
				if LibraryRS2("OnSiteManagement") then CommunityAmenities = addToAry(CommunityAmenities, "On Site Management")
				if LibraryRS2("PackageReceiving") then CommunityAmenities = addToAry(CommunityAmenities, "Package Receiving")
				if LibraryRS2("Racquetball") then CommunityAmenities = addToAry(CommunityAmenities, "Racquetball")
				if LibraryRS2("RecRoom") then CommunityAmenities = addToAry(CommunityAmenities, "Rec Room")
				if LibraryRS2("StorageSpace") then CommunityAmenities = addToAry(CommunityAmenities, "Storage Space")
				if LibraryRS2("Sundeck") then CommunityAmenities = addToAry(CommunityAmenities, "Sundeck")
				if LibraryRS2("Transportation") then CommunityAmenities = addToAry(CommunityAmenities, "Transportation")
				if LibraryRS2("TVLounge") then CommunityAmenities = addToAry(CommunityAmenities, "24 Hour Availability")
				if LibraryRS2("Vintage") then CommunityAmenities = addToAry(CommunityAmenities, "Vintage")
				if LibraryRS2("VolleyballCourt") then CommunityAmenities = addToAry(CommunityAmenities, "Volleyball Court")
			End If
		End If
	end function

	private function GetFloorplans ()
		sql = "SELECT FloorplanID FROM Floorplans WHERE ApartmentID = " & CLng(c_AID) & " ORDER BY " & c_FloorplanOrder
		set LibraryRS = AptConnRead(sql)
		Floorplans = array()
		do while not LibraryRS.eof
			Floorplans = addToAry(Floorplans, LibraryRS("FloorplanID"))
			LibraryRS.movenext
		loop
	end function	
	
	private function GetVTours()
		sql = "SELECT ID FROM vtours WHERE AID = " & CLng(c_AID) & " ORDER BY OrderPosition, Title"
		set LibraryRS = AptConnRead(sql)
		VTours = array()
		do while not LibraryRS.eof
			VTours = addToAry(VTours, LibraryRS("ID"))
			LibraryRS.movenext
		loop
	end function
	
	private function classError(details)
		response.write("<div style=""color: #FF0000;""><b>clsProperty Error: "&details&"</b></div><br>")
		haserrors = true
	end function
end class
%>
