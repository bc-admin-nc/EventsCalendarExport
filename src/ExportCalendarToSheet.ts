const VenueSheetId = "1y55O_-39m1YBUEWInNKYUFbxrKLehpLKpyklgV_tFqA";
const BandsSheetId = "1Y_c0qHMdTQlBrSzHYWT1cdelnucRU2rTLae3fQW4UHU";
const calendarId = "localmusic@bandsandclubsofthetriangle.com";

export type MMDDYYYY = string & { __format: "MM/DD/YYYY" };
export type RECURRING = "S" | "W" | "B" | "M" | "Y";
export type CATEGORY1 =
	| "Live Artists"
	| "Dance"
	| "Jam/Open Mic"
	| "Karaoke"
	| "Bingo/Trivia";
export type CATEGORY2 =
	| "Band/Choir"
	| "DJ/Emcee"
	| "Duo/Trio/Quartet"
	| "Solo"
	| "Symphony/Big Band";
export type CATEGORY3 =
	| "Local Venue"
	| "Indoor Theatre"
	| "Outdoor Concert"
	| "Dance Hall"
	| "Festival";
export type CATEGORY4 =
	| "Alternative"
	| "Bluegrass"
	| "Blues"
	| "Beach/Shag"
	| "Classical"
	| "Folk/Americana"
	| "Gospel"
	| "Hip Hop/Rap"
	| "Jazz"
	| "Latin"
	| "Metal"
	| "Oldies"
	| "Punk"
	| "R&B/Motown"
	| "Rock"
	| "Theatre/Show"
	| "Swing"
	| "Top 40"
	| "Variety"
	| "Reggae"
	| "Country"
	| "Pop"
	| "Funk"
	| "Soul"
	| "Electronic"
	| "World"
	| "Other";
export type rowType = {
	[key: string]: string | number | Date | undefined;
};

interface Event {
	getStartTime(): Date;
	getEndTime(): Date;
	getTitle(): string;
	getLocation(): string;
}

function onOpen() {
	const ui = SpreadsheetApp.getUi();
	ui.createMenu("📌 Get Data") // Menu name
		.addItem("Run Script", "showDateForm") // Option name & function name
		.addToUi();
}

function showDateForm() {
	const html = HtmlService.createHtmlOutputFromFile("src/DateForm")
		.setWidth(300)
		.setHeight(200);
	SpreadsheetApp.getUi().showModalDialog(html, "Query data range");
}

function readSheetById(
	sheetId = "1y55O_-39m1YBUEWInNKYUFbxrKLehpLKpyklgV_tFqA",
): string[][] {
	const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Sheet1"); // Change sheet name if needed
	const data = sheet?.getDataRange().getValues();
	return data || [];
}

function getVenueData() {
	const data = readSheetById(VenueSheetId);
	const venueData = data.map((row) => {
		return {
			venueName: row[0],
			venuAddress: `${row[1]}, ${row[2]}, ${row[3]}, ${row[4]}`,
			venueStreet: row[1],
			venueCity: row[2],
			venueState: row[3],
			venueZip: row[4],
		};
	});
	return venueData.sort((a, b) => a.venueName.localeCompare(b.venueName));
}

function getArtistData() {
	const data = readSheetById(BandsSheetId);
	const artistData = data.map((row) => {
		return {
			artistName: row[0],
			genre: row[1],
		};
	});
	return artistData.sort((a, b) => a.artistName.localeCompare(b.artistName));
}

export function processUserInput(
	startDateString: string,
	endDateString: string,
): void {
	const startDate: Date = new Date(startDateString);
	const endDate: Date = new Date(endDateString);

	const sheet: GoogleAppsScript.Spreadsheet.Sheet =
		SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	sheet.clear(); // Clear old data

  const range = sheet.getRange("A2:A200"); // Dropdown in A2 to 200
  const values = getArtistData().map((artist) => artist.artistName);
  const rule = SpreadsheetApp.newDataValidation()
  .requireValueInList(values)
  .setAllowInvalid(true) // Prevents invalid entries
  .build();
  // range.setDataValidation(rule)

	// Fetch events
	const calendarEvents = CalendarApp.getCalendarById(calendarId).getEvents(
		startDate,
		endDate,
	);
	const rows = calendarEvents.map((event) => convertEventToRow(event));

	// Set headers
	const headers: string[] = [
		"Event Title",
		"Event SEO Title",
		"Event email",
		"Event URL",
		"Event Venue",
		"Event Address",
		"Event Contact Name",
		"Event Start Date",
		"Event End Date",
		"Event Start Time",
		"Event End Time",
		"Event Country",
		"Event Country Abbreviation",
		"Event Region",
		"Event State",
		"Event State Abbreviation",
		"Event City",
		"Event City Abbreviation",
		"Event Neighborhood",
		"Event Neighborhood Abbreviation",
		"Event Postal Code",
		"Event Latitude",
		"Event Longitude",
		"Event Phone",
		"Event Short Description",
		"Event Long Description",
		"Event SEO Description",
		"Event Keywords",
		"Event Renewal Date",
		"Event Status",
		"Event Category 1",
		"Event Category 2",
		"Event Category 3",
		"Event Category 4",
		"Event Category 5",
		"Event DB ID",
		"Custom ID",
		"Account Username",
		"Account Password",
		"Account Contact First Name",
		"Account Contact Last Name",
		"Account Contact Company",
		"Account Contact Address",
		"Account Contact Address2",
		"Account Contact Country",
		"Account Contact State",
		"Account Contact City",
		"Account Contact Postal Code",
		"Account Contact Phone",
		"Account Contact Email",
		"Account Contact URL",
		"Gallery Main Image",
		"Gallery Image 1",
		"Gallery Image 2",
		"Gallery Image 3",
		"Gallery Image 4",
		"Gallery Image 5",
		"Gallery Image 6",
		"Gallery Image 7",
		"Gallery Image 8",
		"Gallery Image 9",
	];
	sheet.appendRow(headers);
	rows.map((row) => sheet.appendRow(Object.values(row)));

	Logger.log("Calendar events exported successfully!");
}

export function convertEventToRow(
	calEvent: GoogleAppsScript.Calendar.CalendarEvent,
): rowType {
	const row = {
		"Event Title": calEvent.getTitle(),
		"Event SEO Title": "",
		"Event email": "",
		"Event URL": "",
		"Event Venue": calEvent.getLocation(),
		"Event Address": "",
		"Event Contact Name": "",
		"Event Start Date": Utilities.formatDate(
			calEvent.getStartTime(),
			Session.getScriptTimeZone(),
			"MM/dd/yyyy",
		),
		"Event End Date": Utilities.formatDate(
			calEvent.getEndTime(),
			Session.getScriptTimeZone(),
			"MM/dd/yyyy",
		),
		"Event Start Time": calEvent.getStartTime().toLocaleTimeString(),
		"Event End Time": calEvent.getEndTime().toLocaleTimeString(),
		"Event Country": "United States",
		"Event Country Abbreviation": "USA",
		"Event Region": "",
		"Event State": "",
		"Event State Abbreviation": "",
		"Event City": "",
		"Event City Abbreviation": "",
		"Event Neighborhood": "",
		"Event Neighborhood Abbreviation": "",
		"Event Postal Code": "",
		"Event Latitude": "",
		"Event Longitude": "",
		"Event Phone": "",
		"Event Short Description": "",
		"Event Long Description": "",
		"Event SEO Description": "",
		"Event Keywords": "",
		"Event Renewal Date": "",
		"Event Status": "",
		"Event Category 1": "",
		"Event Category 2": "",
		"Event Category 3": "",
		"Event Category 4": "",
		"Event Category 5": "",
		"Event DB ID": "",
		"Custom ID": "",
		"Account Username": "",
		"Account Password": "",
		"Account Contact First Name": "",
		"Account Contact Last Name": "",
		"Account Contact Company": "",
		"Account Contact Address": "",
		"Account Contact Address2": "",
		"Account Contact Country": "",
		"Account Contact State": "",
		"Account Contact City": "",
		"Account Contact Postal Code": "",
		"Account Contact Phone": "",
		"Account Contact Email": "",
		"Account Contact URL": "",
		"Gallery Main Image": "",
		"Gallery Image 1": "",
		"Gallery Image 2": "",
		"Gallery Image 3": "",
		"Gallery Image 4": "",
		"Gallery Image 5": "",
		"Gallery Image 6": "",
		"Gallery Image 7": "",
		"Gallery Image 8": "",
		"Gallery Image 9": "",
	};
	return row;
}
