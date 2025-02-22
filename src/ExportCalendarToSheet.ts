const VenueSheetId = "1y55O_-39m1YBUEWInNKYUFbxrKLehpLKpyklgV_tFqA";
const BandsSheetId = "1Y_c0qHMdTQlBrSzHYWT1cdelnucRU2rTLae3fQW4UHU";
const calendarId = "localmusic@bandsandclubsofthetriangle.com";

const FIRST_DATA_ROW = 2;	// start populating sheet on this row

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

function test_main() {
	const [startDate, endDate] = getTestDates();
	const rows = main(startDate.toDateString(), endDate.toDateString());
	Logger.log(`${rows.length} rows created successfully!`);
	return rows;
}

function test_getCalendarEvents() {
	const [startDate, endDate] = getTestDates();
	const events = getCalenderEvents(startDate, endDate);
	Logger.log(`${events.length} calendar events exported successfully!`);
	return events;
}

const test_parseLocation = () => {
	const testString = "";
	const { venue, street, city, state, zip } = parseLocation(testString);

	Logger.log(`Venue: ${venue}`);
	Logger.log(`Street: ${street}`);
	Logger.log(`City: ${city}`);
	Logger.log(`State: ${state}`);
	Logger.log(`Zip: ${zip}`);
};

export function getTestDates() {
	const startDate = new Date();
	const endDate = new Date();
	endDate.setDate(endDate.getDate() + 1);
	return [startDate, endDate];
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

function getCalenderEvents(
	startDate: Date,
	endDate: Date,
): GoogleAppsScript.Calendar.CalendarEvent[] {
	const calendarEvents = CalendarApp.getCalendarById(calendarId).getEvents(
		startDate,
		endDate,
	);
	return calendarEvents;
}

export const parseLocation = (eventLocation: string) => {
	let venue = "VENUE";
	let street = "STREET";
	let city = "CITY";
	let state = "STATE";
	let zip = "ZIP";
	const country = "United States";
	const countryAbbreviation = "USA";
	const parsed = eventLocation.split(",");

	const regex = /^(.*?);?\s*(.*?),\s*(.*?),\s*([A-Z]{2})\s*(\d{5}),\s*(\w+)$/;

	const match = eventLocation.match(regex);

	if (match) {
		venue = match[1] || "N/A";
		street = match[2];
		city = match[3];
		state = match[4];
		zip = match[5];
	}

	return {
		venue,
		street,
		city,
		state,
		zip,
		country,
		countryAbbreviation,
	};
};

export function convertEventToRow(
	calEvent: GoogleAppsScript.Calendar.CalendarEvent,
): rowType {
	const { venue, street, city, state, zip, country, countryAbbreviation } =
		parseLocation(calEvent.getLocation());

	const row = {
		"Event Title": calEvent.getTitle(),
		"Event SEO Title": calEvent.getTitle(),
		"Event email": "",
		"Event URL": "",
		"Event Venue": venue,
		"Event Address": `${street}, ${city}, ${state}, ${zip}`,
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
		"Event Country": country,
		"Event Country Abbreviation": countryAbbreviation,
		"": "",
		"Event Region": "",
		"Event Region Abbreviation": "",
		"Event State": state,
		"Event State Abbreviation": "",
		"Event City": city,
		"Event City Abbreviation": "",
		"Event Neighborhood": "",
		"Event Neighborhood Abbreviation": "",
		"Event Postal Code": zip,
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

function main(
	startDateString: string,
	endDateString: string,
): rowType[] {
	const startDate: Date = new Date(startDateString);
	const endDate: Date = new Date(endDateString);

	const sheet: GoogleAppsScript.Spreadsheet.Sheet =
		SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	const lastRow = sheet.getLastRow();
	sheet
		.getRange(FIRST_DATA_ROW, 1, lastRow - 3, sheet.getLastColumn())
		.clearContent();

	// Fetch events
	const calendarEvents = getCalenderEvents(startDate, endDate);
	const rows = calendarEvents.map((event) => convertEventToRow(event));
	rows.map((row) => sheet.appendRow(Object.values(row)));
	return rows;
}



// function readSheetById(
// 	sheetId = "1y55O_-39m1YBUEWInNKYUFbxrKLehpLKpyklgV_tFqA",
// ): string[][] {
// 	const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Sheet1"); // Change sheet name if needed
// 	const data = sheet?.getDataRange().getValues();
// 	return data || [];
// }

// function getVenueData() {
// 	const data = readSheetById(VenueSheetId);
// 	const venueData = data.map((row) => {
// 		return {
// 			venueName: row[0],
// 			venuAddress: `${row[1]}, ${row[2]}, ${row[3]}, ${row[4]}`,
// 			venueStreet: row[1],
// 			venueCity: row[2],
// 			venueState: row[3],
// 			venueZip: row[4],
// 		};
// 	});
// 	return venueData.sort((a, b) => a.venueName.localeCompare(b.venueName));
// }

// function getArtistData() {
// 	const data = readSheetById(BandsSheetId);
// 	const artistData = data.map((row) => {
// 		return {
// 			artistName: row[0],
// 			genre: row[1],
// 		};
// 	});
// 	return artistData.sort((a, b) => a.artistName.localeCompare(b.artistName));
// }
