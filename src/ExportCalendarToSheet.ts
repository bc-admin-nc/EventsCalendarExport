// Google ID references (sheets, calendar, etc)
const VenueSheetId = "1y55O_-39m1YBUEWInNKYUFbxrKLehpLKpyklgV_tFqA";
const BandsSheetId = "1Y_c0qHMdTQlBrSzHYWT1cdelnucRU2rTLae3fQW4UHU";
const calendarId = "localmusic@bandsandclubsofthetriangle.com";

const FIRST_DATA_ROW = 2; // start populating sheet on this row
const MATCH_THRESHOLD = 0.9; // used for fuzzy matching
const matchAlgorithm: matchAlgorithmType = "jaroWinkler";

// Read the sheets values into these variables
const venueDataRows = readSheetById(VenueSheetId);
const artistDataRows = readSheetById(BandsSheetId);
const lookupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOOKUP");
lookupSheet?.clearContents();

type matchAlgorithmType = "jaroWinkler" | "levenshteinDistance" | "off";
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

export type VenueDataType = {
	venueName: string,
	venueAddress?: string,
	venueCountry?: string,
	venueCountryAbbreviation?: string,
	venueRegion?: string,
	venueState?: string,
	venueStateAbbreviation?: string,
	venueCity?: string,
	venueNeighborHood?: string,
	venueZip?: string,
	venuePhone?: string,
}

export type ArtistDataType = {
	artistName: string,
	genre?: string
}

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

export const parseLocation = (eventLocation: string): string => {
	// try to parse the venue from the eventLocation
	const regex = /^(.*?);?\s*(.*?),\s*(.*?),\s*([A-Z]{2})\s*(\d{5}),\s*(\w+)$/;
	const match = eventLocation.match(regex);

	// either the parsed value or the entire eventLocation
	return match ? match[2] : eventLocation;
};

export function convertEventToRow(
	calEvent: GoogleAppsScript.Calendar.CalendarEvent,
): rowType {
	const searchVenue = parseLocation(calEvent.getLocation());
	const searchArtitst = calEvent.getTitle();
 	const { venue, found: foundVenue } = matchVenue(searchVenue);
	const { artist, found: foundArtist } = matchArtist(searchArtitst);

	let misMatchData = "";

	if (!foundVenue) {
		Logger.log(`No match found for venue: ${venue.venueName}`);
		lookupSheet?.appendRow(["Venue", venue.venueName])
	}

	if (!foundArtist) {
		Logger.log(`No match found for artist: ${artist.artistName}`);
		lookupSheet?.appendRow(["Artist", artist.artistName])
	}

	if (!foundVenue && !foundArtist) {
		misMatchData += 'Artist & Venue mismatch'
	} else if (!foundVenue && foundArtist) {
		misMatchData += 'Venue mismatch'
	} else if (foundVenue && !foundArtist) {
		misMatchData += 'Artist mismatch'
	}
	
	const row = {
		"Event Title": artist.artistName,
		"Event SEO Title": misMatchData,
		"Event email": "",
		"Event URL": "",
		"Event Venue": venue.venueName,
		"Event Address": venue?.venueAddress,
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
		"Event Country": venue?.venueCountry,
		"Event Country Abbreviation": venue?.venueCountryAbbreviation,
		"Event Region": venue?.venueRegion,
		"Event Region Abbreviation": "",
		"Event State": venue?.venueState,
		"Event State Abbreviation": venue?.venueStateAbbreviation,
		"Event City": venue?.venueCity,
		"Event City Abbreviation": "",
		"Event Neighborhood": venue?.venueNeighborHood,
		"Event Neighborhood Abbreviation": "",
		"Event Postal Code": venue?.venueZip,
		"Event Latitude": "",
		"Event Longitude": "",
		"Event Phone": venue.venuePhone,
		"Event Short Description": "",
		"Event Long Description": calEvent.getDescription(),
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

function writeRowsToSheet(rawData: rowType[]) {
	const data = rawData.map((item) => Object.values(item));

	const sheet: GoogleAppsScript.Spreadsheet.Sheet =
		SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	const numRows = data.length;
	const numCols = data[0].length;
	sheet
		.getRange(FIRST_DATA_ROW, 1, numRows, numCols)
		.setValues(Object.values(data));
}

// for a given date range:
// 1. query a google calendar
// 2. convert event data to row data (attempt to match with known values)
// 3. write data rows to associated Google Sheet
function main(startDateString: string, endDateString: string): rowType[] {
	const startDate: Date = new Date(startDateString);
	const endDate: Date = new Date(endDateString);
	const calendarEvents = getCalenderEvents(startDate, endDate); // query a google calendar
	const dataRows = calendarEvents.map((event) => convertEventToRow(event)); // convert event to data to row data
	writeRowsToSheet(dataRows); // write data rows to associated Google Sheet
	return dataRows;
}

// fetch and match routines from other sheets
function readSheetById(sheetId: string): string[][] {
	const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Sheet1"); // Change sheet name if needed
	const data = sheet?.getDataRange().getValues();
	return data || [];
}

function getVenueData(): VenueDataType[] {
	const venueData = venueDataRows.map((row) => {
		return {
			venueName: row[0],
			venueAddress: row[4],
			venueCountry: row[6],
			venueCountryAbbreviation: row[7],
			venueRegion: row[8],
			venueState: row[10],
			venueStateAbbreviation: row[11],
			venueCity: row[12],
			venueNeighborHood: row[14],
			venueZip: row[16],
			venuePhone: row[19],
		};
	});
	return venueData.sort((a, b) => a.venueName.localeCompare(b.venueName));
}

function getArtistData(): ArtistDataType[] {
	const artistData = artistDataRows.map((row) => {
		return {
			artistName: row[0],
			genre: row[1],
		};
	});
	return artistData.sort((a, b) => a.artistName.localeCompare(b.artistName));
}

function matchVenue(venueName: string): { venue: VenueDataType, found: boolean } {
	const venueData = getVenueData();
	const filtered = venueData.filter((venue) => matchWrapper(venue.venueName, venueName));
	return (filtered.length === 0) ? {venue: {venueName}, found: false} : {venue: filtered[0], found: true}
}

function matchArtist(artistName: string): { artist: ArtistDataType; found: boolean } {
	const artistData = getArtistData();
	const filtered = artistData.filter((artist) => matchWrapper(artist.artistName, artistName));
	return (filtered.length === 0) ? { artist: {artistName}, found: false } : { artist: filtered[0], found: true };
}

function matchWrapper(string1: string, string2: string) {
	if (matchAlgorithm === "jaroWinkler") {
		return jaroWinkler(string1, string2) > MATCH_THRESHOLD;
	}

	if (matchAlgorithm === "levenshteinDistance") {
		return levenshteinDistance(string1, string2) > 1 - MATCH_THRESHOLD;
	}
	return true;
}

// Alogrithm used to fuzzy match 2 strings.  0-1, where higher is a better match
// Higher similarity score (0 to 1) = More similar strings
function jaroWinkler(s1: string, s2: string) {
	let m = 0;
	let i: number;
	let j: number;
	let t = 0;
	const matchDistance = Math.floor(Math.max(s1.length, s2.length) / 2) - 1;
	const s1Matches = new Array(s1.length).fill(false);
	const s2Matches = new Array(s2.length).fill(false);

	for (i = 0; i < s1.length; i++) {
		const start = Math.max(0, i - matchDistance);
		const end = Math.min(i + matchDistance + 1, s2.length);

		for (j = start; j < end; j++) {
			if (s2Matches[j]) continue;
			if (s1[i] !== s2[j]) continue;
			s1Matches[i] = s2Matches[j] = true;
			m++;
			break;
		}
	}

	if (m === 0) return 0;

	let k = 0;
	for (i = 0; i < s1.length; i++) {
		if (!s1Matches[i]) continue;
		while (!s2Matches[k]) k++;
		if (s1[i] !== s2[k]) t++;
		k++;
	}

	t /= 2;

	const jaro = (m / s1.length + m / s2.length + (m - t) / m) / 3;
	let prefixLength = 0;

	for (i = 0; i < Math.min(s1.length, s2.length, 4); i++) {
		if (s1[i] === s2[i]) prefixLength++;
		else break;
	}

	return jaro + prefixLength * 0.1 * (1 - jaro);
}

// Lower distance = More similar strings
function levenshteinDistance(str1: string, str2: string) {
	const len1 = str1.length;
	const len2 = str2.length;
	const matrix = [];

	if (len1 === 0) return len2;
	if (len2 === 0) return len1;

	for (let i = 0; i <= len1; i++) {
		matrix[i] = [i];
	}
	for (let j = 0; j <= len2; j++) {
		matrix[0][j] = j;
	}

	for (let i = 1; i <= len1; i++) {
		for (let j = 1; j <= len2; j++) {
			if (str1[i - 1] === str2[j - 1]) {
				matrix[i][j] = matrix[i - 1][j - 1];
			} else {
				matrix[i][j] = Math.min(
					matrix[i - 1][j] + 1, // Deletion
					matrix[i][j - 1] + 1, // Insertion
					matrix[i - 1][j - 1] + 1, // Substitution
				);
			}
		}
	}

	return matrix[len1][len2];
}

// Example usage
function test_fuzzyMatching() {
	const str1 = "fuzzy matching";
	const str2 = "fuzzi matcing";
	const distance1 = jaroWinkler(str1, str2);
	Logger.log(`jaroWinkler Distance: ${distance1}`);
	const distance2 = levenshteinDistance(str1, str2);
	Logger.log(`Levenshtein Distance: ${distance2}`);
}
