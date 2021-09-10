// ░░░░░░░░░▓ GLOBAL VARIABLES
let uploads;
let pubNum;
let adDiag;
let epDiag;
let rowDiag;
let errorMessage;
let errorFailure;


// ░░░░░░░░░▓ FUNCTION THAT GETS CALLED FROM THE MENU ITEM
function closeOutButton() {
	
	const foundRow = findRow();
	const epLabel = tab2.getRange("G" + foundRow).getValue();
	const adLabel = adInput(foundRow);

	rowDiag = foundRow;
	adDiag = adLabel[0];
	if (epLabel == ""){ epDiag = "Submitting nothing" }
	else { epDiag = epLabel;}

	getYoutubeData();
	findMatches(epLabel, adLabel);
}

// ░░░░░░░░░▓ FIND THE LOWEST WHITE BACKGROUND ROW (UNPUBLISHED) IN TAB 2
function findRow(){

	let rowCand2D = tab2.getRange("G1:G40").getBackgrounds();	// returns a 2D array of all background values in first 40 rows
	let rowCandidates = [].concat.apply([], rowCand2D);			// converts rowCand2D from a 2D array to a 1D array
	let foundRow = rowCandidates.lastIndexOf("#ffffff");		// returns last row of white background (counting from 0)
	
	return foundRow + 1;
}

// ░░░░░░░░░▓ FILTERS THE SPONSOR DATA TO REMOVE EVERYTHING AFTER " ("
function adInput(row){
	if (tab2.getRange("K" + row).getValue().match(/.*?(?= \()/) == null){
		return tab2.getRange("K" + row).getValue().match(/.*/);
	} else {
		return tab2.getRange("K" + row).getValue().match(/.*?(?= \()/);
	};
}

// ░░░░░░░░░▓ PULLS YOUTUBE DATA, STORES IT IN GLOBAL VARIABLE "UPLOADS"
function getYoutubeData(){
	
	let results = YouTube.Channels.list('contentDetails', {
		id: "UC42VsoDtra5hMiXZSsD6eGg"
	});

	for (var i = 0; i < results.items.length; i++) {
		var item = results.items[i];
		var playlistId = item.contentDetails.relatedPlaylists.uploads;
		var playlistResponse = YouTube.PlaylistItems.list('snippet', {
			playlistId: playlistId,
			maxResults: 5,
		});
	}

	uploads = playlistResponse;
}

// ░░░░░░░░░▓ COMPARE EP AND AD LABELS TO DATA FROM THE PRODUCTION TAB
function findMatches(ep, ad){
	
	const lastRow = tab1.getLastRow();
	const labelRange = tab1.getRange((lastRow - 39), 12, 40, 1);
	var matchedEpRows = [];
	var matchedAdRows = [];
	
	for ( i = 0; i < 40; i++) {
		// SENDS THE ROW NUMBER OF MATCHING *EPISODE* ENTRIES FROM TAB 1 TO matchedEpRows ARRAY
		if (labelRange.getValues()[i].toString() == ep) {
			matchedEpRows.push(lastRow - (39 - Number([i]))); };
		
		// SENDS THE ROW NUMBER OF MATCHING *AD* ENTRIES FROM TAB 1 TO matchedAdRows ARRAY
		if (ad != null) {
			if (new RegExp(ad[0],'i').test(labelRange.getValues()[i]) == true && labelRange.getBackgrounds()[i] == "#90eba6") {
				matchedAdRows.push(Number(lastRow - (39 - Number([i])))) };
		};
	};

	errorCheck(matchedEpRows, matchedAdRows);
}

// ░░░░░░░░░▓ IF A MATCH ISN'T FOUND, RUN AN ERROR MESSAGE AND END PROCESS
function errorCheck(eps, ads){
	
	if (eps[0] == null){
		errorMessage = "Your episode did not close successfully.";
		errorFailure = epDiag;
		
		const htmlForModal = HtmlService.createTemplateFromFile("COB-error");
		const htmlOutput = htmlForModal.evaluate();
			htmlOutput.setWidth(430);
			htmlOutput.setHeight(110);
		const ui = SpreadsheetApp.getUi();
			ui.showModalDialog(htmlOutput, "Episode Error");
	} else {		
		if (ads[0] == null){
			errorMessage = "Your episode closed successfully, but your ad didn't.";
			errorFailure = adDiag;
			
			const htmlForModal = HtmlService.createTemplateFromFile("COB-error");
			const htmlOutput = htmlForModal.evaluate();
				htmlOutput.setWidth(430);
				htmlOutput.setHeight(90);
			const ui = SpreadsheetApp.getUi();
				ui.showModalDialog(htmlOutput, "Sponsor Error");
		} else {
			submitDataTab2(rowDiag);
			closeOutEp(eps, ads);
			closeOutAd(ads);
		}
	}
}

// ░░░░░░░░░▓ EACH MATCHED EP ENTRY | SUBMITS DATA, FORMATS DATA, COLORS ROW, HIDES ROW
function closeOutEp(eps, ads){
	eps.forEach(function(rowNum){
		if (tab1.getRange("M" + rowNum).getValue() > 1){
			tab1.getRange(rowNum + ":" + rowNum).setBackground("#b7b7b7");
		} else {
			tab1.getRange(rowNum + ":" + rowNum).setBackground("#d9d9d9"); };

		tab1.getRange("O" + rowNum).setValue(pubNum);
		tab1.getRange("P" + rowNum).setValue(new Date(uploads.items[0].snippet.publishedAt))
			.setNumberFormat("yyyy-mm-dd");
		tab1.getRange("Q" + rowNum).setValue("https://youtu.be/" + uploads.items[0].snippet.resourceId.videoId)
			.setHorizontalAlignment("right");
		tab1.getRange("R" + rowNum).setValue(ads[0]);
		tab1.hideRows(rowNum);
	});
}

// ░░░░░░░░░▓ OLDEST MATCHED AD ENTRY | SUBMITS DATA, FORMATS DATA, COLORS ROW, HIDES ROW
function closeOutAd(ads){
	if (ads != null){
		let oldestAd;
		
		oldestAd = Math.min(...ads);
		
		tab1.getRange(oldestAd + ":" + oldestAd).setBackground("#b7b7b7");
		tab1.getRange("O" + oldestAd).setValue(pubNum);
		tab1.getRange("P" + oldestAd).setValue(new Date(uploads.items[0].snippet.publishedAt))
			.setNumberFormat("yyyy-mm-dd");
		tab1.getRange("Q" + oldestAd).setValue("https://youtu.be/" + uploads.items[0].snippet.resourceId.videoId)
			.setHorizontalAlignment("right");
		tab1.hideRows(oldestAd);
	}
}

// ░░░░░░░░░▓ SUBMITS AND MODIFIES DATA IN TAB 2
function submitDataTab2(row){
	
	// sets background color for the entire row
	tab2.getRange("B" + row + ":O" + row).setBackground("#d9d9d9");
	
	// fills out release number
	tab2.getRange("B" + row)
		.setValue("MR");
	tab2.getRange("C" + row)
		.setValue(determinePubNum(row));

	// sets the air date to the date the selected video was published
	tab2.getRange("D" + row).setValue(new Date(uploads.items[0].snippet.publishedAt))
		.setNumberFormat("yyyy-mm-dd");
	
	// submits the corresponding youtube link
	tab2.getRange("E" + row).setValue("https://youtu.be/" + uploads.items[0].snippet.resourceId.videoId)
		.setFontSize(8)
		.setHorizontalAlignment("right");
	
	// changes the production label to the published episode title
	tab2.getRange("G" + row).setValue(uploads.items[0].snippet.title);

	// unbolds the sponsor
	tab2.getRange("K" + row)
		.setFontWeight("normal");
	
	// sets entire row to vertical center
	tab2.getRange(row + ":" + row)
		.setVerticalAlignment("middle");
}

// ░░░░░░░░░▓ DETERMINES PUBLISH NUMBER
function determinePubNum(row){
	
	const rangePubNums = tab2.getRange(row, 3, 10, 1);
	var recentPubNums = [];

	for ( i = 0; i < 10; i++ ){
		if (rangePubNums.getValues()[i] != "" && rangePubNums.getValues()[i] != "-")
		{ recentPubNums.push(Number(rangePubNums.getValues()[i])); };
	};

	pubNum = Math.max(...recentPubNums) + 1;
	return pubNum;
}