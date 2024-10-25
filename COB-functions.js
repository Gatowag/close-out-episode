// ░░░░░░░░░▓ FUNCTION THAT GETS CALLED FROM THE MENU ITEM
function closeOutButton() {
	console.log(`DIAG: ${new Date().toISOString().slice(14,23)}, starting COB`);
	
	const foundRow = findRow();
	const epLabel = tab2.getRange(tab2TitleCol + foundRow).getValue();
	const adLabel = adInput(foundRow);

	rowDiag = foundRow;
	adDiag = adLabel[0];
	isEpMissing(epLabel);
	getYoutubeData();
	getSheetDataCoB(newRow - 70);
	findMatches(epLabel, adLabel);
	updateMusicDoc();
}

// ░░░░░░░░░▓ FIND THE LOWEST WHITE BACKGROUND ROW (UNPUBLISHED) IN TAB 2
function findRow(){

	let rowCand2D = tab2.getRange(`${tab2TitleCol}1:${tab2TitleCol}40`).getBackgrounds();	// returns a 2D array of all background values in first 40 rows
	let rowCandidates = [].concat.apply([], rowCand2D);			// converts rowCand2D from a 2D array to a 1D array
	let foundRow = rowCandidates.lastIndexOf("#ffffff");		// returns last row of white background (counting from 0)
	
	return foundRow + 1;
}

// ░░░░░░░░░▓ FILTERS THE SPONSOR DATA TO REMOVE EVERYTHING AFTER " ("
function adInput(row){
	
	if (tab2.getRange(tab2SponsCol + row).getValue().match(/.*?(?= \()/) == null){
		return tab2.getRange(tab2SponsCol + row).getValue().match(/.*/);
	} else {
		return tab2.getRange(tab2SponsCol + row).getValue().match(/.*?(?= \()/);
	};
}

// ░░░░░░░░░▓ IF NO TEXT IS FOUND FOR EPISODE TITLE, CUSTOM ERROR MESSAGE IS PASSED TO epDiag
function isEpMissing(ep){
	
	if (ep == ""){ epDiag = "not found" }
	else { epDiag = ep }
}

// ░░░░░░░░░▓ PULLS YOUTUBE DATA, STORES IT IN GLOBAL VARIABLE "UPLOADS"
function getYoutubeData(){
	
	// get uploads playlist ID from given channel
	const mrChannel = YouTube.Channels.list('contentDetails', { id: "UC42VsoDtra5hMiXZSsD6eGg" });
	const mrUploadsID = mrChannel.items[0].contentDetails.relatedPlaylists.uploads;
	// get data from 10 most recent uploads
	const mrUploads = YouTube.PlaylistItems.list('snippet', {
		playlistId: mrUploadsID,
		maxResults: 50,
	});

	let IDlist = [];
	// loop through latest videos to store their IDs (end of url) to IDlist
	for (let f = 0; f < 50; f++) {
	  let tempID = mrUploads.items[f].snippet.resourceId.videoId;
		IDlist.push(tempID);
	};

	// gets data to filter out YT shorts
	const mrVideos = YouTube.Videos.list(
		'contentDetails',
		{ id: [IDlist.toString()] } );

	// loop through latest videos to store their IDs (end of url) to IDlist
	for (let s = 0; s < 50; s++) {
		let sD = mrVideos.items[s].contentDetails.duration;
		let sM = sD.indexOf("M"), sT = sD.indexOf("T"), sH = sD.indexOf("H");

		if (sH != -1) { recentFull.push(s) }
		else if (sM == -1) { continue; }
		else if (Number(sD.slice(sT + 1, sM)) >= 2) { recentFull.push(s) };
	};

	uploads = mrUploads;
}

// ░░░░░░░░░▓ READS DATA FROM THE SPREADSHEET WHEN THE SIDEBAR LOADS,
// ░░░░░░░░░▓ ALL RELEVANT DATA GETS PASSED TO THEIR RESPECTIVE ARRAYS IN OBJECT "dataArray"
function getSheetDataCoB(offset){
	const labelRange = tab1.getRange(offset,12,newRow - (offset-1),1);

		// cycle through all rows of selected data
		for ( i = 0; i < ((newRow + 1) - offset); i++){
			// if ad entry is detected
			if (labelRange.getBackgrounds()[i] == "#90eba6") {
				// store the ad title
				dataArray.recSponsors.push(labelRange.getValues()[i]);
			};
		};
	
	// strip "ad: " text from all ad titles and store them separately
	dataArray.allSponsors = labelRange.getValues().filter(value => /^ad: /i.test(value));

	return dataArray;
}

// ░░░░░░░░░▓ COMPARE EP AND AD LABELS TO DATA FROM THE PRODUCTION TAB
function findMatches(ep, ad){
	
	const lastRow = tab1.getLastRow();
	const labelRange = tab1.getRange((lastRow - 39), 12, 40, 1);
	var matchedEpRows = [];
	var matchedAdRows = [];
	
		// SENDS THE ROW NUMBER OF MATCHING *AD* ENTRIES FROM TAB 1 TO matchedAdRows ARRAY
	if (ad != "") {
		suggestMatch();
	};
	
	for ( i = 0; i < 40; i++) {
		let rollingCell = labelRange.getValues()[i].toString();
		// SENDS THE ROW NUMBER OF MATCHING *EPISODE* ENTRIES FROM TAB 1 TO matchedEpRows ARRAY
		if (rollingCell == ep)
			{ matchedEpRows.push(lastRow - (39 - Number([i]))) };
		if (rollingCell == "ad: " + bestAdMatch && labelRange.getBackgrounds()[i] == "#90eba6")
			{ matchedAdRows.push(Number(lastRow - (39 - Number([i])))) };
	};

	errorCheck(matchedAdRows, matchedEpRows, ep);
}

// ░░░░░░░░░▓ SUGGESTS THE NEAREST CANDIDATE TO CORRECT A TYPO
function suggestMatch(){
	
	let inputString = adDiag.toLocaleLowerCase();									// user input set to all lower case
	let candidates = [].concat.apply([], dataArray.recSponsors);					// a flattened array of recorded sponsors
	let candConfidences = [];														// a list of each candidate's confidence rating
					
	candidates.forEach(function(cand) {												// run through each ad candidate
		let clone = cand.slice(4).toLocaleLowerCase();								// removes "ad: " from the beginning of the candidate
		let candLength = clone.length;												// length of candidate string before filtering
		let penalties = 0;															// how many letters are input but not matched
		let score = candLength;
		let candReward = 0;															// cumulative consecutive matches
		let consMatches = -1;														// variable consecutive match counter
		let inputLength = inputString.length;										// length of user input as a number
							
		for ( i = 0; i < inputLength; i++) {										// runs through each letter of the user input
			let indx = clone.indexOf(inputString[i]);								// define the location of this character in the cand, -1 if none
			if (indx !== -1) {														// if the letter exists in the candidate, then...
				let str1 = clone.slice(0, indx);									// cut everything before the letter
				let str2 = clone.slice(indx + 1);									// cut everything after the letter
				clone = str1 += str2;												// combine the strings to remove the letter
				let proxPoint = roundDec(1 - Math.abs((indx / candLength) - (i / inputLength)),2);
				score = score - proxPoint;
				consMatches++;	
					
			} else {															// but if it can't be found in the candidate...
				penalties++;														// ... then it increases total length
				if (consMatches > 0){ candReward = candReward + consMatches };		// send consecutive bonus if it's built up
				consMatches = -1;													// reset consecutive bonus	
			};
		};
		if (consMatches > 0) { candReward = candReward + consMatches };				// settles any bonuses from the last loop

		let reward = score * (candReward / candLength);

		// adds mismatches from the user input to the total length
		score = roundDec(score - reward + penalties, 2);

		let confidence = Math.round((2 - (score / (candLength))) * 50);

		candConfidences.push(confidence);											// send filtered candidate string to list
	});
	
	let bestMatch = Math.max(...candConfidences);									// find the candidate with the highest confidence
	let bestMatchName = candidates[candConfidences.indexOf(bestMatch)].slice(4);	// return the corresponding title without "ad: "
	bestAdMatch = bestMatchName;													// global version of the above
	adCorrectiveConf = candConfidences[candConfidences.indexOf(bestMatch)];			// sets adCorrectiveConf to the confidence rating
}

// ░░░░░░░░░▓ IF A MATCH ISN'T FOUND, RUN AN ERROR MESSAGE AND END PROCESS
function errorCheck(ads, eps, epLabel){
	if (eps[0] == null){
		errorMessage = "Your episode did not close successfully.";
		errorFailure = epDiag;
		
		const htmlForModal = HtmlService.createTemplateFromFile("COB-ep-error");
		const htmlOutput = htmlForModal.evaluate();
			htmlOutput.setWidth(410);
			htmlOutput.setHeight(130);
			SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Episode Error");
	} else {
		if (adCorrectiveConf != 100 && adDiag != "") {				// if the user input ad anchor has text but isn't a perfect match
			commitClose(eps, ads);
			const htmlForModal = HtmlService.createTemplateFromFile("COB-ad-error");
			const htmlOutput = htmlForModal.evaluate();
				htmlOutput.setWidth(314);
				htmlOutput.setHeight(250);
			SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Corrective Assumption");
		} else {
			commitClose(eps, ads);
		}
	}
}

// ░░░░░░░░░▓ RUN ALL OF THE FUNCTIONS THAT ACTUALLY CHANGE CELLS
function commitClose(eps, ads){
	closeOutEp(eps);
	closeOutAd(eps, ads);
	submitDataTab2(rowDiag);
}

// ░░░░░░░░░▓ EACH MATCHED EP ENTRY | SUBMITS DATA, FORMATS DATA, COLORS ROW, HIDES ROW
function closeOutEp(eps){
	
	eps.forEach(function(rowNum){
		if (tab1.getRange(tab1PartCol + rowNum).getValue() > 1){
			tab1.getRange(rowNum + ":" + rowNum).setBackground("#b7b7b7");
		} else {
			tab1.getRange(rowNum + ":" + rowNum).setBackground("#d9d9d9"); };
		tab1.getRange(tab1PubCol + rowNum).setValue(pubNum);
		tab1.getRange(tab1PubDateCol + rowNum).setValue(new Date(uploads.items[recentFull[0]].snippet.publishedAt))
			.setNumberFormat("yyyy-mm-dd");
		tab1.getRange(tab1LinkCol + rowNum).setValue("https://youtu.be/" + uploads.items[recentFull[0]].snippet.resourceId.videoId)
			.setHorizontalAlignment("right");
		prodNum = tab1.getRange(tab1ProdCol + rowNum).getValue();
		tab1.hideRows(rowNum);
	});
	
}

// ░░░░░░░░░▓ OLDEST MATCHED AD ENTRY | SUBMITS DATA, FORMATS DATA, COLORS ROW, HIDES ROW
function closeOutAd(eps, ads){
	if (ads != ""){
		let firstAdRow;
		
		firstAdRow = Math.min(...ads);
		let firstAdLabel = tab1.getRange(tab1TitleCol + firstAdRow).getValue().slice(4);
		
		tab1.getRange(firstAdRow + ":" + firstAdRow).setBackground("#b7b7b7");
		tab1.getRange(tab1PubCol + firstAdRow).setValue(pubNum);
		tab1.getRange(tab1PubDateCol + firstAdRow).setValue(new Date(uploads.items[recentFull[0]].snippet.publishedAt))
			.setNumberFormat("yyyy-mm-dd");
		tab1.getRange(tab1LinkCol + firstAdRow).setValue("https://youtu.be/" + uploads.items[recentFull[0]].snippet.resourceId.videoId)
			.setHorizontalAlignment("right");
		eps.forEach(function(rowNum){
			tab1.getRange(tab1SponsCol + rowNum).setValue(firstAdLabel);
		});
		tab1.hideRows(firstAdRow);
	}
}

// ░░░░░░░░░▓ SUBMITS AND MODIFIES DATA IN TAB 2
function submitDataTab2(row){
	
	tab2.getRange(tab2ProdCol + row + ":" + tab2LinkCol + row)
		.setFontFamily("Roboto Mono")
		.setFontSize(8);

	// fills out release number
	tab2.getRange(tab2ProdCol + row)
		.setValue(`p${prodNum}`)
	tab2.getRange(tab2PubCol + row)
		.setValue(determinePubNum(row))
		.setFontSize(10);

	// sets the air date to the date the selected video was published
	tab2.getRange(tab2DateCol + row)
		.setValue(new Date(uploads.items[recentFull[0]].snippet.publishedAt))
		.setNumberFormat("yyyy-mm-dd");
	
	// submits the corresponding youtube link
	tab2.getRange(tab2LinkCol + row)
		.setValue("https://youtu.be/" + uploads.items[recentFull[0]].snippet.resourceId.videoId)
		.setHorizontalAlignment("right");
	
	// changes the production label to the published episode title
	tab2.getRange(tab2TitleCol + row)
		.setValue(uploads.items[recentFull[0]].snippet.title)
		.setFontFamily("Roboto")
		.setFontSize(10);

	// unbolds the sponsor
	tab2.getRange(tab2SponsCol + row)
		.setFontFamily("Roboto")
		.setFontSize(10)
		.setFontWeight("normal");
	
	// sets background color and vertical center for the entire row
	tab2.getRange(row + ":" + row)
		.setVerticalAlignment("middle")
		.setBackground("#d9d9d9");
}

// ░░░░░░░░░▓ DETERMINES PUBLISH NUMBER
function determinePubNum(row){
	
	const rangePubNums = tab2.getRange(row, 2, 10, 1);
	var recentPubNums = [];

	for ( i = 0; i < 10; i++ ){
		if (rangePubNums.getValues()[i] != "" && rangePubNums.getValues()[i] != "-")
		{ recentPubNums.push(Number(rangePubNums.getValues()[i])); };
	};

	pubNum = Math.max(...recentPubNums) + 1;
	return pubNum;
}

// ░░░░░░░░░▓ WRITES DATA FROM CLOSED-OUT EPISODE TO MUSIC DOC
function updateMusicDoc() {
	const mDoc = SpreadsheetApp.openById(`PASTE ID FROM URL HERE`);
	const mTab = mDoc.getSheets()[0];
	const mVals = mTab.getRange("A2:C25").getValues();
	let mRow, mRow2;
	
	for (m = 23; m >= 0; m--) {
		if (Date.parse(mVals[m][2]) != null) { mRow2 = m + 2; };
		if (mVals[m][0] == "") {
			mRow = m + 2;
			break;
		}
	};
	
	mTab.getRange(`A${mRow}`)
		.setValue(pubNum)
    	.setFontFamily("Roboto Mono")
		.setFontSize(12)
		.setFontWeight("bold")
		.setHorizontalAlignment("center");

	// fills in published episode title
  	mTab.getRange(`B${mRow}`)
		.setValue(uploads.items[recentFull[0]].snippet.title)
		.setHorizontalAlignment("left")
		.setFontSize(14);

	// sets the air date to the date the selected video was published
  	mTab.getRange(`C${mRow}`)
		.setValue(new Date(uploads.items[recentFull[0]].snippet.publishedAt))
		.setNumberFormat("yyyy-mm-dd")
		.setHorizontalAlignment("center")
		.setFontSize(10);

	// submits the corresponding youtube link
	mTab.getRange(`D${mRow}`)
		.setValue("https://youtu.be/" + uploads.items[recentFull[0]].snippet.resourceId.videoId)
		.setHorizontalAlignment("left")
		.setFontSize(10);

	// format entire row after production number
	mTab.getRange(`B${mRow}:2`)
		.setFontFamily("Roboto");
	
	// format entire episode's entry
	mTab.getRange(`A${mRow}:D${mRow2}`)
		.setVerticalAlignment("middle");
}