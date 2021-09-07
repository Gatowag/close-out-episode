// ░░░░░░░░░▓ GLOBAL VARIABLES
let uploads;


// ░░░░░░░░░▓ FUNCTION THAT GETS CALLED FROM THE MENU ITEM
function closeOutEpisode() {
	
	const cell = ss.getActiveSheet().getCurrentCell();
	const epLabel = cell.getValue();
	const thisRow = cell.getRow();
	
	tab2.getRange("B" + thisRow + ":O" + thisRow).setBackground("#d9d9d9");
	getYoutubeData();
	submitDataTab2(thisRow);
	submitDataTab1(epLabel);
	// formatDataTab1();
}


// ░░░░░░░░░▓ SUBMITS AND MODIFIES DATA IN TAB 2
function submitDataTab2(row){
	
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

function submitDataTab1(label){
	
	let lastRow = tab1.getLastRow();
	const labelRange = tab1.getRange((lastRow - 40), 12, 40, 1);
	var matchingRows = [];
		
	for ( i = 0; i < 40; i++){
		if (labelRange.getValues()[i].toString() == label) {
			matchingRows.push(lastRow - (40 - Number([i])));
		};
	};
	
	matchingRows.forEach(function(el){
		tab1.getRange(el + ":" + el).setBackground("#d9d9d9");
		});
}


// ░░░░░░░░░▓ DETERMINES PUBLISH NUMBER
function determinePubNum(row){
	
	const rangePubNums = tab2.getRange(row, 3, 10, 1);
	var recentPubNums = [];

	for ( i = 0; i < 10; i++ ){
		if (rangePubNums.getValues()[i] != "" && rangePubNums.getValues()[i] != "-")
		{ recentPubNums.push(Number(rangePubNums.getValues()[i])); };
	};

	return Math.max(...recentPubNums) + 1;
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
