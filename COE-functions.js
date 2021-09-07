let uploads;

function closeOutEpisode() {
	
	let cell = ss.getActiveSheet().getCurrentCell();
	let epLabel = cell.value;
	let thisRow = cell.getRow();
	
	tab2.getRange("B" + thisRow + ":O" + thisRow).setBackground("#d9d9d9");
	getYoutubeData();
	submitDataTab2(cell, thisRow);
}

function submitDataTab2(cell, row){
	// changes the production label to the published episode title
	tab2.getRange("G" + row).setValue(uploads.items[0].snippet.title);
	
	// submits the corresponding youtube link
	tab2.getRange("E" + row).setValue("https://youtu.be/" + uploads.items[0].snippet.resourceId.videoId);
}


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

		for (var j = 0; j < playlistResponse.items.length; j++) {
			var playlistItem = playlistResponse.items[j];
			Logger.log('[%s] Title: %s',
					playlistItem.snippet.resourceId.videoId,
					playlistItem.snippet.title);
		}
	}
	
	uploads = playlistResponse;
}