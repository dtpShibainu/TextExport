var doc = app.activeDocument;
var extractedText = "";

for (var i = 0; i < doc.pages.length; i++) {
    var page = doc.pages[i];
    var pageText = "";
    var pageStories = [];
    
    for (var j = 0; j < doc.stories.length; j++) {
        var story = doc.stories[j];
        if (story.textContainers.length > 0) {
            var firstContainer = story.textContainers[0];
            if (firstContainer.itemLayer.locked || !firstContainer.itemLayer.visible) continue;
            if (firstContainer.parentPage === null || firstContainer.parentPage.parent instanceof MasterSpread) continue;
            if (firstContainer.parentPage == page) {
                var storyData = {
                    story: story,
                    x: firstContainer.geometricBounds[1], // 左端 (X座標)
                    y: firstContainer.geometricBounds[0], // 上端 (Y座標)
                    isGrouped: false,
                    groupId: null,
                    id: null
                };
                if (firstContainer.parent instanceof Group) {
                    storyData.isGrouped = true;
                    storyData.groupId = firstContainer.parent.id;
                } else {
                    storyData.groupId = "single_" + (i + 1) + "_" + (j + 1);
                }
                pageStories.push(storyData);
            }
        }
    }

    // Y座標で降順ソート（上から下へ）
    // Y座標が同じ場合はX座標で降順ソート（右から左へ）
    pageStories.sort(function(a, b) {
        if (a.y !== b.y) {
            return b.y - a.y; // Y座標の降順
        }
        return b.x - a.x; // X座標の降順（右から左へ）
    });

    for (var k = 0; k < pageStories.length; k++) {
        pageStories[k].id = k + 1;
    }

    var groupedTextFrames = {};
    for (var k = 0; k < pageStories.length; k++) {
        var storyData = pageStories[k];
        if (!groupedTextFrames[storyData.groupId]) {
            groupedTextFrames[storyData.groupId] = [];
        }
        groupedTextFrames[storyData.groupId].push(storyData);
    }

    var finalSortedStories = [];
    for (var groupId in groupedTextFrames) {
        if (groupedTextFrames.hasOwnProperty(groupId)) {
            finalSortedStories = finalSortedStories.concat(groupedTextFrames[groupId]);
        }
    }

    var previousGroupId = null;
    for (var k = 0; k < finalSortedStories.length; k++) {
        var storyContent = finalSortedStories[k].story.contents;
        var tableContent = "";
        
        // テーブルが含まれる場合の処理
        var tables = finalSortedStories[k].story.tables.everyItem().getElements();
        if (tables.length > 0) {
            for (var t = 0; t < tables.length; t++) {
                var table = tables[t];
                for (var r = 0; r < table.rows.length; r++) {
                    var rowText = [];
                    for (var c = 0; c < table.columns.length; c++) {
                        rowText.push(table.rows[r].cells[c].contents);
                    }
                    tableContent += rowText.join(" | ") + "\n";
                }
            }
        }
        
        var text = storyContent + (tableContent ? "\n" + tableContent : "");
        if (text.length > 0) {
            text = text.replace(/\n+/g, "\n");
            if (previousGroupId && previousGroupId !== finalSortedStories[k].groupId) {
                pageText += "\n";
            }
            pageText += text + "\n";
            previousGroupId = finalSortedStories[k].groupId;
        }
    }

    if (pageText.length > 0) {
        extractedText += "【ページ " + (page.documentOffset + 1) + "】\n" + pageText + "\n";
    }
}

outputFile(extractedText);
function outputFile(str) {
    var file = File.saveDialog("ファイル書き出し", "テキストファイル:*.txt");
    if (file) {
        if (!file.name.match(/\.txt$/)) {
            file = new File(file.fsName + ".txt");
        }
        file.encoding = "UTF-8";
        file.open("w");
        file.write(str);
        file.close();
        alert("ファイルが保存されました: " + file.fsName);
    } else {
        alert("保存がキャンセルされました");
    }
}
