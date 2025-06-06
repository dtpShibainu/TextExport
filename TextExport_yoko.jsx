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
            // レイヤーロック or 非表示レイヤーはスキップ
            if (firstContainer.itemLayer.locked || !firstContainer.itemLayer.visible) {
                continue;
            }
            if (firstContainer.parentPage === null || firstContainer.parentPage.parent instanceof MasterSpread) {
                continue;
            }
            // ページに属しているストーリーを収集
            if (firstContainer.parentPage == page) {
                var storyData = {
                    story: story,
                    x: firstContainer.geometricBounds[1], // 左端 (X座標)
                    y: firstContainer.geometricBounds[0], // 上端 (Y座標)
                    isGrouped: false, // グループフラグ（初期値はfalse）
                    groupId: null, // グループIDを追加（最初はnull）
                    id: null // IDを追加
                };
                // テキストフレームがグループ化されている場合、同じグループに属するテキストフレームをまとめる
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

    // Y座標で並び替え、同じYの場合はX座標で並べ替え
    pageStories.sort(function(a, b) {
        // Y座標が異なる場合は、Y座標で昇順にソート（上にあるものを優先）
        if (a.y !== b.y) {
            return a.y - b.y;
        }
        // 同じY座標にある場合、X座標で昇順にソート（左にあるものを優先）
        return a.x - b.x;
    });
    // 並べ替えた順番でIDを割り当て
    for (var k = 0; k < pageStories.length; k++) {
        pageStories[k].id = k + 1; // 1からIDを付与
    }
    // グループ化されたテキストフレームをまとめる
    var groupedTextFrames = {};
    for (var k = 0; k < pageStories.length; k++) {
        var storyData = pageStories[k];
        if (storyData.isGrouped && storyData.groupId) {
            if (!groupedTextFrames[storyData.groupId]) {
                groupedTextFrames[storyData.groupId] = [];
            }
            groupedTextFrames[storyData.groupId].push(storyData);
        } else {
            if (!groupedTextFrames[storyData.groupId]) {
                groupedTextFrames[storyData.groupId] = [];
            }
            groupedTextFrames[storyData.groupId].push(storyData);
        }
    }

    // グループIDごとにまとめて並べ替え
    var finalSortedStories = [];
    for (var groupId in groupedTextFrames) {
    if (groupedTextFrames.hasOwnProperty(groupId)) {
        finalSortedStories = finalSortedStories.concat(groupedTextFrames[groupId]);
    }
    }

    // 最終的に並べ替えた順番でテキストを取得
    var previousGroupId = null;
    for (var k = 0; k < finalSortedStories.length; k++) {
    var text = finalSortedStories[k].story.contents;
    if (text.length > 0) {
        // 改行が複数入る場合、1つの改行にまとめる
        text = text.replace(/\n+/g, "\n");
        // 異なるグループ間では改行を2つにする
        if (previousGroupId && previousGroupId !== finalSortedStories[k].groupId) {
            pageText += "\n";
        }
        pageText += text + "\n";
        previousGroupId = finalSortedStories[k].groupId;
    }
    }

// ページごとのテキストを追加
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
        file.encoding = "UTF-8"; // UTF-8でエンコード
        file.open("w");
        file.write(str);
        file.close();
        alert("ファイルが保存されました: " + file.fsName);
    } else {
        alert("保存がキャンセルされました");
    }
}
