function myFunction() {
  // if (MailApp.getRemainingDailyQuota() == 0.0) return;
  // else Logger.log(MailApp.getRemainingDailyQuota());
  
  Logger.log(MailApp.getRemainingDailyQuota());
  retryEmailByRow(546);

  // retryEmailByRow(865);
  // <-- Stop -->

  // for (let i = 826; i <= 832; i++) {
	//  	retryEmailByRow(i);
	// }

  // let rows = [834, 835, 836, 837, 838, 841, 845, 849, 850, 851, 858, 859, 860, 862, 877, 878, 880, 881, 888, 891, 892, 893, 895, 897 , 898];

	// for (let row of rows) {
	//  	retryEmailByRow(row);
	// }
}

function doGet() {
	getData();
	getDataByMembership();
	return HtmlService
        .createHtmlOutputFromFile('index.html')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getData() {
	const cache = CacheService.getScriptCache();
	const cached = cache.get("examMap");
	if (cached) return JSON.parse(cached);

	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
	const values = sheet.getDataRange().getValues();

	const examMap = Object.create(null);

	for (let i = 1; i < values.length; i++) {
		const examId = String(values[i][9]).trim();
		if (!examId) continue;

		examMap[examId] = {
			id: String(values[i][2] ?? "").trim(),
			phone: String(values[i][7] ?? "").trim(),
		};
	}

	try {
		cache.put("examMap", JSON.stringify(examMap), 60);
	} catch (e) {
		Logger.log("Cache put failed: " + e.message);
	}

	return examMap;
}

function findID(examId, phone) {
	try {
		const examMap = getData();
		const key = String(examId).trim();

		if (!Object.prototype.hasOwnProperty.call(examMap, key)) {
			return { found: false };
		}

		const entry = examMap[key];
		if (String(entry.phone).trim() !== String(phone).trim()) {
			return { found: false };
		}

		return { found: true, membershipId: entry.id };
	} catch (error) {
		return { found: false, error: error.message };
	}
}

function invalidateExamCache() {
	const cache = CacheService.getScriptCache();
	cache.remove("examMap");
	cache.remove("membershipMap");
}

function getDataByMembership() {
	const cache = CacheService.getScriptCache();
	const cached = cache.get("membershipMap");
	if (cached) return JSON.parse(cached);

	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
	const allValues = sheet.getDataRange().getValues();
	const headers = allValues[0];

	const map = Object.create(null);

	for (let i = 1; i < allValues.length; i++) {
		const membershipId = String(allValues[i][2] ?? "").trim();
		if (!membershipId) continue;

		const record = {};
		headers.forEach((h, j) => { record[String(h)] = allValues[i][j]; });

		const isTrue = (v) => v === true || String(v).toUpperCase() === 'TRUE';

		map[membershipId] = {
			row:            i + 1,
			jerseySize:     String(record["ไซส์เสื้อ Jersey"]        ?? "").trim() || "-",
			jerseyText:     String(record["ตัวอักษรบนเสื้อ Jersey"]  ?? "").trim() || "-",
			jerseyNumber:   String(record["ตัวเลขบนเสื้อ Jersey"]    ?? "").trim() || "-",
			poloSize:       String(record["ไซส์เสื้อโปโล"]           ?? "").trim() || "-",
			poloGender:     String(record["เพศเสื้อโปโล"]            ?? "").trim() || "-",
			bagCollected:   isTrue(allValues[i][33]),
			poloCollected:  isTrue(allValues[i][34]),
			jerseyCollected:isTrue(allValues[i][35]),
		};
	}

	try {
		cache.put("membershipMap", JSON.stringify(map), 60);
	} catch (e) {
		Logger.log("Cache put failed (membershipMap): " + e.message);
	}

	return map;
}

function findByMembershipId(rawInput) {
	try {
		const num = parseInt(String(rawInput).replace(/\D/g, ""), 10);
		if (isNaN(num) || num < 1) return { found: false };

		const map = getDataByMembership();
		const key = "89-" + num;

		if (!Object.prototype.hasOwnProperty.call(map, key)) {
			return { found: false };
		}

		return { found: true, membershipId: key, ...map[key] };
	} catch (err) {
		return { found: false, error: err.message };
	}
}

function toggleCollectedStatus(membershipId, type) {
	// AH=34 (bag), AI=35 (polo), AJ=36 (jersey) — 1-indexed
	const COL = { bag: 34, polo: 35, jersey: 36 };
	const col = COL[type];
	if (!col) return { success: false, error: "Invalid type" };

	try {
		const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
		const allValues = sheet.getDataRange().getValues();
		const key = String(membershipId).trim();

		for (let i = 1; i < allValues.length; i++) {
			if (String(allValues[i][2] ?? "").trim() !== key) continue;

			const current = allValues[i][col - 1];
			const newStatus = !(current === true || String(current).toUpperCase() === 'TRUE');
			sheet.getRange(i + 1, col).setValue(newStatus);
			CacheService.getScriptCache().remove("membershipMap");
			return { success: true, newStatus };
		}

		return { success: false, error: "Member not found" };
	} catch (err) {
		return { success: false, error: err.message };
	}
}