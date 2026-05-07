function myFunction() {
}

function doGet() {
	buildAndCacheMaps_();
	return HtmlService
		.createHtmlOutputFromFile('index.html')
		.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheet_() {
	return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
}

function isTrue_(v) {
	return v === true || String(v).toUpperCase() === "TRUE";
}

function isDenied_(v) {
	const s = String(v ?? "").trim().toLowerCase();
	return s === "denied";
}

function str_(v) {
	return String(v ?? "").trim() || "-";
}

function joinName_(first, last) {
	return [String(first ?? "").trim(), String(last ?? "").trim()].filter(Boolean).join(" ") || "-";
}

// Reads the sheet once and populates both caches atomically.
function buildAndCacheMaps_() {
	const values  = getSheet_().getDataRange().getValues();
	const headers = values[0];

	const hIdx = Object.create(null);
	headers.forEach((h, i) => { hIdx[String(h)] = i; });

	const jerseySize_idx   = hIdx["ไซส์เสื้อ Jersey"];
	const jerseyText_idx   = hIdx["ตัวอักษรบนเสื้อ Jersey"];
	const jerseyNumber_idx = hIdx["ตัวเลขบนเสื้อ Jersey"];
	const poloSize_idx     = hIdx["ไซส์เสื้อโปโล"];
	const poloGender_idx   = hIdx["เพศเสื้อโปโล"];

	const examMap       = Object.create(null);
	const membershipMap = Object.create(null);

	for (let i = 1; i < values.length; i++) {
		const row          = values[i];
		const membershipId = String(row[2] ?? "").trim();
		const examId       = String(row[9] ?? "").trim();

		if (examId) {
			examMap[examId] = {
				id:    membershipId,
				phone: String(row[7] ?? "").trim(),
			};
		}

		if (membershipId) {
			membershipMap[membershipId] = {
				row:             i + 1,
				parentName:      joinName_(row[3],  row[4]),
				studentName:     joinName_(row[13], row[14]),
				jerseySize:      str_(row[jerseySize_idx]),
				jerseyText:      str_(row[jerseyText_idx]),
				jerseyNumber:    str_(row[jerseyNumber_idx]),
				poloSize:        str_(row[poloSize_idx]),
				poloGender:      str_(row[poloGender_idx]),
				denyStatus: str_(row[32]),
				isDenied: isDenied_(row[32]),
				bagCollected:    isTrue_(row[33]),
				poloCollected:   isTrue_(row[34]),
				jerseyCollected: isTrue_(row[35]),
			};
		}
	}

	const cache = CacheService.getScriptCache();
	try { cache.put("examMap",       JSON.stringify(examMap),       60); } catch (e) { Logger.log(e.message); }
	try { cache.put("membershipMap", JSON.stringify(membershipMap), 60); } catch (e) { Logger.log(e.message); }

	return { examMap, membershipMap };
}

function getData() {
	const cached = CacheService.getScriptCache().get("examMap");
	if (cached) return JSON.parse(cached);
	return buildAndCacheMaps_().examMap;
}

function getDataByMembership() {
	const cached = CacheService.getScriptCache().get("membershipMap");
	if (cached) return JSON.parse(cached);
	return buildAndCacheMaps_().membershipMap;
}

function invalidateExamCache() {
	const cache = CacheService.getScriptCache();
	cache.remove("examMap");
	cache.remove("membershipMap");
}

function findID(examId, phone) {
	try {
		const entry = getData()[String(examId).trim()];
		if (!entry) return { found: false };
		if (String(entry.phone).trim() !== String(phone).trim()) return { found: false };
		return { found: true, membershipId: entry.id };
	} catch (error) {
		return { found: false, error: error.message };
	}
}

function findByMembershipId(rawInput) {
	try {
		const num = parseInt(String(rawInput).replace(/\D/g, ""), 10);
		if (isNaN(num) || num < 1) return { found: false };

		const map = getDataByMembership();
		const key = "89-" + num;
		if (!map[key]) return { found: false };

		return { found: true, membershipId: key, ...map[key] };
	} catch (err) {
		return { found: false, error: err.message };
	}
}

function toggleCollectedStatus(membershipId, type) {
	const COL = { bag: 34, polo: 35, jersey: 36 };
	const col = COL[type];
	if (!col) return { success: false, error: "Invalid type" };

	try {
		const key   = String(membershipId).trim();
		const entry = getDataByMembership()[key];
		if (!entry) return { success: false, error: "Member not found" };

		const sheet     = getSheet_();
		const current   = sheet.getRange(entry.row, col).getValue();
		const newStatus = !isTrue_(current);
		sheet.getRange(entry.row, col).setValue(newStatus);
		CacheService.getScriptCache().remove("membershipMap");
		return { success: true, newStatus };
	} catch (err) {
		return { success: false, error: err.message };
	}
}
