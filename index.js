import express from "express";
import { google } from "googleapis";

const app = express();
app.use(express.json({ limit: "1mb" }));

const PORT = process.env.PORT || 8080;

// Google Sheets
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const DEBUG_SHEET_NAME = process.env.SHEET_NAME || "debug";
const TARGET_SHEET_NAME = process.env.TARGET_SHEET_NAME || "辻立ち";
const MEMBER_SHEET_NAME = process.env.MEMBER_SHEET_NAME || "メンバー対応表";

// Google Maps Platform
const MAPS_API_KEY = process.env.MAPS_API_KEY;

if (!SPREADSHEET_ID) console.warn("SPREADSHEET_ID is not set.");
if (!MAPS_API_KEY) console.warn("MAPS_API_KEY is not set.");

// ---------- Helpers ----------
function getMessageText(body) {
  if (!body || typeof body !== "object") return "";

  if (
    body.content &&
    body.content.type === "text" &&
    typeof body.content.text === "string"
  ) {
    return body.content.text.trim();
  }

  if (body.content && typeof body.content.text === "string") {
    return body.content.text.trim();
  }

  if (typeof body.text === "string") {
    return body.text.trim();
  }

  if (body.message && typeof body.message.text === "string") {
    return body.message.text.trim();
  }

  return "";
}

function getEventType(body) {
  return body?.type || "";
}

function getAccountId(body) {
  return (
    body?.source?.accountId ||
    body?.source?.userId ||
    body?.source?.senderId ||
    body?.source?.memberId ||
    body?.accountId ||
    body?.userId ||
    body?.senderId ||
    ""
  );
}

function getRoomId(body) {
  return (
    body?.source?.roomId ||
    body?.source?.channelId ||
    body?.roomId ||
    body?.channelId ||
    ""
  );
}

function getJapanMonthDay() {
  const formatter = new Intl.DateTimeFormat("ja-JP", {
    timeZone: "Asia/Tokyo",
    month: "numeric",
    day: "numeric"
  });

  const parts = formatter.formatToParts(new Date());
  const month = parts.find((p) => p.type === "month")?.value || "";
  const day = parts.find((p) => p.type === "day")?.value || "";

  return `${month}/${day}`;
}

function extractUrls(text) {
  if (!text) return [];
  const re = /https?:\/\/[^\s<>\u3000]+/g;
  return text.match(re) || [];
}

function isGoogleMapsUrl(url) {
  return /(^https?:\/\/)?(www\.)?(maps\.app\.goo\.gl|goo\.gl\/maps|google\.[^\/]+\/maps|maps\.google\.[^\/]+)/i.test(
    url
  );
}

function normalizeUrl(url) {
  return String(url || "").trim().replace(/\/+$/, "");
}

async function expandShortUrl(url) {
  let current = url;

  for (let i = 0; i < 5; i++) {
    const res = await fetch(current, {
      method: "GET",
      redirect: "manual"
    });

    const status = res.status;
    const location = res.headers.get("location");

    if ([301, 302, 303, 307, 308].includes(status) && location) {
      current = location;
      continue;
    }

    break;
  }

  return current;
}

function decodeSafe(value) {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function getQueryParam(url, key) {
  const m = url.match(new RegExp(`[?&]${key}=([^&#]+)`, "i"));
  return m ? decodeSafe(m[1]) : "";
}

function cleanupMapsText(s) {
  return decodeSafe(s)
    .replace(/\+/g, " ")
    .replace(/_/g, " ")
    .trim();
}

function extractLikelyQueryFromMapsUrl(text) {
  let m = text.match(/\/place\/([^\/?#]+)/i);
  if (m) return cleanupMapsText(m[1]);

  m = text.match(/\/search\/([^\/?#]+)/i);
  if (m) return cleanupMapsText(m[1]);

  return "";
}

function extractPlaceIdFromText(text) {
  const m = text.match(/(?:place_id|query_place_id)=([^&#]+)/i);
  if (m) return decodeSafe(m[1]);
  return "";
}

function extractLatLng(text) {
  let m = text.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) return { lat: m[1], lng: m[2] };

  m = text.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
  if (m) return { lat: m[1], lng: m[2] };

  return null;
}


function isCoordinateQuery(query) {
  if (!query) return false;

  const cleaned = String(query).trim();

  return /^-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?$/.test(cleaned);
}

function parseCoordinateQuery(query) {
  const m = String(query).trim().match(/^(-?\d+(\.\d+)?)\s*,\s*(-?\d+(\.\d+)?)$/);
  if (!m) return null;

  return {
    lat: m[1],
    lng: m[3]
  };
}

function isCoordinateQuery(query) {
  if (!query) return false;
  return /^-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?$/.test(String(query).trim());
}

function parseCoordinateQuery(query) {
  const m = String(query).trim().match(/^(-?\d+(\.\d+)?)\s*,\s*(-?\d+(\.\d+)?)$/);
  if (!m) return null;

  return {
    lat: m[1],
    lng: m[3]
  };
}

// ---------- Google Sheets ----------
async function getSheetsClient(scopes = ["https://www.googleapis.com/auth/spreadsheets"]) {
  const auth = new google.auth.GoogleAuth({ scopes });
  return google.sheets({ version: "v4", auth });
}

async function appendToSheet(sheetName, rowValues, range = "A:Z") {
  const sheets = await getSheetsClient();

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!${range}`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [rowValues]
    }
  });
}

async function getSenderNameFromSheet(accountId) {
  if (!accountId) return "";

  const sheets = await getSheetsClient([
    "https://www.googleapis.com/auth/spreadsheets.readonly"
  ]);

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${MEMBER_SHEET_NAME}!A:B`
  });

  const rows = res.data.values || [];

  for (const row of rows) {
    if ((row[0] || "").trim() === accountId) {
      return (row[1] || "").trim();
    }
  }

  return accountId;
}

/**
 * 辻立ちシートの J列(URL) に同じURLが既にあるか確認
 */
async function urlAlreadyExists(sheetName, url) {
  const sheets = await getSheetsClient([
    "https://www.googleapis.com/auth/spreadsheets.readonly"
  ]);

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!J:J`
  });

  const values = res.data.values || [];
  const normalized = normalizeUrl(url);

  return values.some((row) => normalizeUrl(row?.[0] || "") === normalized);
}

// ---------- Google Maps / Places ----------
async function getPlaceDetailsByPlaceId(placeId) {
  const url = `https://places.googleapis.com/v1/places/${encodeURIComponent(
    placeId
  )}`;

  const res = await fetch(url, {
    method: "GET",
    headers: {
      "X-Goog-Api-Key": MAPS_API_KEY,
      "X-Goog-FieldMask": "displayName,formattedAddress,plusCode"
    }
  });

  if (!res.ok) {
    throw new Error(`Place Details failed: ${res.status} ${await res.text()}`);
  }

  const data = await res.json();

  return {
    address: data?.formattedAddress || "",
    plusCode:
      data?.plusCode?.globalCode ||
      data?.plusCode?.compoundCode ||
      ""
  };
}

async function searchPlaceByText(query) {
  const url = "https://places.googleapis.com/v1/places:searchText";

  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Goog-Api-Key": MAPS_API_KEY,
      "X-Goog-FieldMask": "places.formattedAddress,places.plusCode"
    },
    body: JSON.stringify({
      textQuery: query,
      languageCode: "ja",
      regionCode: "JP"
    })
  });

  if (!res.ok) {
    throw new Error(`Text Search failed: ${res.status} ${await res.text()}`);
  }

  const data = await res.json();
  const place = data?.places?.[0];

  if (!place) {
    return {
      address: "",
      plusCode: ""
    };
  }

  return {
    address: place?.formattedAddress || "",
    plusCode:
      place?.plusCode?.globalCode ||
      place?.plusCode?.compoundCode ||
      ""
  };
}

async function reverseGeocode(lat, lng) {
  const url = `https://maps.googleapis.com/maps/api/geocode/json?latlng=${lat},${lng}&language=ja&key=${encodeURIComponent(
    MAPS_API_KEY
  )}`;

  const res = await fetch(url);

  if (!res.ok) {
    throw new Error(`Geocoding failed: ${res.status} ${await res.text()}`);
  }

  const data = await res.json();
  const result = data?.results?.[0];

  if (!result) {
    return {
      address: "",
      plusCode: ""
    };
  }

  return {
    address: result?.formatted_address || "",
    plusCode:
      data?.plus_code?.global_code ||
      data?.plus_code?.compound_code ||
      ""
  };
}

async function resolveGoogleMapsLocation(originalUrl) {
  const expandedUrl = await expandShortUrl(originalUrl);
  const decodedUrl = decodeSafe(expandedUrl);

  const placeId =
    getQueryParam(expandedUrl, "query_place_id") ||
    getQueryParam(expandedUrl, "place_id") ||
    extractPlaceIdFromText(decodedUrl);

  if (placeId) {
    try {
      const detail = await getPlaceDetailsByPlaceId(placeId);
      return {
        address: detail.address || "",
        plusCode: detail.plusCode || "",
        url: originalUrl
      };
    } catch (err) {
      console.warn("Place Details fallback:", String(err));
    }
  }

  const query =
    getQueryParam(expandedUrl, "query") ||
    getQueryParam(expandedUrl, "q") ||
    extractLikelyQueryFromMapsUrl(decodedUrl);

if (query) {
  try {
    const searched = await searchPlaceByText(query);
    return {
      address: searched.address || "",
      plusCode: searched.plusCode || "",
      url: originalUrl
    };
  } catch (err) {
    console.warn("Text Search failed, but continue with blank address:", String(err));

    return {
      address: "",
      plusCode: "",
      url: originalUrl
    };
  }
}


  throw err;
}
  const latlng = extractLatLng(decodedUrl);
  if (latlng) {
    const geo = await reverseGeocode(latlng.lat, latlng.lng);
    return {
      address: geo.address || "",
      plusCode: geo.plusCode || "",
      url: originalUrl
    };
  }

  return {
    address: "",
    plusCode: "",
    url: originalUrl
  };
}

// ---------- Routes ----------
app.get("/", (_req, res) => {
  res.status(200).send("cloud run is alive v2");
});

app.post("/", async (req, res) => {
  const body = req.body || {};
  const rawBody = JSON.stringify(body);

  const eventType = getEventType(body);
  const text = getMessageText(body);
  const accountId = getAccountId(body);
  const roomId = getRoomId(body);
  const senderName = await getSenderNameFromSheet(accountId);
  const urls = extractUrls(text);
  const mapUrls = urls.filter(isGoogleMapsUrl);

  try {
    // debugシートへ記録
    await appendToSheet(DEBUG_SHEET_NAME, [
      new Date().toISOString(),
      eventType,
      text,
      senderName,
      accountId,
      roomId,
      rawBody
    ]);

    if (eventType !== "message") {
      return res.status(200).json({ ok: true, skipped: "not message" });
    }

    if (!mapUrls.length) {
      return res.status(200).json({ ok: true, skipped: "no google maps url" });
    }

    for (const mapUrl of mapUrls) {
      const exists = await urlAlreadyExists(TARGET_SHEET_NAME, mapUrl);
      if (exists) continue;

      const location = await resolveGoogleMapsLocation(mapUrl);

      const addressOrPlusCode =
        location.address || location.plusCode || "";

      // B:K に追記
      // B:番号（手入力）
      // C:日付（自動）
      // D:東国原（手入力）
      // E:入力者（自動）
      // F:G 時刻（手入力）
      // H: 時間（手入力
      // I:所在地（自動）
      // J:所在地URL（自動）
      // K:備考（手入力）
      await appendToSheet(
        TARGET_SHEET_NAME,
        [
          "", // B 番号
          getJapanMonthDay(), // C 日付
          "", // D 東国原
          senderName || accountId || "", // E 入力者
          "", // F 時刻
          "", // G 時刻
          "", // H 実施時間
          addressOrPlusCode, // H 所在地
          mapUrl, // I 所在地URL
          "" // J 備考
        ],
        "B:K"
      );
    }

    return res.status(200).json({ ok: true });
  } catch (err) {
    console.error(err);

    try {
      await appendToSheet(DEBUG_SHEET_NAME, [
        new Date().toISOString(),
        "ERROR",
        "",
        senderName,
        accountId,
        roomId,
        String(err)
      ]);
    } catch (e) {
      console.error("append debug error failed", e);
    }

    return res.status(200).json({
      ok: false,
      error: String(err)
    });
  }
});

app.listen(PORT, () => {
  console.log(`Listening on port ${PORT}`);
});
