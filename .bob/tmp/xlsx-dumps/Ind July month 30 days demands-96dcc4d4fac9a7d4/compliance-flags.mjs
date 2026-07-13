import { readFileSync } from "fs";

const DUMP_DIR = ".bob/tmp/xlsx-dumps/Ind July month 30 days demands-96dcc4d4fac9a7d4";
const { rows, headers } = JSON.parse(readFileSync(`${DUMP_DIR}/30days.json`, "utf8"));

const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

const today = new Date();
today.setHours(0, 0, 0, 0);

const fifteenDaysAgo = new Date(today);
fifteenDaysAgo.setDate(today.getDate() - 15);

const results = [];

for (const row of rows) {
  const openSeatID     = row[idx["Open Seat ID"]];
  const clientName     = row[idx["Client Name"]];
  const openSeatTitle  = row[idx["Open Seat Title"]];
  const estStrtDt      = row[idx["Est Strt Dt"]];
  const addlComments   = row[idx["Additional Comments"]];
  const candTrackType  = row[idx["Candidate Track Type"]];
  const fulfillAction  = row[idx["Fulfillment Action"]];
  const fieldglassFlag = row[idx["Fieldglass Request Flag"]];

  const flags = [];

  // Rule 1: EST Non Compliant
  // Est Strt Dt is blank OR <= today
  const esdIsBlank = estStrtDt === null || estStrtDt === "" || estStrtDt === undefined;
  let esdDate = null;
  if (!esdIsBlank && typeof estStrtDt === "string") {
    esdDate = new Date(estStrtDt);
  }
  if (esdIsBlank || (esdDate && esdDate <= today)) {
    flags.push("EST Non Compliant");
  }

  // Rule 2: Comment Non Compliant
  // Additional Comments is blank OR last comment is >= 15 days old
  const commentBlank = addlComments === null || addlComments === "" || addlComments === undefined;
  if (commentBlank) {
    flags.push("Comment Non Compliant");
  } else {
    // Extract the most recent date from the comment string
    // Comments follow pattern: "email@domain 2026-07-07 <text>: ..."
    const dateMatches = String(addlComments).match(/\d{4}-\d{2}-\d{2}/g);
    if (dateMatches && dateMatches.length > 0) {
      // Take the first (most recent) date in the comment
      const latestCommentDate = new Date(dateMatches[0]);
      latestCommentDate.setHours(0, 0, 0, 0);
      if (latestCommentDate <= fifteenDaysAgo) {
        flags.push("Comment Non Compliant");
      }
    } else {
      // No parseable date found in comment — treat as non-compliant
      flags.push("Comment Non Compliant");
    }
  }

  // Rule 3: Mismatch in Track Type vs Fieldglass
  // Candidate Track Type = "contractor" but Fieldglass Request Flag is blank or "N"
  const trackTypeLower = (candTrackType || "").toString().trim().toLowerCase();
  const fgFlag = (fieldglassFlag || "").toString().trim().toUpperCase();
  if (trackTypeLower === "contractor") {
    if (fgFlag === "" || fgFlag === "N" || fgFlag === null) {
      flags.push("Mismatch in Track Type vs Fieldglass");
    }
  }

  // Rule 4: Mismatch in Track Type vs Fulfillment Action
  // Candidate Track Type = "Actively recruiting" but Fulfillment Action is NOT "External Hire"
  const openSeatStatus = (row[idx["Open Seat Status"]] || "").toString().trim().toLowerCase();
  if (openSeatStatus === "actively recruiting") {
    const fa = (fulfillAction || "").toString().trim().toLowerCase();
    if (fa !== "external hire") {
      flags.push("Mismatch in Track Type vs Fulfillment Action");
    }
  }

  if (flags.length > 0) {
    results.push({
      "Open Seat ID": openSeatID,
      "Client Name": clientName,
      "Open Seat Title": openSeatTitle,
      "Est Strt Dt": estStrtDt || "",
      "Candidate Track Type": candTrackType || "",
      "Open Seat Status": row[idx["Open Seat Status"]] || "",
      "Fulfillment Action": fulfillAction || "",
      "Fieldglass Request Flag": fieldglassFlag || "",
      "Flags": flags.join(" | ")
    });
  }
}

// Summary counts
const summary = {
  "EST Non Compliant": 0,
  "Comment Non Compliant": 0,
  "Mismatch in Track Type vs Fieldglass": 0,
  "Mismatch in Track Type vs Fulfillment Action": 0
};
for (const r of results) {
  for (const f of r["Flags"].split(" | ")) {
    if (summary[f] !== undefined) summary[f]++;
  }
}

console.log(JSON.stringify({ summary, total: results.length, records: results }, null, 2));
