import { readFileSync, writeFileSync } from "fs";

const currentYear = new Date().getFullYear();
const years = [currentYear, currentYear + 1];
const result = {};

for (const year of years) {
  try {
    const res = await fetch(`https://api.jiejiariapi.com/v1/holidays/${year}`);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    const offDays = [];
    const workDays = [];
    for (const [, item] of Object.entries(data)) {
      if (!item?.date) continue;
      const entry = { date: item.date, name: item.name || "" };
      if (item.isOffDay) {
        offDays.push(entry);
      } else {
        const day = new Date(item.date + "T00:00:00").getDay();
        if (day === 0 || day === 6) workDays.push(entry);
      }
    }
    offDays.sort((a, b) => a.date.localeCompare(b.date));
    workDays.sort((a, b) => a.date.localeCompare(b.date));
    result[year] = {
      source: "api.jiejiariapi.com",
      updatedAt: new Date().toISOString(),
      offDays,
      workDays,
    };
    console.log(`Fetched ${year}: ${offDays.length} off days, ${workDays.length} work days`);
  } catch (e) {
    console.error(`Failed to fetch ${year}:`, e.message);
  }
}

let existing = {};
try {
  existing = JSON.parse(readFileSync("holidays.json", "utf-8"));
} catch {}

const merged = { ...existing, ...result };
writeFileSync("holidays.json", JSON.stringify(merged, null, 2) + "\n", "utf-8");
console.log("holidays.json updated, years:", Object.keys(merged));
