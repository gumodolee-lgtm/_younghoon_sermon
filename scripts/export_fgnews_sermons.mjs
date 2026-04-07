import { promises as fs } from "node:fs";
import path from "node:path";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";

const LIST_BASE = "http://pastorlee.fgtv.com/fgnews/pastorlee_list.asp";
const START_DATE = "2021-01-03";
const END_DATE = "2025-12-28";
const MAX_WORDS_PER_NOTEBOOKLM_SOURCE = 420000;
const OUTPUT_ROOT = path.resolve(process.cwd(), "output", "fgnews_2021_2025");

const USER_AGENT =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36";

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchText(url, { retries = 3, timeoutMs = 30000, waitMs = 250 } = {}) {
  let lastError;
  for (let i = 0; i < retries; i += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const res = await fetch(url, {
        headers: {
          "user-agent": USER_AGENT
        },
        signal: controller.signal
      });
      if (!res.ok) {
        throw new Error("No sermons found in the requested date range.");
      }
      const text = await res.text();
      clearTimeout(timer);
      return text;
    } catch (error) {
      clearTimeout(timer);
      lastError = error;
      if (i < retries - 1) {
        await sleep(waitMs * (i + 1));
      }
    }
  }
  throw new Error("No sermons found in the requested date range.");
}

function decodeHtmlEntities(input) {
  if (!input) return "";
  const named = {
    amp: "&",
    lt: "<",
    gt: ">",
    quot: '"',
    apos: "'",
    nbsp: " ",
    ldquo: '"',
    rdquo: '"',
    lsquo: "'",
    rsquo: "'",
    hellip: "...",
    mdash: "-",
    ndash: "-"
  };

  return input
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, code) => String.fromCodePoint(Number.parseInt(code, 16)))
    .replace(/&([a-zA-Z]+);/g, (full, key) => named[key] ?? full);
}

function stripTagsToText(html) {
  let out = html ?? "";
  out = out.replace(/<script[\s\S]*?<\/script>/gi, "");
  out = out.replace(/<style[\s\S]*?<\/style>/gi, "");
  out = out.replace(/<br\s*\/?>/gi, "\n");
  out = out.replace(/<\/(p|div|h1|h2|h3|h4|h5|h6|li|tr|td|table|blockquote)>/gi, "\n");
  out = out.replace(/<[^>]+>/g, "");
  out = decodeHtmlEntities(out);
  out = out.replace(/\u00a0/g, " ");
  out = out.replace(/\r/g, "");
  return out;
}

function normalizeLine(line) {
  return line.replace(/\s+/g, " ").trim();
}

function stripPrayerAndHymn(lines) {
  const explicitLiturgyRegex =
    /(<\s*\uCC2C\uC1A1\uAC00|<\s*\uAE30\uB3C4\s*>|<\s*\uCD95\uB3C4\s*>)/i;
  const genericTailRegex =
    /(\uC624\uB298 \uB9C8\uC9C0\uB9C9 \uCC2C\uC591|\uCC2C\uC1A1\uAC00\s*\d+\s*\uC7A5|\uAE30\uB3C4\uD558\uACA0\uC2B5\uB2C8\uB2E4|^\(?\uAE30\uB3C4\)?$)/i;

  const cutAtMatch = (startIndex, regex) => {
    for (let i = startIndex; i < lines.length; i += 1) {
      const matched = lines[i].match(regex);
      if (!matched) continue;

      const prefix = lines[i].slice(0, matched.index).trim();
      const before = lines.slice(0, i);
      if (prefix) before.push(prefix);
      return before;
    }
    return null;
  };

  const explicitStart = Math.floor(lines.length * 0.2);
  const explicitCut = cutAtMatch(explicitStart, explicitLiturgyRegex);
  if (explicitCut) return explicitCut;

  const genericStart = Math.floor(lines.length * 0.35);
  const genericCut = cutAtMatch(genericStart, genericTailRegex);
  if (genericCut) return genericCut;

  return lines;
}

function stripTailByMarkers(text) {
  const markers = [
    /\uCC2C\uC1A1\uAC00\s*\d+\s*\uC7A5/i,
    /\uAE30\uB3C4\uD558\uACA0\uC2B5\uB2C8\uB2E4/i,
    /\uC624\uB298 \uB9C8\uC9C0\uB9C9 \uCC2C\uC591/i,
    /<\s*\uCC2C\uC1A1\uAC00/i,
    /<\s*\uAE30\uB3C4\s*>/i
  ];

  const start = Math.floor(text.length * 0.35);
  let cut = -1;
  const tail = text.slice(start);
  for (const re of markers) {
    const localIndex = tail.search(re);
    if (localIndex < 0) continue;
    const absoluteIndex = start + localIndex;
    if (cut < 0 || absoluteIndex < cut) {
      cut = absoluteIndex;
    }
  }

  if (cut >= 0) {
    return text.slice(0, cut).trim();
  }
  return text;
}
function toWordCount(text) {
  if (!text.trim()) return 0;
  return text.trim().split(/\s+/).filter(Boolean).length;
}

function toDateIso(dotDate) {
  const m = String(dotDate).match(/^(\d{4})\.(\d{2})\.(\d{2})\.$/);
  if (!m) return null;
  return `${m[1]}-${m[2]}-${m[3]}`;
}

function cleanInlineText(text) {
  return normalizeLine(stripTagsToText(text));
}

function parseMaxPage(listHtml) {
  const nums = [...listHtml.matchAll(/pastorlee_list\.asp\?page=(\d+)/gi)].map((m) => Number(m[1]));
  if (nums.length === 0) return 1;
  return Math.max(...nums);
}

function parseListRows(listHtml) {
  const rows = [];
  const trRegex = /<tr>[\s\S]*?<\/tr>/gi;
  for (const trMatch of listHtml.matchAll(trRegex)) {
    const rowHtml = trMatch[0];
    const datePieces = [...rowHtml.matchAll(/<p class="sermon_date">([\s\S]*?)<\/p>/gi)].map((m) =>
      cleanInlineText(m[1])
    );
    const dateIso = toDateIso(datePieces[0]);
    if (!dateIso) continue;

    const titleMatch = rowHtml.match(/<a class="sermon_title bolder"[^>]*>([\s\S]*?)<a><p class="sermon_date">/i);
    const textUrlMatch = rowHtml.match(/<a class="sermon_bt_13" href="([^"]+)"/i);
    if (!textUrlMatch) continue;

    rows.push({
      date: dateIso,
      category: datePieces[1] ?? "",
      scripture: datePieces[2] ?? "",
      title: cleanInlineText(titleMatch?.[1] ?? ""),
      textUrl: textUrlMatch[1].replace("https://", "http://")
    });
  }
  return rows;
}

function extractBetween(source, startToken, endToken) {
  const startIdx = source.indexOf(startToken);
  if (startIdx < 0) return "";
  const fromStart = source.slice(startIdx + startToken.length);
  const endIdx = fromStart.indexOf(endToken);
  if (endIdx < 0) return "";
  return fromStart.slice(0, endIdx);
}

function parseDetail(detailHtml, fallbackTitle) {
  const titleMatch = detailHtml.match(/<div class="view_title[^"]*">\s*([\s\S]*?)\s*<\/div>/i);
  const title = cleanInlineText(titleMatch?.[1] ?? fallbackTitle ?? "");

  let contentHtml = extractBetween(detailHtml, '<div class="cls-contents">', '<textarea id="tts_area"');
  if (!contentHtml.trim()) {
    const ttsMatch = detailHtml.match(/<textarea id="tts_area"[^>]*>([\s\S]*?)<\/textarea>/i);
    contentHtml = ttsMatch?.[1] ?? "";
  }

  contentHtml = contentHtml.replace(/<\/div>\s*$/i, "");
  let lines = stripTagsToText(contentHtml)
    .split("\n")
    .map((line) => normalizeLine(line))
    .filter(Boolean);

  // Some pages include front liturgy blocks before sermon body.
  const sermonStartIdx = lines.findIndex((line) => /<\s*설교말씀\s*>|^\s*설교말씀\s*$/i.test(line));
  if (sermonStartIdx >= 0 && sermonStartIdx < Math.floor(lines.length * 0.5)) {
    lines = lines.slice(sermonStartIdx + 1);
  }

  lines = stripPrayerAndHymn(lines);

  const deduped = [];
  for (const line of lines) {
    if (deduped.length === 0 || deduped[deduped.length - 1] !== line) {
      deduped.push(line);
    }
  }

  let body = deduped.join("\n").trim();
  body = stripTailByMarkers(body);
  return { title, body };
}

async function mapWithConcurrency(items, limit, worker) {
  const out = new Array(items.length);
  let cursor = 0;

  async function runner() {
    while (true) {
      const idx = cursor;
      cursor += 1;
      if (idx >= items.length) break;
      out[idx] = await worker(items[idx], idx);
    }
  }

  const workers = [];
  for (let i = 0; i < Math.min(limit, items.length); i += 1) {
    workers.push(runner());
  }
  await Promise.all(workers);
  return out;
}

function groupByYear(sermons) {
  const map = new Map();
  for (const sermon of sermons) {
    const year = sermon.date.slice(0, 4);
    if (!map.has(year)) map.set(year, []);
    map.get(year).push(sermon);
  }
  for (const [, list] of map) {
    list.sort((a, b) => a.date.localeCompare(b.date));
  }
  return map;
}

function buildNotebookLmPacks(yearEntries, maxWords) {
  const packs = [];
  let current = { sermons: [], words: 0, years: new Set() };

  const pushCurrent = () => {
    if (current.sermons.length > 0) {
      packs.push({
        sermons: current.sermons,
        words: current.words,
        years: [...current.years].sort()
      });
      current = { sermons: [], words: 0, years: new Set() };
    }
  };

  const appendSermon = (sermon) => {
    current.sermons.push(sermon);
    current.words += sermon.wordCount;
    current.years.add(sermon.date.slice(0, 4));
  };

  for (const [year, yearSermons] of yearEntries) {
    const yearWords = yearSermons.reduce((acc, s) => acc + s.wordCount, 0);
    if (yearWords <= maxWords) {
      if (current.words + yearWords > maxWords) {
        pushCurrent();
      }
      for (const sermon of yearSermons) appendSermon(sermon);
      continue;
    }

    pushCurrent();
    for (const sermon of yearSermons) {
      if (current.words + sermon.wordCount > maxWords) {
        pushCurrent();
      }
      appendSermon(sermon);
    }
    pushCurrent();
    current = { sermons: [], words: 0, years: new Set([year]) };
    current.years.clear();
  }
  pushCurrent();
  return packs;
}

async function writeDocx(filePath, title, sermons) {
  const children = [];
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text: title, bold: true })]
    })
  );
  children.push(
    new Paragraph({
      children: [new TextRun(`Date Range: ${START_DATE} ~ ${END_DATE}`)]
    })
  );
  children.push(
    new Paragraph({
      children: [new TextRun(`Sermon Count: ${sermons.length}`)]
    })
  );
  children.push(new Paragraph({ text: "" }));

  sermons.forEach((sermon, idx) => {
    if (idx > 0) {
      children.push(new Paragraph({ text: "", pageBreakBefore: true }));
    }
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun(`${sermon.date} | ${sermon.title}`)]
      })
    );
    children.push(new Paragraph({ children: [new TextRun(`Scripture: ${sermon.scripture || "-"}`)] }));
    children.push(new Paragraph({ children: [new TextRun(`Category: ${sermon.category || "-"}`)] }));
    children.push(new Paragraph({ children: [new TextRun(`Source: ${sermon.textUrl}`)] }));
    children.push(new Paragraph({ text: "" }));

    const lines = sermon.body.split("\n");
    for (const line of lines) {
      children.push(new Paragraph({ children: [new TextRun(line)] }));
    }
  });

  const doc = new Document({
    sections: [{ children }]
  });
  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(filePath, buffer);
}

async function main() {
  await fs.mkdir(OUTPUT_ROOT, { recursive: true });

  console.log("[1/6] Collecting list pages");
  const firstListHtml = await fetchText(LIST_BASE);
  const maxPage = parseMaxPage(firstListHtml);
  console.log(`- list pages: ${maxPage}`);

  const allRows = [];
  for (let page = 1; page <= maxPage; page += 1) {
    const url = page === 1 ? LIST_BASE : `${LIST_BASE}?page=${page}`;
    const html = page === 1 ? firstListHtml : await fetchText(url);
    const rows = parseListRows(html);
    allRows.push(...rows);
    if (page % 10 === 0 || page === maxPage) {
      console.log(`- list progress: ${page}/${maxPage}`);
    }
    await sleep(80);
  }

  const rowMap = new Map();
  for (const row of allRows) {
    const key = `${row.date}|${row.textUrl}`;
    if (!rowMap.has(key)) rowMap.set(key, row);
  }
  const dedupedRows = [...rowMap.values()];

  const filtered = dedupedRows
    .filter((row) => row.date >= START_DATE && row.date <= END_DATE)
    .sort((a, b) => a.date.localeCompare(b.date));

  console.log(`[2/6] Filtered range rows: ${filtered.length}`);
  if (filtered.length === 0) {
    throw new Error("No sermons found in the requested date range.");
  }

  console.log("[3/6] Fetching sermon details");
  const detailed = await mapWithConcurrency(filtered, 4, async (row, idx) => {
    const html = await fetchText(row.textUrl);
    const parsed = parseDetail(html, row.title);
    const body = parsed.body;
    const wordCount = toWordCount(body);
    const result = {
      ...row,
      title: parsed.title || row.title,
      body,
      wordCount
    };
    if ((idx + 1) % 20 === 0 || idx + 1 === filtered.length) {
      console.log(`- detail progress: ${idx + 1}/${filtered.length}`);
    }
    await sleep(40);
    return result;
  });

  const usable = detailed.filter((item) => item.body.trim().length > 0);
  if (usable.length !== detailed.length) {
    console.log(`- dropped empty bodies: ${detailed.length - usable.length}`);
  }
  console.log(`[4/6] Cleaned sermons: ${usable.length}`);

  const byYear = groupByYear(usable);
  const years = [...byYear.keys()].sort();

  const yearlyOutputDir = path.join(OUTPUT_ROOT, "yearly_docx");
  const mergedOutputDir = path.join(OUTPUT_ROOT, "notebooklm_merged_docx");
  await fs.mkdir(yearlyOutputDir, { recursive: true });
  await fs.mkdir(mergedOutputDir, { recursive: true });

  console.log("[5/6] Writing yearly DOCX files");
  const yearlyStats = [];
  for (const year of years) {
    const sermons = byYear.get(year);
    const words = sermons.reduce((acc, s) => acc + s.wordCount, 0);
    const filePath = path.join(yearlyOutputDir, `${year}_sermons.docx`);
    await writeDocx(filePath, `Younghoon Lee Sunday Sermons ${year}`, sermons);
    yearlyStats.push({ year, sermons: sermons.length, words, file: filePath });
    console.log(`- ${year}: ${sermons.length} sermons, ${words.toLocaleString()} words`);
  }

  console.log("[6/6] Writing NotebookLM merged DOCX files");
  const yearEntries = [...byYear.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const packs = buildNotebookLmPacks(yearEntries, MAX_WORDS_PER_NOTEBOOKLM_SOURCE);
  const mergedStats = [];
  for (let i = 0; i < packs.length; i += 1) {
    const pack = packs[i];
    const startYear = pack.years[0];
    const endYear = pack.years[pack.years.length - 1];
    const yearLabel = startYear === endYear ? startYear : `${startYear}-${endYear}`;
    const fileName = `NotebookLM_pack_${String(i + 1).padStart(2, "0")}_${yearLabel}.docx`;
    const filePath = path.join(mergedOutputDir, fileName);
    await writeDocx(
      filePath,
      `NotebookLM Upload Pack ${i + 1} (${yearLabel})`,
      pack.sermons
    );
    mergedStats.push({
      pack: i + 1,
      years: pack.years,
      sermons: pack.sermons.length,
      words: pack.words,
      file: filePath
    });
    console.log(`- pack ${i + 1}: ${pack.sermons.length} sermons, ${pack.words.toLocaleString()} words`);
  }

  const indexPath = path.join(OUTPUT_ROOT, "sermons_index.json");
  const summaryPath = path.join(OUTPUT_ROOT, "summary.json");
  await fs.writeFile(indexPath, JSON.stringify(usable, null, 2), "utf8");
  await fs.writeFile(
    summaryPath,
    JSON.stringify(
      {
        range: { start: START_DATE, end: END_DATE },
        totalSermons: usable.length,
        notebooklmLimits: {
          assumedMaxWordsPerSource: 500000,
          appliedSafeLimitWordsPerSource: MAX_WORDS_PER_NOTEBOOKLM_SOURCE
        },
        yearly: yearlyStats.map((x) => ({
          ...x,
          file: path.relative(OUTPUT_ROOT, x.file)
        })),
        mergedForNotebookLm: mergedStats.map((x) => ({
          ...x,
          file: path.relative(OUTPUT_ROOT, x.file)
        }))
      },
      null,
      2
    ),
    "utf8"
  );

  console.log("");
  console.log("Done:");
  console.log(`- output root: ${OUTPUT_ROOT}`);
  console.log(`- yearly docx: ${yearlyOutputDir}`);
  console.log(`- merged docx: ${mergedOutputDir}`);
  console.log(`- summary: ${summaryPath}`);
}

main().catch((error) => {
  console.error(`Failed: ${error.message}`);
  process.exitCode = 1;
});












