import { promises as fs } from "node:fs";
import path from "node:path";
import { AlignmentType, Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";

const LIST_BASE = "http://pastorlee.fgtv.com/fgnews/pastorlee_list.asp";
const SOURCE_LIST_URL =
  "http://www.fgnews.co.kr/front/sub/list.do?first_category_id=1&second_category_id=8";
const OUTPUT_DIR = path.resolve(process.cwd(), "output", "absolute_positive_gratitude");
const OUTPUT_DOCX = path.resolve(
  process.cwd(),
  "이영훈목사_절대긍정_절대감사_설교모음_20260411.docx"
);
const OUTPUT_JSON = path.join(OUTPUT_DIR, "matched_sermons.json");
const OUTPUT_TXT = path.join(OUTPUT_DIR, "matched_sermons.txt");
const OUTPUT_SUMMARY = path.join(OUTPUT_DIR, "summary.json");

const USER_AGENT =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";

const KEYWORDS = [
  { key: "절대 긍정", regex: /절대\s*긍정/g },
  { key: "절대 감사", regex: /절대\s*감사/g }
];

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchText(url, { retries = 4, timeoutMs = 30000, waitMs = 250 } = {}) {
  let lastError;
  for (let i = 0; i < retries; i += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const res = await fetch(url, {
        headers: { "user-agent": USER_AGENT },
        signal: controller.signal
      });
      if (!res.ok) {
        throw new Error(`HTTP ${res.status} for ${url}`);
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
  throw new Error(`Failed to fetch ${url}: ${lastError?.message ?? "unknown error"}`);
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
    ndash: "-",
    sim: "∼",
    middot: "·"
  };

  return input
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, code) => String.fromCodePoint(Number.parseInt(code, 16)))
    .replace(/&([a-zA-Z]+);/g, (full, key) => named[key] ?? full);
}

function normalizeLine(line) {
  return line.replace(/\s+/g, " ").trim();
}

function stripTagsToText(html) {
  let out = html ?? "";
  out = out.replace(/<script[\s\S]*?<\/script>/gi, "");
  out = out.replace(/<style[\s\S]*?<\/style>/gi, "");
  out = out.replace(/<br\s*\/?>/gi, "\n");
  out = out.replace(/<\/(p|div|h1|h2|h3|h4|h5|h6|li|tr|td|table|blockquote|ul|ol)>/gi, "\n");
  out = out.replace(/<(p|div|h1|h2|h3|h4|h5|h6|li|tr|td|table|blockquote|ul|ol)[^>]*>/gi, "\n");
  out = out.replace(/<[^>]+>/g, "");
  out = decodeHtmlEntities(out);
  out = out.replace(/\u00a0/g, " ");
  out = out.replace(/\r/g, "");
  return out;
}

function cleanInlineText(text) {
  return normalizeLine(stripTagsToText(text));
}

function toDateIso(dotDate) {
  const m = String(dotDate).match(/^(\d{4})\.(\d{2})\.(\d{2})\.$/);
  if (!m) return null;
  return `${m[1]}-${m[2]}-${m[3]}`;
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

function normalizeForSearch(text) {
  return decodeHtmlEntities(text).replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function countMatches(text, regex) {
  regex.lastIndex = 0;
  const matches = text.match(regex);
  return matches ? matches.length : 0;
}

function matchesKeywords(text) {
  const counts = {};
  let total = 0;
  for (const { key, regex } of KEYWORDS) {
    const c = countMatches(text, regex);
    counts[key] = c;
    total += c;
  }
  return { total, counts, matched: total > 0 };
}

function htmlToLinesPreserveBlocks(contentHtml) {
  const text = stripTagsToText(contentHtml ?? "");
  const lines = [];
  let blankStreak = 0;

  for (const raw of text.split("\n")) {
    const cleaned = raw.replace(/\s+/g, " ").trim();
    if (!cleaned) {
      blankStreak += 1;
      if (blankStreak <= 1) lines.push("");
      continue;
    }
    blankStreak = 0;
    lines.push(cleaned);
  }

  while (lines.length > 0 && lines[0] === "") lines.shift();
  while (lines.length > 0 && lines[lines.length - 1] === "") lines.pop();
  return lines;
}

function parseDetail(detailHtml, fallbackTitle) {
  const titleMatch = detailHtml.match(/<div class="view_title[^"]*">\s*([\s\S]*?)\s*<\/div>/i);
  const title = cleanInlineText(titleMatch?.[1] ?? fallbackTitle ?? "");

  let contentHtml = extractBetween(detailHtml, '<div class="cls-contents">', '<textarea id="tts_area"');
  const ttsMatch = detailHtml.match(/<textarea id="tts_area"[^>]*>([\s\S]*?)<\/textarea>/i);
  const ttsText = decodeHtmlEntities(ttsMatch?.[1] ?? "").trim();

  if (!contentHtml.trim() && ttsText) {
    contentHtml = `<p>${ttsText}</p>`;
  }

  contentHtml = contentHtml.replace(/<\/div>\s*$/i, "").trim();
  const searchable = normalizeForSearch(`${stripTagsToText(contentHtml)}\n${ttsText}`);
  const lines = htmlToLinesPreserveBlocks(contentHtml);
  const keywordInfo = matchesKeywords(searchable);

  return {
    title,
    contentHtml,
    lines,
    searchable,
    keywordInfo
  };
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

function buildTextExport(matched) {
  const lines = [];
  lines.push("이영훈 목사 설교 중 절대 긍정/절대 감사 관련 설교 모음");
  lines.push(`생성일시: ${new Date().toISOString()}`);
  lines.push(`출처 목록: ${SOURCE_LIST_URL}`);
  lines.push(`전체 수집 설교: ${matched.totalCollected}`);
  lines.push(`매칭 설교 수: ${matched.items.length}`);
  lines.push("");

  matched.items.forEach((item, idx) => {
    lines.push("=".repeat(80));
    lines.push(`번호: ${idx + 1}`);
    lines.push(`제목: ${item.title}`);
    lines.push(`날짜: ${item.date}`);
    lines.push(`분류: ${item.category || "-"}`);
    lines.push(`본문: ${item.scripture || "-"}`);
    lines.push(`출처: ${item.textUrl}`);
    lines.push(`키워드 매칭: 절대 긍정(${item.keywordCounts["절대 긍정"]}), 절대 감사(${item.keywordCounts["절대 감사"]})`);
    lines.push("-".repeat(80));
    lines.push(...item.lines);
    lines.push("");
  });

  return `${lines.join("\n")}\n`;
}

function createBodyParagraph(line) {
  if (!line) return new Paragraph({ text: "" });

  const isSectionHeading = /^\d+\.\s+/.test(line);
  const isVerseReference = /^[가-힣]+복음|^[가-힣]{1,4}\s*\d+장\s*\d+절/.test(line);
  const run = new TextRun({
    text: line,
    bold: isSectionHeading,
    italics: !isSectionHeading && isVerseReference
  });

  return new Paragraph({
    children: [run],
    spacing: {
      before: isSectionHeading ? 220 : 60,
      after: isSectionHeading ? 120 : 60
    }
  });
}

async function writeDocx(filePath, sermons, totalCollected) {
  const children = [];

  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "이영훈 목사 설교: 절대 긍정/절대 감사 포함 설교 모음", bold: true })]
    })
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun(`생성일: ${new Date().toISOString().slice(0, 10)}`)]
    })
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun(`원본 카테고리 전체 설교 수집 수: ${totalCollected}`)]
    })
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun(`키워드 매칭 설교 수: ${sermons.length}`)]
    })
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun(`출처 목록: ${SOURCE_LIST_URL}`)]
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
        children: [new TextRun(`${idx + 1}. ${sermon.title}`)]
      })
    );
    children.push(new Paragraph({ children: [new TextRun(`날짜: ${sermon.date}`)] }));
    children.push(new Paragraph({ children: [new TextRun(`분류: ${sermon.category || "-"}`)] }));
    children.push(new Paragraph({ children: [new TextRun(`본문: ${sermon.scripture || "-"}`)] }));
    children.push(new Paragraph({ children: [new TextRun(`출처: ${sermon.textUrl}`)] }));
    children.push(
      new Paragraph({
        children: [
          new TextRun(
            `키워드 매칭: 절대 긍정(${sermon.keywordCounts["절대 긍정"]}), 절대 감사(${sermon.keywordCounts["절대 감사"]})`
          )
        ]
      })
    );
    children.push(new Paragraph({ text: "" }));

    for (const line of sermon.lines) {
      children.push(createBodyParagraph(line));
    }
  });

  const doc = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(filePath, buffer);
}

async function main() {
  await fs.mkdir(OUTPUT_DIR, { recursive: true });

  console.log("[1/6] 목록 페이지 수집 시작");
  const firstHtml = await fetchText(LIST_BASE);
  const maxPage = parseMaxPage(firstHtml);
  console.log(`- 목록 페이지 수: ${maxPage}`);

  const allRows = [];
  for (let page = 1; page <= maxPage; page += 1) {
    const url = page === 1 ? LIST_BASE : `${LIST_BASE}?page=${page}`;
    const html = page === 1 ? firstHtml : await fetchText(url);
    const rows = parseListRows(html);
    allRows.push(...rows);
    if (page % 10 === 0 || page === maxPage) {
      console.log(`- 목록 진행: ${page}/${maxPage}`);
    }
    await sleep(40);
  }

  const rowMap = new Map();
  for (const row of allRows) {
    const key = `${row.date}|${row.textUrl}`;
    if (!rowMap.has(key)) rowMap.set(key, row);
  }
  const dedupedRows = [...rowMap.values()];
  console.log(`[2/6] 중복 제거 후 설교 수: ${dedupedRows.length}`);

  console.log("[3/6] 설교 본문 수집");
  const detailed = await mapWithConcurrency(dedupedRows, 5, async (row, idx) => {
    const html = await fetchText(row.textUrl);
    const detail = parseDetail(html, row.title);
    if ((idx + 1) % 20 === 0 || idx + 1 === dedupedRows.length) {
      console.log(`- 본문 진행: ${idx + 1}/${dedupedRows.length}`);
    }
    await sleep(25);
    return {
      ...row,
      title: detail.title || row.title,
      lines: detail.lines,
      searchable: detail.searchable,
      keywordInfo: detail.keywordInfo
    };
  });

  console.log("[4/6] 키워드 필터링");
  const matched = detailed.filter((item) => item.keywordInfo.matched && item.lines.length > 0);
  console.log(`- 매칭 설교 수: ${matched.length}`);

  const exportItems = matched.map((item) => ({
    title: item.title,
    date: item.date,
    category: item.category,
    scripture: item.scripture,
    textUrl: item.textUrl,
    keywordCounts: item.keywordInfo.counts,
    lines: item.lines
  }));

  await fs.writeFile(
    OUTPUT_JSON,
    JSON.stringify(
      {
        sourceListUrl: SOURCE_LIST_URL,
        collectedAt: new Date().toISOString(),
        totalCollected: dedupedRows.length,
        totalMatched: exportItems.length,
        items: exportItems
      },
      null,
      2
    ),
    "utf8"
  );
  await fs.writeFile(
    OUTPUT_TXT,
    buildTextExport({
      totalCollected: dedupedRows.length,
      items: exportItems
    }),
    "utf8"
  );

  console.log("[5/6] DOCX 생성");
  await writeDocx(OUTPUT_DOCX, exportItems, dedupedRows.length);

  await fs.writeFile(
    OUTPUT_SUMMARY,
    JSON.stringify(
      {
        sourceListUrl: SOURCE_LIST_URL,
        listBase: LIST_BASE,
        collectedAt: new Date().toISOString(),
        totalPages: maxPage,
        totalCollected: dedupedRows.length,
        totalMatched: exportItems.length,
        output: {
          docx: path.relative(process.cwd(), OUTPUT_DOCX),
          json: path.relative(process.cwd(), OUTPUT_JSON),
          txt: path.relative(process.cwd(), OUTPUT_TXT)
        }
      },
      null,
      2
    ),
    "utf8"
  );

  console.log("[6/6] 완료");
  console.log(`- DOCX: ${OUTPUT_DOCX}`);
  console.log(`- JSON: ${OUTPUT_JSON}`);
  console.log(`- TXT: ${OUTPUT_TXT}`);
  console.log(`- 매칭/전체: ${exportItems.length}/${dedupedRows.length}`);
}

main().catch((error) => {
  console.error(`Failed: ${error.message}`);
  process.exitCode = 1;
});
