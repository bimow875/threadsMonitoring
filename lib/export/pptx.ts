import JSZip from "jszip";
import { MentionStatus, Sentiment, Priority, InputSourceType } from "@prisma/client";

export interface MentionExportData {
  id: string;
  tanggalDitemukan: Date | string;
  topik: string | null;
  sentimen: Sentiment | string;
  prioritas: Priority | string;
  status: MentionStatus | string;
  sumberInput: InputSourceType | string;
  headline: string | null;
  ringkasan: string | null;
  teksPostingan: string;
  linkAsli: string;
  linkTernormalisasi?: string | null;
}

export interface PptxExportOptions {
  startDate?: string | null;
  endDate?: string | null;
  topik?: string | null;
  sentimen?: string | null;
  includeSample?: boolean;
}

/**
 * Membersihkan string dari karakter kontrol ilegal XML dan melakukan HTML/XML entity encoding.
 */
export function escapeXml(unsafe: unknown): string {
  if (unsafe === null || unsafe === undefined) return "";
  return String(unsafe)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function getSentimentColor(sentimen: string): { bg: string; text: string } {
  const s = (sentimen || "").toUpperCase();
  if (s === "POSITIF") return { bg: "EBF3FF", text: "0055FF" };
  if (s === "NEGATIF") return { bg: "FEF2F2", text: "C10010" };
  return { bg: "F1F5F9", text: "475569" };
}

function getPriorityColor(prioritas: string): { bg: string; text: string } {
  const p = (prioritas || "").toUpperCase();
  if (p === "PERLU_PERHATIAN") return { bg: "FEE2E2", text: "E60012" };
  if (p === "PANTAU") return { bg: "FEF3C7", text: "D97706" };
  return { bg: "F8FAFC", text: "334155" };
}

/**
 * Menghasilkan presentasi PowerPoint (.pptx) terstruktur untuk Bank BTN Social Listening.
 */
export async function generatePptxReport(
  mentions: MentionExportData[],
  options: PptxExportOptions
): Promise<Buffer> {
  const zip = new JSZip();

  const totalCount = mentions.length;
  const positifCount = mentions.filter((m) => m.sentimen === "POSITIF").length;
  const netralCount = mentions.filter((m) => m.sentimen === "NETRAL").length;
  const negatifCount = mentions.filter((m) => m.sentimen === "NEGATIF").length;
  const urgentCount = mentions.filter((m) => m.prioritas === "PERLU_PERHATIAN").length;
  const contohCount = mentions.filter((m) => m.sumberInput === "CONTOH").length;

  const pctPos = totalCount > 0 ? ((positifCount / totalCount) * 100).toFixed(1) : "0.0";
  const pctNet = totalCount > 0 ? ((netralCount / totalCount) * 100).toFixed(1) : "0.0";
  const pctNeg = totalCount > 0 ? ((negatifCount / totalCount) * 100).toFixed(1) : "0.0";
  const pctUrg = totalCount > 0 ? ((urgentCount / totalCount) * 100).toFixed(1) : "0.0";

  // Topik distribusi
  const topicCounts: Record<string, number> = {};
  mentions.forEach((m) => {
    const t = m.topik || "Lainnya";
    topicCounts[t] = (topicCounts[t] || 0) + 1;
  });
  const sortedTopics = Object.entries(topicCounts).sort((a, b) => b[1] - a[1]);

  const dateStr =
    options.startDate || options.endDate
      ? `${options.startDate || "Awal"} s/d ${options.endDate || "Sekarang"}`
      : "Semua Rentang Waktu";

  const exportTimestamp = new Date().toLocaleString("id-ID", {
    timeZone: "Asia/Jakarta",
    dateStyle: "full",
    timeStyle: "short",
  });

  const slides: string[] = [];

  // ==========================================
  // SLIDE 1: COVER SLIDE
  // ==========================================
  slides.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>

      <!-- Pill Badge: PT BANK TABUNGAN NEGARA -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="BadgeBTN"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1000000" y="1100000"/><a:ext cx="3600000" cy="450000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="0055FF"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>
              <a:t>PT BANK TABUNGAN NEGARA (PERSERO) TBK</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- Main Title & Subtitle -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="MainTitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1000000" y="1800000"/><a:ext cx="10192000" cy="1700000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="id-ID" sz="3400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>
              <a:t>Laporan Monitoring Media Sosial Threads</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="100000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1800"><a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill></a:rPr>
              <a:t>Social Listening &amp; Analytics Dashboard</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- Parameters Card Box -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="InfoCard"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1000000" y="3800000"/><a:ext cx="10192000" cy="2250000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="1E293B"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="200000" tIns="160000" rIns="200000" bIns="160000"/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="id-ID" sz="1100" b="1"><a:solidFill><a:srgbClr val="0055FF"/></a:solidFill></a:rPr>
              <a:t>PARAMETER LAPORAN</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="80000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1300"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>
              <a:t>• Rentang Tanggal : ${escapeXml(dateStr)}</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="60000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1300"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill></a:rPr>
              <a:t>• Filter Topik : ${escapeXml(options.topik || "Semua Topik")}    |    Filter Sentimen : ${escapeXml(options.sentimen || "Semua")}</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="60000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1300"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill></a:rPr>
              <a:t>• Data Contoh Simulasi : ${options.includeSample ? `Disertakan (${contohCount} item simulasi)` : "Dikecualikan (Hanya Data Riil)"}</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="60000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1150"><a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill></a:rPr>
              <a:t>• Tanggal Pembuatan : ${escapeXml(exportTimestamp)} WIB   (Draf selalu dikecualikan dari ekspor resmi)</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`);

  // ==========================================
  // SLIDE 2: RINGKASAN EKSEKUTIF & METRIK
  // ==========================================
  const topicListText =
    sortedTopics.length > 0
      ? sortedTopics
          .slice(0, 5)
          .map(([t, count], i) => `${i + 1}. ${escapeXml(t)}: ${count} mention (${totalCount > 0 ? ((count / totalCount) * 100).toFixed(0) : 0}%)`)
          .join("   •   ")
      : "Tidak ada topik tercatat.";

  slides.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr><a:solidFill><a:srgbClr val="F4F7FB"/></a:solidFill></p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>

      <!-- Slide Header -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="H1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="609600" y="450000"/><a:ext cx="10972800" cy="800000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="id-ID" sz="2400" b="1"><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:rPr>
              <a:t>Ringkasan Eksekutif &amp; Statistik Sentimen</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="30000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1200"><a:solidFill><a:srgbClr val="5A6A7E"/></a:solidFill></a:rPr>
              <a:t>Statistik agregat percakapan publik pada Threads untuk periode yang dipilih</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- KPI 1: Total Mention -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="KpiTotal"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="609600" y="1450000"/><a:ext cx="2050000" cy="1550000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="91440" tIns="91440" rIns="91440" bIns="91440" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="id-ID" sz="3200" b="1"><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:rPr>
              <a:t>${totalCount}</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr" spaceBefore="50000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="5A6A7E"/></a:solidFill></a:rPr>
              <a:t>TOTAL MENTION</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- KPI 2: Sentimen Positif -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="KpiPos"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="2840000" y="1450000"/><a:ext cx="2050000" cy="1550000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="91440" tIns="91440" rIns="91440" bIns="91440" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="id-ID" sz="3200" b="1"><a:solidFill><a:srgbClr val="0055FF"/></a:solidFill></a:rPr>
              <a:t>${positifCount}</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr" spaceBefore="50000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="0055FF"/></a:solidFill></a:rPr>
              <a:t>POSITIF (${pctPos}%)</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- KPI 3: Sentimen Netral -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="5" name="KpiNet"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="5070000" y="1450000"/><a:ext cx="2050000" cy="1550000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="91440" tIns="91440" rIns="91440" bIns="91440" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="id-ID" sz="3200" b="1"><a:solidFill><a:srgbClr val="475569"/></a:solidFill></a:rPr>
              <a:t>${netralCount}</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr" spaceBefore="50000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="475569"/></a:solidFill></a:rPr>
              <a:t>NETRAL (${pctNet}%)</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- KPI 4: Sentimen Negatif -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="6" name="KpiNeg"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="7300000" y="1450000"/><a:ext cx="2050000" cy="1550000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="91440" tIns="91440" rIns="91440" bIns="91440" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="id-ID" sz="3200" b="1"><a:solidFill><a:srgbClr val="C10010"/></a:solidFill></a:rPr>
              <a:t>${negatifCount}</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr" spaceBefore="50000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="C10010"/></a:solidFill></a:rPr>
              <a:t>NEGATIF (${pctNeg}%)</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- KPI 5: Perlu Perhatian -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="7" name="KpiUrg"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="9530000" y="1450000"/><a:ext cx="2050000" cy="1550000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="91440" tIns="91440" rIns="91440" bIns="91440" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="id-ID" sz="3200" b="1"><a:solidFill><a:srgbClr val="E60012"/></a:solidFill></a:rPr>
              <a:t>${urgentCount}</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr" spaceBefore="50000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="E60012"/></a:solidFill></a:rPr>
              <a:t>PERLU PERHATIAN (${pctUrg}%)</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- Panel Catatan Analisis & Topik -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="8" name="AnalysisPanel"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="609600" y="3250000"/><a:ext cx="10970400" cy="2900000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="200000" tIns="180000" rIns="200000" bIns="180000"/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="id-ID" sz="1400" b="1"><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:rPr>
              <a:t>Ikhtisar Tren &amp; Distribusi Topik Utama</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="100000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1200"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr>
              <a:t>• Topik Utama : ${topicListText}</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="80000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1200"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr>
              <a:t>• Sentimen Positif tercatat sebesar ${pctPos}% (${positifCount} mention), Netral sebesar ${pctNet}% (${netralCount} mention), dan Negatif sebesar ${pctNeg}% (${negatifCount} mention).</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="80000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1200"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr>
              <a:t>• Terdapat ${urgentCount} mention (${pctUrg}%) yang memerlukan perhatian prioritas penanganan atau eskalasi ke divisi operasional/pelayanan.</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="80000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1150"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill></a:rPr>
              <a:t>• Seluruh data postingan yang memuat informasi pribadi nasabah (nomor kartu, nomor rekening, nomor ponsel) telah otomatis disamarkan (PII masking) sesuai standar privasi data perbankan.</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`);

  // ==========================================
  // SLIDE 3+: TABEL DETAIL MENTIONS (PAGINATED 4 ITEM PER SLIDE)
  // ==========================================
  const ITEMS_PER_SLIDE = 4;
  const totalPages = Math.max(1, Math.ceil(mentions.length / ITEMS_PER_SLIDE));

  if (mentions.length === 0) {
    // Slide kosong jika 0 mention
    slides.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F4F7FB"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="EmptyState"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="2000000" y="2500000"/><a:ext cx="8192000" cy="1800000"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="100000" tIns="100000" rIns="100000" bIns="100000" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1800" b="1"><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:rPr>
              <a:t>Tidak Ada Data Mention</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr" spaceBefore="60000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1300"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill></a:rPr>
              <a:t>Tidak ada postingan yang sesuai dengan filter atau periode yang dipilih.</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`);
  } else {
    for (let page = 1; page <= totalPages; page++) {
      const startIndex = (page - 1) * ITEMS_PER_SLIDE;
      const pageMentions = mentions.slice(startIndex, startIndex + ITEMS_PER_SLIDE);

      let tableRows = `
              <!-- Header Row -->
              <a:tr h="450000">
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>TANGGAL</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>TOPIK</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>SENTIMEN</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>PRIORITAS</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:r><a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>HEADLINE &amp; RINGKASAN</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>SUMBER</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:tcPr>
                </a:tc>
              </a:tr>`;

      pageMentions.forEach((m) => {
        const tglStr = new Date(m.tanggalDitemukan).toLocaleDateString("id-ID", {
          day: "2-digit",
          month: "2-digit",
          year: "numeric",
        });

        const sentCol = getSentimentColor(m.sentimen);
        const priCol = getPriorityColor(m.prioritas);

        const sumberLabel =
          m.sumberInput === "CONTOH" ? "CONTOH" : m.sumberInput || "MANUAL";
        const sumberBg = m.sumberInput === "CONTOH" ? "FEF3C7" : "FFFFFF";
        const sumberTextCol = m.sumberInput === "CONTOH" ? "92400E" : "475569";

        const headlineText = escapeXml(m.headline || "Mention Baru");
        const ringkasanText = escapeXml(
          m.ringkasan || (m.teksPostingan ? m.teksPostingan.slice(0, 160) : "")
        );

        tableRows += `
              <!-- Data Row -->
              <a:tr h="850000">
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="900"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr><a:t>${escapeXml(tglStr)}</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:lnB w="9525"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:lnB></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="950" b="1"><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:rPr><a:t>${escapeXml(m.topik || "Lainnya")}</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:lnB w="9525"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:lnB></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="950" b="1"><a:solidFill><a:srgbClr val="${sentCol.text}"/></a:solidFill></a:rPr><a:t>${escapeXml(m.sentimen)}</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="${sentCol.bg}"/></a:solidFill><a:lnB w="9525"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:lnB></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="900" b="1"><a:solidFill><a:srgbClr val="${priCol.text}"/></a:solidFill></a:rPr><a:t>${escapeXml(m.prioritas)}</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="${priCol.bg}"/></a:solidFill><a:lnB w="9525"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:lnB></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>
                    <a:p><a:r><a:rPr lang="id-ID" sz="950" b="1"><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:rPr><a:t>${headlineText}</a:t></a:r></a:p>
                    <a:p><a:pPr spaceBefore="30000"/><a:r><a:rPr lang="id-ID" sz="850"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill></a:rPr><a:t>${ringkasanText}</a:t></a:r></a:p>
                  </a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:lnB w="9525"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:lnB></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="id-ID" sz="850" b="1"><a:solidFill><a:srgbClr val="${sumberTextCol}"/></a:solidFill></a:rPr><a:t>${escapeXml(sumberLabel)}</a:t></a:r></a:p></a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="${sumberBg}"/></a:solidFill><a:lnB w="9525"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:lnB></a:tcPr>
                </a:tc>
              </a:tr>`;
      });

      slides.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F4F7FB"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>

      <!-- Slide Title & Page Indicator -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TitlePage"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="609600" y="400000"/><a:ext cx="10972800" cy="650000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="id-ID" sz="2200" b="1"><a:solidFill><a:srgbClr val="0A192F"/></a:solidFill></a:rPr>
              <a:t>Daftar Mention Teridentifikasi</a:t>
            </a:r>
            <a:r>
              <a:rPr lang="id-ID" sz="1400"><a:solidFill><a:srgbClr val="5A6A7E"/></a:solidFill></a:rPr>
              <a:t>  (Halaman ${page} dari ${totalPages})</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr spaceBefore="20000"/>
            <a:r>
              <a:rPr lang="id-ID" sz="1100"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill></a:rPr>
              <a:t>Menampilkan mention ${startIndex + 1} s/d ${Math.min(startIndex + ITEMS_PER_SLIDE, totalCount)} dari total ${totalCount} data</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- Graphic Frame Table -->
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="3" name="MentionsTable${page}"/>
          <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="609600" y="1150000"/>
          <a:ext cx="10972800" cy="5200000"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr firstRow="1"/>
              <a:tblGrid>
                <a:gridCol w="1100000"/>
                <a:gridCol w="1600000"/>
                <a:gridCol w="1200000"/>
                <a:gridCol w="1500000"/>
                <a:gridCol w="4172800"/>
                <a:gridCol w="1400000"/>
              </a:tblGrid>
${tableRows}
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>

    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`);
    }
  }

  // ==========================================
  // PACKAGE ASSEMBLY (ECMA-376 / OPENXML)
  // ==========================================
  // 1. [Content_Types].xml
  let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n`;

  slides.forEach((_, idx) => {
    contentTypes += `  <Override PartName="/ppt/slides/slide${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n`;
  });
  contentTypes += `</Types>`;
  zip.file("[Content_Types].xml", contentTypes);

  // 2. _rels/.rels
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
  );

  // 3. docProps/core.xml
  zip.file(
    "docProps/core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Laporan Threads Monitoring - Bank BTN</dc:title>
  <dc:creator>BTN Threads Monitoring</dc:creator>
  <cp:lastModifiedBy>BTN Threads Monitoring</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>
</cp:coreProperties>`
  );

  // 4. docProps/app.xml
  zip.file(
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Application>BTN Threads Monitoring</Application>
</Properties>`
  );

  // 5. ppt/_rels/presentation.xml.rels
  let presRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>\n`;

  slides.forEach((_, idx) => {
    presRels += `  <Relationship Id="rId${idx + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${idx + 1}.xml"/>\n`;
  });
  const themeRelId = `rId${slides.length + 2}`;
  presRels += `  <Relationship Id="${themeRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>\n`;
  presRels += `</Relationships>`;
  zip.file("ppt/_rels/presentation.xml.rels", presRels);

  // 6. ppt/presentation.xml
  let sldIds = "";
  slides.forEach((_, idx) => {
    sldIds += `    <p:sldId id="${256 + idx}" r:id="rId${idx + 2}"/>\n`;
  });

  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
${sldIds}  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="12192000"/>
</p:presentation>`
  );

  // 7. Theme, master, layout
  zip.file(
    "ppt/theme/theme1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="0A192F"/></a:dk2>
      <a:lt2><a:srgbClr val="F4F7FB"/></a:lt2>
      <a:accent1><a:srgbClr val="0055FF"/></a:accent1>
      <a:accent2><a:srgbClr val="475569"/></a:accent2>
      <a:accent3><a:srgbClr val="047857"/></a:accent3>
      <a:accent4><a:srgbClr val="C10010"/></a:accent4>
      <a:accent5><a:srgbClr val="E60012"/></a:accent5>
      <a:accent6><a:srgbClr val="F59E0B"/></a:accent6>
      <a:hlink><a:srgbClr val="0055FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="0043CC"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`
  );

  zip.file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`
  );

  zip.file(
    "ppt/slideMasters/slideMaster1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`
  );

  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`
  );

  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`
  );

  // 8. Tulis slide dan slide rels
  slides.forEach((sldXml, idx) => {
    const slideNum = idx + 1;
    zip.file(
      `ppt/slides/_rels/slide${slideNum}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
    );
    zip.file(`ppt/slides/slide${slideNum}.xml`, sldXml);
  });

  return (await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  })) as Buffer;
}
