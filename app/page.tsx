"use client";
import { useMemo, useState } from "react";
import * as XLSX from "xlsx-js-style";
import { Download, FileSpreadsheet, Upload } from "lucide-react";
type R = Record<string, string>;
type A = { counselor: string; customer: string };
const keys = [
  "고객사명",
  "협력사명",
  "상담명",
  "상담자명",
  "상담일자",
  "상담구분",
  "우선순위",
  "상담내용",
];
const aliases: Record<string, string[]> = {
  고객사명: ["고객사명", "고객사", "상담기관"],
  협력사명: ["협력사명", "협력사"],
  상담명: ["상담명", "상담제목"],
  상담자명: ["상담자명", "상담자", "상담사", "담당자", "상담사(담당자)"],
  상담일자: ["상담일자", "상담일시"],
  상담구분: ["상담구분", "구분"],
  우선순위: ["우선순위"],
  상담내용: ["상담내용", "내용"],
};
const norm = (v: unknown) =>
  String(v ?? "")
    .trim()
    .replace(/\s+/g, " ");
const head = (v: unknown) => norm(v).replace(/\s/g, "").toLowerCase();
const formatContent = (v: string) =>
  v.replace(/\s*(?:->|→)\s*/g, "\n→ ").trim();
const kind = (v: string) => {
  const x = head(v);
  if (["일반", "일반상담"].includes(x)) return "일반상담";
  if (["클레임", "클레임상담"].includes(x)) return "클레임상담";
  return x ? v.trim() : "미분류";
};
function readSheet(wb: XLSX.WorkBook, sheetName: string) {
  const data = XLSX.utils.sheet_to_json<unknown[]>(
    wb.Sheets[sheetName],
    { header: 1, defval: "", raw: false },
  );
  const i = data.findIndex((r) => r.some((v) => norm(v)));
  if (i < 0) throw Error(`${sheetName} 시트에 데이터가 없습니다.`);
  return { headers: data[i].map(norm), data: data.slice(i + 1) };
}

const findSheet = (wb: XLSX.WorkBook, preferredName: string, fallbackIndex: number) =>
  wb.SheetNames.find((name) => head(name) === head(preferredName)) ??
  wb.SheetNames[fallbackIndex];

function readEnterprisePartners(wb: XLSX.WorkBook) {
  const preferred = wb.SheetNames.find((name) => head(name) === head("대기업협력사"));
  const sheetName = preferred ?? wb.SheetNames[2];
  if (!sheetName) return [];
  const data = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: false,
  });
  const headerRow = data.findIndex((row) =>
    row.some((value) => ["협력사명", "협력사"].some((alias) => head(value) === head(alias))),
  );
  if (headerRow < 0) throw Error("대기업협력사 시트에서 협력사명 열을 찾을 수 없습니다.");
  const partnerColumn = data[headerRow].findIndex((value) =>
    ["협력사명", "협력사"].some((alias) => head(value) === head(alias)),
  );
  return data
    .slice(headerRow + 1)
    .map((row) => norm(row[partnerColumn]))
    .filter(Boolean);
}
function source(p: ReturnType<typeof readSheet>) {
  const ix: Record<string, number> = {};
  keys.forEach((k) => {
    ix[k] = p.headers.findIndex((h) =>
      aliases[k].some((a) => head(a) === head(h)),
    );
    if (ix[k] < 0) throw Error(`필수 열을 찾을 수 없습니다: ${k}`);
  });
  return p.data
    .filter((r) => r.some((v) => norm(v)))
    .map(
      (r) =>
        Object.fromEntries(
          keys.map((k) => [
            k,
            k === "상담내용" ? formatContent(norm(r[ix[k]])) : norm(r[ix[k]]),
          ]),
        ) as R,
    );
}
function assign(p: ReturnType<typeof readSheet>) {
  const c = p.headers.findIndex((h) =>
      aliases.상담자명.some((a) => head(a) === head(h)),
    ),
    u = p.headers.findIndex((h) =>
      aliases.고객사명.some((a) => head(a) === head(h)),
    );
  if (c < 0 || u < 0)
    throw Error("담당 고객사 파일에는 상담사(담당자), 고객사 열이 필요합니다.");
  return p.data
    .map((r) => ({ counselor: norm(r[c]), customer: norm(r[u]) }))
    .filter((x) => x.counselor && x.customer);
}
const border = {
  top: { style: "thin", color: { rgb: "111827" } },
  bottom: { style: "thin", color: { rgb: "111827" } },
  left: { style: "thin", color: { rgb: "111827" } },
  right: { style: "thin", color: { rgb: "111827" } },
};
function sheet(
  title: string,
  headers: string[],
  rows: string[][],
  color = "F7C85E",
) {
  const ws = XLSX.utils.aoa_to_sheet([[title], [], headers, ...rows]);
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }];
  ws["!autofilter"] = {
    ref: `A3:${XLSX.utils.encode_col(headers.length - 1)}${rows.length + 3}`,
  };
  ws["!cols"] = headers.map((h) => ({
    wch: h === "상담내용" ? 70 : h.includes("명") ? 24 : 14,
  }));
  headers.forEach((_, c) => {
    const t = ws[XLSX.utils.encode_cell({ r: 0, c })];
    if (t)
      t.s = {
        font: { bold: true, sz: 16 },
        alignment: { horizontal: "center" },
      };
    const h = ws[XLSX.utils.encode_cell({ r: 2, c })];
    if (h)
      h.s = {
        fill: { fgColor: { rgb: color } },
        font: { bold: true },
        border,
        alignment: { horizontal: "center" },
      };
  });
  rows.forEach((_, ri) =>
    headers.forEach((__, c) => {
      const x = ws[XLSX.utils.encode_cell({ r: ri + 3, c })];
      if (x) x.s = { border, alignment: { vertical: "top", wrapText: true } };
    }),
  );
  const contentIndex = headers.indexOf("상담내용");
  ws["!rows"] = [
    { hpt: 28 },
    { hpt: 8 },
    { hpt: 22 },
    ...rows.map((r) => {
      const v = contentIndex >= 0 ? (r[contentIndex] ?? "") : "";
      const lines = Math.max(
        1,
        v
          .split("\n")
          .reduce((n, line) => n + Math.max(1, Math.ceil(line.length / 65)), 0),
      );
      return { hpt: Math.min(180, Math.max(42, lines * 18)) };
    }),
  ];
  return ws;
}

function mainSummarySheet(title: string, left: string[][], complaints: string[][]) {
  const rowCount = Math.max(left.length, complaints.length + 1);
  const complaintTotal = complaints.reduce((sum, row) => sum + Number(row[1] || 0), 0);
  const body = Array.from({ length: rowCount }, (_, index) => {
    const leftRow = left[index] ?? ["", ""];
    const rightRow = index < complaints.length
      ? complaints[index]
      : index === complaints.length
        ? ["총합계", String(complaintTotal)]
        : ["", ""];
    return [...leftRow, "", ...rightRow];
  });
  const ws = XLSX.utils.aoa_to_sheet([
    [title],
    [],
    ["TOTAL 상담", "", "", "고객사별 클레임", ""],
    ["상담자 / 상담기관", "상담 수", "", "고객사", "상담 수"],
    ...body,
  ]);
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } },
    { s: { r: 2, c: 3 }, e: { r: 2, c: 4 } },
  ];
  ws["!cols"] = [{ wch: 34 }, { wch: 10 }, { wch: 5 }, { wch: 34 }, { wch: 10 }];
  ws["!rows"] = [{ hpt: 30 }, { hpt: 10 }, { hpt: 22 }, { hpt: 22 }, ...body.map(() => ({ hpt: 20 }))];
  const titleCell = ws.A1;
  if (titleCell) titleCell.s = { font: { bold: true, sz: 16 }, alignment: { horizontal: "center" } };
  ["A3", "D3"].forEach((ref) => {
    if (ws[ref]) ws[ref].s = { font: { bold: true }, alignment: { horizontal: "center" } };
  });
  [0, 1].forEach((c) => {
    const cell = ws[XLSX.utils.encode_cell({ r: 3, c })];
    if (cell) cell.s = { fill: { fgColor: { rgb: "CDE3A4" } }, font: { bold: true }, border, alignment: { horizontal: "center" } };
  });
  [3, 4].forEach((c) => {
    const cell = ws[XLSX.utils.encode_cell({ r: 3, c })];
    if (cell) cell.s = { fill: { fgColor: { rgb: "D89B9B" } }, font: { bold: true }, border, alignment: { horizontal: "center" } };
  });
  body.forEach((row, index) => {
    const excelRow = index + 4;
    [0, 1, 3, 4].forEach((c) => {
      const cell = ws[XLSX.utils.encode_cell({ r: excelRow, c })];
      if (cell) cell.s = { border, alignment: { vertical: "center", horizontal: c === 1 || c === 4 ? "center" : "left" } };
    });
    if (row[0] && !row[0].startsWith("  ")) {
      [0, 1].forEach((c) => {
        const cell = ws[XLSX.utils.encode_cell({ r: excelRow, c })];
        if (cell) cell.s = { ...cell.s, fill: { fgColor: { rgb: "E2EFCB" } }, font: { bold: true } };
      });
    }
    if (row[3] === "총합계") {
      [3, 4].forEach((c) => {
        const cell = ws[XLSX.utils.encode_cell({ r: excelRow, c })];
        if (cell) cell.s = { ...cell.s, fill: { fgColor: { rgb: "E7C1C1" } }, font: { bold: true } };
      });
    }
  });
  return ws;
}

function partnerSummarySheet(
  title: string,
  market: [string, R[]][],
  enterprise: [string, R[]][],
  summaries: Record<string, string>,
) {
  const section = (label: string, entries: [string, R[]][]) => [
    [label, "건수", "요약"],
    ...entries.map(([name, rows]) => [name, `${rows.length}건`, summaries[name] ?? ""]),
  ];
  const marketRows = section("협력사명 (시장)", market);
  const enterpriseRows = section("협력사명 (대기업)", enterprise);
  const enterpriseHeaderRow = marketRows.length + 3;
  const values = [[title], [], ...marketRows, [], ...enterpriseRows];
  const ws = XLSX.utils.aoa_to_sheet(values);
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
  ws["!cols"] = [{ wch: 38 }, { wch: 10 }, { wch: 72 }];
  ws["!rows"] = values.map((_, index) => ({ hpt: index === 0 ? 30 : index === 1 || index === marketRows.length + 2 ? 9 : 24 }));
  if (ws.A1) ws.A1.s = { font: { bold: true, sz: 16 }, alignment: { horizontal: "center" } };
  [2, enterpriseHeaderRow].forEach((row) => {
    [0, 1, 2].forEach((column) => {
      const cell = ws[XLSX.utils.encode_cell({ r: row, c: column })];
      if (cell) cell.s = { fill: { fgColor: { rgb: "D8CCE9" } }, font: { bold: true }, border, alignment: { horizontal: "center" } };
    });
  });
  values.forEach((_, row) => {
    if (row <= 2 || row === enterpriseHeaderRow || row === marketRows.length + 2) return;
    [0, 1, 2].forEach((column) => {
      const cell = ws[XLSX.utils.encode_cell({ r: row, c: column })];
      if (cell) cell.s = { border, alignment: { vertical: "top", horizontal: column === 1 ? "center" : "left", wrapText: true } };
    });
  });
  return ws;
}
export default function Home() {
  const [file, setFile] = useState<File | null>(null),
    [rows, setRows] = useState<R[]>([]),
    [as, setAs] = useState<A[]>([]),
    [enterprisePartners, setEnterprisePartners] = useState<string[]>([]),
    [msg, setMsg] = useState(""),
    [sum, setSum] = useState<Record<string, string>>({});
  const partners = useMemo(() => {
    const m = new Map<string, R[]>();
    rows
      .filter((r) => r.협력사명)
      .forEach((r) => m.set(r.협력사명, [...(m.get(r.협력사명) ?? []), r]));
    return [...m].sort(
      (a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0], "ko"),
    );
  }, [rows]);
  async function analyze() {
    if (!file) return setMsg("통합 XLSX 파일을 선택해주세요.");
    try {
      const wb = XLSX.read(await file.arrayBuffer(), {
          type: "array",
          cellDates: true,
        }),
        consultationSheet = findSheet(wb, "상담내역", 0),
        assignmentSheet = findSheet(wb, "담당고객사", 1);
      if (!consultationSheet)
        throw Error("상담내역 시트를 찾을 수 없습니다.");
      if (!assignmentSheet)
        throw Error("담당고객사 시트를 찾을 수 없습니다.");
      const r1 = readSheet(wb, consultationSheet),
        r2 = readSheet(wb, assignmentSheet),
        enterprise = readEnterprisePartners(wb),
        s = source(r1),
        a = assign(r2),
        dups = a.map((x) => x.customer).filter((x, i, z) => z.indexOf(x) !== i);
      if (dups.length)
        throw Error(`중복 고객사: ${[...new Set(dups)].join(", ")}`);
      setRows(s);
      setAs(a);
      setEnterprisePartners(enterprise);
      const u = [
        ...new Set(
          s
            .map((r) => kind(r.상담구분))
            .filter((x) => !["일반상담", "클레임상담"].includes(x)),
        ),
      ];
      setMsg(
        u.length
          ? `확인할 상담구분: ${u.join(", ")}`
          : `${s.length}건을 확인했습니다. 대기업 협력사 ${enterprise.length}개를 불러왔습니다.`,
      );
    } catch (e) {
      setRows([]);
      setMsg(e instanceof Error ? e.message : "파일을 읽지 못했습니다.");
    }
  }
  function download() {
    const d = rows.find((r) => r.상담일자)?.상담일자 ?? "",
      m = d.match(/(\d{4})[-./년\s]+(\d{1,2})/),
      period = m ? `${m[1]}. ${String(m[2]).padStart(2, "0")}월` : "월간",
      wb = XLSX.utils.book_new();
    const allH = [
      "고객사명",
      "상담명",
      "상담자명",
      "상담일시",
      "상담구분",
      "우선순위",
      "분류",
      "상담내용",
      "협력사명",
    ];
    XLSX.utils.book_append_sheet(
      wb,
      sheet(
        `${period} 상담이력관리`,
        allH,
        rows.map((r) => [
          r.고객사명,
          r.상담명,
          r.상담자명,
          r.상담일자,
          r.상담구분,
          r.우선순위,
          "",
          r.상담내용,
          r.협력사명,
        ]),
      ),
      "전체",
    );
    XLSX.utils.book_append_sheet(
      wb,
      sheet(
        `${period} 협력사`,
        ["상담일시", "고객사명", "상담명", "우선순위", "협력사명", "상담내용"],
        partners.flatMap(([, rs]) =>
          rs.map((r) => [
            r.상담일자,
            r.고객사명,
            r.상담명,
            r.우선순위,
            r.협력사명,
            r.상담내용,
          ]),
        ),
      ),
      "협력사",
    );
    const ch = [
      "고객사명",
      "상담명",
      "상담자명",
      "상담일시",
      "우선순위",
      "분류",
      "상담내용",
    ];
    [
      ["일반상담", "일반 상담"],
      ["클레임상담", "클레임 상담"],
    ].forEach(([sn, title]) =>
      XLSX.utils.book_append_sheet(
        wb,
        sheet(
          title,
          ch,
          rows
            .filter((r) => kind(r.상담구분) === sn)
            .map((r) => [
              r.고객사명,
              r.상담명,
              r.상담자명,
              r.상담일자,
              r.우선순위,
              "",
              r.상담내용,
            ]),
        ),
        sn,
      ),
    );
    const enterpriseSet = new Set(enterprisePartners.map(head));
    const enterprisePartnerRows = partners.filter(([name]) => enterpriseSet.has(head(name)));
    const marketPartnerRows = partners.filter(([name]) => !enterpriseSet.has(head(name)));
    XLSX.utils.book_append_sheet(
      wb,
      partnerSummarySheet(
        `${period} 협력사 요약`,
        marketPartnerRows,
        enterprisePartnerRows,
        sum,
      ),
      "협력사 요약",
    );
    const cnt = new Map<string, number>();
    rows.forEach((r) => cnt.set(r.고객사명, (cnt.get(r.고객사명) ?? 0) + 1));
    const cm = new Map<string, A[]>();
    as.forEach((a) => cm.set(a.counselor, [...(cm.get(a.counselor) ?? []), a]));
    const left: string[][] = [];
    [...cm]
      .map(([c, x]) => ({
        c,
        x,
        t: x.reduce((n, a) => n + (cnt.get(a.customer) ?? 0), 0),
      }))
      .sort((a, b) => b.t - a.t || a.c.localeCompare(b.c, "ko"))
      .forEach(({ c, x, t }) => {
        left.push([c, String(t)]);
        x.sort(
          (a, b) =>
            (cnt.get(b.customer) ?? 0) - (cnt.get(a.customer) ?? 0) ||
            a.customer.localeCompare(b.customer, "ko"),
        ).forEach((a) =>
          left.push([`  ${a.customer}`, String(cnt.get(a.customer) ?? 0)]),
        );
      });
    const cr = [
      ...new Set(
        rows
          .filter((r) => kind(r.상담구분) === "클레임상담")
          .map((r) => r.고객사명),
      ),
    ]
      .map((c) => [
        c,
        String(
          rows.filter(
            (r) => r.고객사명 === c && kind(r.상담구분) === "클레임상담",
          ).length,
        ),
      ])
      .sort(
        (a, b) => Number(b[1]) - Number(a[1]) || a[0].localeCompare(b[0], "ko"),
      );
    XLSX.utils.book_append_sheet(
      wb,
      mainSummarySheet(
        `FS사업2본부 ${period} 상담이력관리 요약`,
        left,
        cr,
      ),
      "메인요약",
    );
    XLSX.writeFile(wb, `VOC_${period.replace(/[ .]/g, "")}.xlsx`);
  }
  return (
    <main className="min-h-screen bg-slate-950 px-4 py-8 text-slate-100">
      <div className="mx-auto max-w-6xl">
        <header className="mb-8 flex items-center gap-4">
          <span className="rounded-2xl bg-amber-400 p-3 text-slate-950">
            <FileSpreadsheet />
          </span>
          <div>
            <h1 className="text-2xl font-bold">VOC 월간 보고서 생성기</h1>
            <p className="text-sm text-slate-400">
              파일은 브라우저 안에서만 처리됩니다.
            </p>
          </div>
        </header>
        <section>
          <Box title="통합 VOC 파일" file={file} set={setFile} />
          <p className="mt-3 text-sm leading-6 text-slate-400">
            시트 구성: 1. 상담내역 · 2. 담당고객사 · 3. 대기업협력사(선택)
          </p>
        </section>
        <button
          onClick={analyze}
          className="mt-4 w-full rounded-xl bg-amber-400 p-4 font-bold text-slate-950"
        >
          파일 확인하기
        </button>
        {msg && (
          <p className="mt-4 rounded-xl border border-slate-700 bg-slate-900 p-4">
            {msg}
          </p>
        )}
        {rows.length > 0 && (
          <section className="mt-8 rounded-2xl border border-slate-800 bg-slate-900 p-5">
            <div className="mb-5 flex items-center justify-between gap-3">
              <div>
                <h2 className="text-lg font-bold">협력사 요약 작성</h2>
                <p className="text-sm text-slate-400">
                  건수가 많은 순서입니다. 대기업 목록에 없으면 시장으로 분류됩니다.
                </p>
              </div>
              <button
                onClick={download}
                className="flex items-center gap-2 rounded-xl bg-emerald-400 px-5 py-3 font-bold text-slate-950"
              >
                <Download size={18} />
                XLSX 다운로드
              </button>
            </div>
            <div className="space-y-3">
              {partners.map(([n, rs]) => (
                <div
                  key={n}
                  className="grid gap-2 rounded-xl border border-slate-800 bg-slate-950 p-4 md:grid-cols-[250px_60px_1fr]"
                >
                  <b>{n}</b>
                  <span className="text-amber-300">{rs.length}건</span>
                  <textarea
                    className="min-h-20 rounded-lg border border-slate-700 bg-slate-900 p-3"
                    placeholder="요약을 직접 입력하세요"
                    value={sum[n] ?? ""}
                    onChange={(e) =>
                      setSum((s) => ({ ...s, [n]: e.target.value }))
                    }
                  />
                </div>
              ))}
            </div>
          </section>
        )}
      </div>
    </main>
  );
}
function Box({
  title,
  file,
  set,
}: {
  title: string;
  file: File | null;
  set: (f: File | null) => void;
}) {
  return (
    <label className="cursor-pointer rounded-2xl border border-dashed border-slate-700 bg-slate-900 p-6 hover:border-amber-400">
      <input
        className="sr-only"
        type="file"
        accept=".xlsx,.xls"
        onChange={(e) => set(e.target.files?.[0] ?? null)}
      />
      <Upload className="mb-4 text-amber-400" />
      <b>{title}</b>
      <p className="mt-1 text-sm text-slate-400">
        {file?.name ?? "XLSX 파일 선택"}
      </p>
    </label>
  );
}
