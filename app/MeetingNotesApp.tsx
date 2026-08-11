"use client";

import { useMemo, useState } from "react";
import mammoth from "mammoth/mammoth.browser";
import { marked } from "marked";
import { Document, HeadingLevel, Packer, Paragraph } from "docx";

type Engine = "openai" | "gemini";
type Template = "standard" | "evaluation" | "weekly" | "kickoff";
type ActionItem = { content: string; owner: string; due: string; status: string };
const defaults: Record<Engine, string> = { openai: "gpt-4o-mini", gemini: "gemini-3.5-flash-lite" };
const templateLabels: Record<Template, string> = { standard: "일반 회의", evaluation: "평가위원회", weekly: "주간 업무회의", kickoff: "프로젝트 킥오프" };
const prompt = `당신은 한국어 회의록 편집자입니다. 아래 회의 전사문을 결정사항과 후속조치 중심의 공식 회의록으로 바꿔 주세요.

반드시 다음 Markdown 구조를 지키세요.
# 회의명
## 회의 개요
| 항목 | 내용 |
|---|---|
...
## 핵심 요약
- 핵심 결과 3~7개
## 주요 논의
### 안건별 제목
- 논의 내용과 쟁점
## 결정·의결 사항
| 결정사항 | 결과 | 조건 |
|---|---|---|
## 후속 조치
| No. | 조치 내용 | 담당자 | 기한 | 상태 |
|---:|---|---|---|---|
## 확인 필요 사항
- 전사문만으로 확정할 수 없는 내용

규칙: 전사문에 없는 날짜·금액·담당자·기한은 추정하지 말고 '확인 필요' 또는 '미정'으로 표시하세요. 질문·제안·검토 중인 내용은 확정된 결정으로 바꾸지 마세요. 인사말과 반복 발언은 줄이고 수치와 결정사항은 원문 그대로 보존하세요. 평가위원회라면 평가 결과·지원금액·조건부 선정 여부·부대조건을 명확히 정리하세요.

전사문:
`;

const download = (blob: Blob, filename: string) => { const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = filename; link.click(); URL.revokeObjectURL(url); };
const splitRow = (line: string) => line.split("|").slice(1, -1).map((cell) => cell.trim());
const parseActions = (markdown: string): ActionItem[] => {
  const section = markdown.match(/##\s*후속 조치([\s\S]*?)(?=\n##\s|$)/i)?.[1] || "";
  return section.split("\n").filter((line) => line.trim().startsWith("|") && !line.includes("---") && !line.includes("조치 내용") && !line.includes("내용"))
    .map(splitRow).filter((cells) => cells.length >= 4).map((cells) => ({ content: cells[1] || cells[0], owner: cells[2] || "미정", due: cells[3] || "미정", status: cells[4] || "확인 필요" }));
};

export function MeetingNotesApp() {
  const [engine, setEngine] = useState<Engine>("openai");
  const [template, setTemplate] = useState<Template>("standard");
  const [model, setModel] = useState(defaults.openai);
  const [apiKey, setApiKey] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [transcript, setTranscript] = useState("");
  const [result, setResult] = useState("");
  const [actions, setActions] = useState<ActionItem[]>([]);
  const [status, setStatus] = useState<"idle" | "reading" | "summarizing" | "done" | "error">("idle");
  const [error, setError] = useState("");
  const rendered = useMemo(() => result ? marked.parse(result, { breaks: true }) as string : "", [result]);

  const selectEngine = (next: Engine) => { setEngine(next); setModel(defaults[next]); setResult(""); setError(""); };
  const readFile = async (selected: File) => {
    setError(""); setFile(selected); setStatus("reading");
    try {
      if (!selected.name.toLowerCase().endsWith(".docx")) throw new Error(".docx 파일만 업로드할 수 있습니다.");
      const extracted = await mammoth.extractRawText({ arrayBuffer: await selected.arrayBuffer() });
      if (!extracted.value.trim()) throw new Error("파일에서 전사문을 읽지 못했습니다.");
      setTranscript(extracted.value); setStatus("idle");
    } catch (caught) { setFile(null); setTranscript(""); setStatus("error"); setError(caught instanceof Error ? caught.message : "파일을 읽지 못했습니다."); }
  };
  const summarize = async () => {
    if (!apiKey.trim()) return setError("API 키를 입력해 주세요.");
    if (!transcript.trim()) return setError("먼저 회의 전사문 DOCX를 업로드해 주세요.");
    setError(""); setStatus("summarizing"); setResult("");
    try {
      let output = "";
      if (engine === "openai") {
        const response = await fetch("https://api.openai.com/v1/chat/completions", { method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey.trim()}` }, body: JSON.stringify({ model: model.trim() || defaults.openai, temperature: 0.2, messages: [{ role: "user", content: `${prompt}\n선택한 회의 유형: ${templateLabels[template]}\n${transcript}` }] }) });
        const data = await response.json(); if (!response.ok) throw new Error(data?.error?.message || "OpenAI API 요청에 실패했습니다."); output = data?.choices?.[0]?.message?.content || "결과를 받지 못했습니다.";
      } else {
        const geminiModel = (model.trim() || defaults.gemini).replace(/^models\//i, "").trim().toLowerCase().replace(/\s+/g, "-");
        const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(geminiModel)}:generateContent?key=${encodeURIComponent(apiKey.trim())}`;
        const response = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ contents: [{ parts: [{ text: `${prompt}\n선택한 회의 유형: ${templateLabels[template]}\n${transcript}` }] }] }) });
        const data = await response.json(); if (!response.ok) throw new Error(data?.error?.message || "Gemini API 요청에 실패했습니다."); output = data?.candidates?.[0]?.content?.parts?.map((part: { text?: string }) => part.text || "").join("") || "결과를 받지 못했습니다.";
      }
      setResult(output); setActions(parseActions(output)); setStatus("done");
    } catch (caught) { setStatus("error"); setError(caught instanceof Error ? caught.message : "요약 중 문제가 발생했습니다."); }
  };

  const baseName = (file?.name || "회의록").replace(/\.docx$/i, "").replace(/전사문|transcript/gi, "").replace(/[_\s]+$/g, "") || "회의록";
  const downloadMarkdown = () => result && download(new Blob([result], { type: "text/markdown;charset=utf-8" }), `${baseName}_회의록.md`);
  const downloadDocx = async () => {
    if (!result) return;
    const children = result.split("\n").map((line) => new Paragraph({ text: line.replace(/^#+\s*/, ""), heading: line.startsWith("# ") ? HeadingLevel.TITLE : line.startsWith("## ") ? HeadingLevel.HEADING_1 : line.startsWith("### ") ? HeadingLevel.HEADING_2 : undefined }));
    download(await Packer.toBlob(new Document({ sections: [{ children }] })), `${baseName}_회의록.docx`);
  };
  const downloadPdf = async () => {
    const target = document.querySelector(".markdown-output") as HTMLElement | null;
    if (!target) return;
    const [{ default: html2canvas }, { jsPDF }] = await Promise.all([import("html2canvas"), import("jspdf")]);
    const canvas = await html2canvas(target, { backgroundColor: "#ffffff", scale: 2 });
    const pdf = new jsPDF("p", "mm", "a4"); const width = 190; const height = canvas.height * width / canvas.width; pdf.addImage(canvas.toDataURL("image/png"), "PNG", 10, 10, width, height); pdf.save(`${baseName}_회의록.pdf`);
  };
  const applyActions = () => {
    if (!result || !actions.length) return;
    const table = ["| No. | 조치 내용 | 담당자 | 기한 | 상태 |", "|---:|---|---|---|---|", ...actions.map((item, index) => `| ${index + 1} | ${item.content} | ${item.owner} | ${item.due} | ${item.status} |`)].join("\n");
    const next = result.replace(/(##\s*후속 조치[\s\S]*?)(?=\n##\s|$)/i, `## 후속 조치\n${table}\n`); setResult(next);
  };

  return <main className="site-shell">
    <nav className="topbar"><div className="brand"><span className="brand-mark">M</span><span>Meeting Notes</span></div><span className="topbar-note">Transcript → minutes</span></nav>
    <section className="hero"><div className="eyebrow"><span className="status-dot" /> AI 회의록 요약 도구</div><h1>긴 전사문을,<br /><em>결정 가능한 회의록</em>으로.</h1><p className="hero-copy">회의 전사문을 올리고 원하는 엔진을 선택하세요.<br />핵심 논의, 결정사항, 후속조치를 한 장으로 정리합니다.</p></section>
    <section className="workspace-grid">
      <div className="control-card card">
        <div className="card-heading"><div><span className="step">01</span><h2>전사문 업로드</h2></div><span className="badge">DOCX</span></div>
        <label className={`dropzone ${file ? "has-file" : ""}`}><input type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={(event) => event.target.files?.[0] && readFile(event.target.files[0])} />{file ? <><span className="file-icon">✓</span><strong>{file.name}</strong><small>{(file.size / 1024).toFixed(1)} KB · 전사문 읽기 완료</small></> : <><span className="upload-icon">↑</span><strong>DOCX 파일을 여기에 놓거나 클릭</strong><small>회의 전사문 .docx · 브라우저에서만 읽습니다</small></>}</label>
        <label className="field-label" htmlFor="template">회의록 템플릿</label><select id="template" className="text-input" value={template} onChange={(event) => setTemplate(event.target.value as Template)}>{Object.entries(templateLabels).map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select>
        <div className="section-divider" />
        <div className="card-heading compact"><div><span className="step">02</span><h2>요약 엔진 선택</h2></div><span className="secure-label">KEY 미저장</span></div>
        <div className="engine-toggle" role="tablist"><button className={engine === "openai" ? "active" : ""} onClick={() => selectEngine("openai")} role="tab"><span className="engine-logo openai-logo">◎</span>OpenAI</button><button className={engine === "gemini" ? "active" : ""} onClick={() => selectEngine("gemini")} role="tab"><span className="engine-logo gemini-logo">✦</span>Gemini</button></div>
        <label className="field-label" htmlFor="model">모델</label><input id="model" className="text-input" value={model} onChange={(event) => setModel(event.target.value)} />
        <label className="field-label" htmlFor="api-key">API 키</label><div className="key-input"><input id="api-key" type="password" placeholder={engine === "openai" ? "sk-..." : "AIza..."} value={apiKey} onChange={(event) => setApiKey(event.target.value)} /><span>●●●</span></div>
        <p className="privacy-note">🔒 키는 저장하지 않고 이 브라우저에서 API 요청에만 사용합니다.</p><button className="primary-button" onClick={summarize} disabled={status === "reading" || status === "summarizing"}>{status === "summarizing" ? <><span className="spinner" /> 회의록 만드는 중...</> : "회의록 생성하기 →"}</button>{error && <p className="error-message" role="alert">{error}</p>}
      </div>
      <div className="result-card card"><div className="result-header"><div><span className="step">03</span><h2>회의록 결과</h2></div>{result && <span className="result-state"><span className="status-dot" /> 생성 완료</span>}</div>{!result ? <div className="empty-result"><div className="empty-mark">✦</div><h3>요약 결과가 여기에 표시됩니다</h3><p>왼쪽에서 전사문을 업로드하고<br />요약 엔진을 선택해 시작하세요.</p><div className="format-note"><span>#</span> Markdown으로 깔끔하게 정리</div></div> : <article className="markdown-output" dangerouslySetInnerHTML={{ __html: rendered }} />}{status === "reading" && <div className="loading-overlay"><span className="spinner dark" /> 파일을 읽는 중...</div>}</div>
      {result && <div className="actions-panel card">
        <div className="actions-heading"><div><span className="step">04</span><h2>후속 조치 확인</h2></div><button className="mini-button" onClick={applyActions}>표에 반영</button></div>
        {actions.length ? <div className="actions-table"><div className="actions-row actions-head"><span>조치 내용</span><span>담당자</span><span>기한</span><span>상태</span></div>{actions.map((item, index) => <div className="actions-row" key={`${item.content}-${index}`}><input value={item.content} onChange={(event) => setActions((items) => items.map((row, i) => i === index ? { ...row, content: event.target.value } : row))} /><input value={item.owner} onChange={(event) => setActions((items) => items.map((row, i) => i === index ? { ...row, owner: event.target.value } : row))} /><input value={item.due} onChange={(event) => setActions((items) => items.map((row, i) => i === index ? { ...row, due: event.target.value } : row))} /><input value={item.status} onChange={(event) => setActions((items) => items.map((row, i) => i === index ? { ...row, status: event.target.value } : row))} /></div>)}</div> : <p className="muted-note">전사문에서 후속 조치를 찾지 못했습니다. 확인 필요 사항으로 검토해 주세요.</p>}
        <div className="download-bar"><span>회의록 저장</span><button onClick={downloadMarkdown}>Markdown</button><button onClick={downloadDocx}>DOCX</button><button onClick={downloadPdf}>PDF</button></div>
      </div>}
    </section>
    <footer><span>MEETING NOTES</span><span>개인 API 키로 안전하게 실행</span></footer>
  </main>;
}
