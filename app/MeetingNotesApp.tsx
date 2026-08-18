"use client";

import { useMemo, useState } from "react";
import mammoth from "mammoth/mammoth.browser";
import { marked } from "marked";
import { Document, HeadingLevel, Packer, Paragraph } from "docx";

type Engine = "openai" | "gemini";
type Stage = 1 | 2 | 3;
const models: Record<Engine, string> = { openai: "gpt-4o-mini", gemini: "gemini-3.5-flash-lite" };
const types = { standard: "일반 회의", evaluation: "평가위원회", weekly: "주간 업무회의", kickoff: "프로젝트 킥오프" };
const instructions = `당신은 한국어 회의록 편집자입니다. 전사문을 공식 회의록으로 변환하세요.
# 회의명
## 회의 개요
| 항목 | 내용 |
|---|---|
## 핵심 요약
- 핵심 결과
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
- 확인이 필요한 내용
원문에 없는 정보는 추정하지 말고 '확인 필요' 또는 '미정'으로 표시하세요. 제안과 검토 의견을 결정으로 바꾸지 마세요.`;

const saveBlob = (blob: Blob, name: string) => { const url = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = url; anchor.download = name; anchor.click(); URL.revokeObjectURL(url); };

export function MeetingNotesApp() {
  const [stage, setStage] = useState<Stage>(1);
  const [engine, setEngine] = useState<Engine>("openai");
  const [model, setModel] = useState(models.openai);
  const [apiKey, setApiKey] = useState("");
  const [showApiKey, setShowApiKey] = useState(false);
  const [meetingType, setMeetingType] = useState("standard");
  const [style, setStyle] = useState("official");
  const [length, setLength] = useState("standard");
  const [file, setFile] = useState<File | null>(null);
  const [transcript, setTranscript] = useState("");
  const [result, setResult] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const rendered = useMemo(() => result ? marked.parse(result, { breaks: true }) as string : "", [result]);

  const selectEngine = (next: Engine) => { setEngine(next); setModel(models[next]); setError(""); };
  const readFile = async (selected: File) => { setFile(selected); setError(""); try { const data = await mammoth.extractRawText({ arrayBuffer: await selected.arrayBuffer() }); setTranscript(data.value); } catch { setError("DOCX 파일을 읽지 못했습니다."); } };
  const next = () => { if (!transcript.trim()) return setError("DOCX 파일을 업로드하거나 전사문을 붙여넣어 주세요."); setError(""); setStage(2); };
  const summarize = async () => {
    if (!apiKey.trim()) return setError("API 키를 입력해 주세요.");
    setBusy(true); setError(""); setStage(3);
    const prompt = `${instructions}\n회의 유형: ${types[meetingType as keyof typeof types]}\n문체: ${style}\n분량: ${length}\n\n전사문:\n${transcript}`;
    try {
      let output = "";
      if (engine === "openai") {
        const response = await fetch("https://api.openai.com/v1/chat/completions", { method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey.trim()}` }, body: JSON.stringify({ model: model.trim() || models.openai, temperature: 0.2, messages: [{ role: "user", content: prompt }] }) });
        const data = await response.json(); if (!response.ok) throw new Error(data?.error?.message || "OpenAI API 요청에 실패했습니다."); output = data?.choices?.[0]?.message?.content || "결과를 받지 못했습니다.";
      } else {
        const geminiModel = (model.trim() || models.gemini).replace(/^models\//i, "").trim().toLowerCase().replace(/\s+/g, "-");
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(geminiModel)}:generateContent?key=${encodeURIComponent(apiKey.trim())}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }) });
        const data = await response.json(); if (!response.ok) throw new Error(data?.error?.message || "Gemini API 요청에 실패했습니다."); output = data?.candidates?.[0]?.content?.parts?.map((part: { text?: string }) => part.text || "").join("") || "결과를 받지 못했습니다.";
      }
      setResult(output);
    } catch (caught) { setError(caught instanceof Error ? caught.message : "요약 중 문제가 발생했습니다."); }
    finally { setBusy(false); }
  };
  const baseName = (file?.name || "회의록").replace(/\.docx$/i, "").replace(/전사문|transcript/gi, "").replace(/[_\s]+$/g, "") || "회의록";
  const downloadMarkdown = () => result && saveBlob(new Blob([result], { type: "text/markdown;charset=utf-8" }), `${baseName}_회의록.md`);
  const downloadDocx = async () => { if (!result) return; const children = result.split("\n").map((line) => new Paragraph({ text: line.replace(/^#+\s*/, ""), heading: line.startsWith("# ") ? HeadingLevel.TITLE : line.startsWith("## ") ? HeadingLevel.HEADING_1 : undefined })); saveBlob(await Packer.toBlob(new Document({ sections: [{ children }] })), `${baseName}_회의록.docx`); };
  const downloadPdf = async () => { const target = document.querySelector(".markdown-output") as HTMLElement | null; if (!target) return; const [{ default: html2canvas }, { jsPDF }] = await Promise.all([import("html2canvas"), import("jspdf")]); const canvas = await html2canvas(target, { backgroundColor: "#fff", scale: 2 }); const pdf = new jsPDF("p", "mm", "a4"); const width = 190; pdf.addImage(canvas.toDataURL("image/png"), "PNG", 10, 10, width, canvas.height * width / canvas.width); pdf.save(`${baseName}_회의록.pdf`); };

  return <main className="site-shell"><nav className="topbar"><div className="brand"><span className="brand-mark">M</span><span>Meeting Notes</span></div><span className="topbar-note">Transcript → minutes</span></nav><section className="hero"><div className="eyebrow"><span className="status-dot" /> AI 회의록 자동 생성기</div><h1>회의록 제작을,<br /><em>3단계로 간결하게.</em></h1><p className="hero-copy">Input에서 자료를 넣고, Process에서 요약 방식을 정한 뒤,<br />Output에서 최종 보고서를 확인하고 내려받으세요.</p></section><section className="stage-shell"><div className="stage-progress">{["Input", "Process", "Output"].map((label, index) => <button key={label} className={`stage-step ${stage === index + 1 ? "active" : ""} ${stage > index + 1 ? "complete" : ""}`} onClick={() => index + 1 <= stage && setStage((index + 1) as Stage)}><span>0{index + 1}</span>{label}</button>)}</div>
    {stage === 1 && <div className="stage-card card"><div className="stage-title"><div><span className="step">INPUT</span><h2>회의 자료를 준비하세요</h2></div><span className="badge">1 / 3</span></div><p className="stage-description">API 키와 회의 전사문을 입력하면 다음 단계로 넘어갈 수 있습니다.</p><div className="input-layout"><div><label className={`dropzone ${file ? "has-file" : ""}`}><input type="file" accept=".docx" onChange={(event) => event.target.files?.[0] && readFile(event.target.files[0])} />{file ? <><span className="file-icon">✓</span><strong>{file.name}</strong><small>전사문 읽기 완료</small></> : <><span className="upload-icon">↑</span><strong>DOCX 파일을 여기에 놓거나 클릭</strong><small>브라우저에서만 읽습니다</small></>}</label><label className="field-label" htmlFor="transcript">또는 전사문 직접 붙여넣기</label><textarea id="transcript" className="transcript-input" value={transcript} onChange={(event) => setTranscript(event.target.value)} placeholder="회의 전사문을 붙여넣으세요." /></div><div className="input-settings"><label className="field-label" htmlFor="meeting-type">회의록 유형</label><select id="meeting-type" className="text-input" value={meetingType} onChange={(event) => setMeetingType(event.target.value)}>{Object.entries(types).map(([key, label]) => <option key={key} value={key}>{label}</option>)}</select><label className="field-label">요약 엔진</label><div className="engine-toggle"><button className={engine === "openai" ? "active" : ""} onClick={() => selectEngine("openai")}>◎ OpenAI</button><button className={engine === "gemini" ? "active" : ""} onClick={() => selectEngine("gemini")}>✦ Gemini</button></div><label className="field-label" htmlFor="engine-model">모델</label><input id="engine-model" className="text-input" value={model} onChange={(event) => setModel(event.target.value)} /><label className="field-label" htmlFor="api-key">API 키</label><div className="api-key-field"><input id="api-key" className="text-input" type={showApiKey ? "text" : "password"} placeholder={engine === "openai" ? "sk-..." : "AIza..."} value={apiKey} onChange={(event) => setApiKey(event.target.value)} /><button type="button" className="api-key-toggle" onClick={() => setShowApiKey((visible) => !visible)} aria-label={showApiKey ? "API 키 숨기기" : "API 키 보기"} title={showApiKey ? "API 키 숨기기" : "API 키 보기"}>{showApiKey ? "숨김" : "보기"}</button></div><p className="privacy-note">🔒 키는 저장하지 않고 API 요청에만 사용합니다.</p></div></div>{error && <p className="error-message" role="alert">{error}</p>}<div className="stage-actions"><span className="stage-hint">API 키는 기본적으로 가려져 있습니다.</span><button className="primary-button stage-button" onClick={next}>Process로 이동 →</button></div></div>}
    {stage === 2 && <div className="stage-card card"><div className="stage-title"><div><span className="step">PROCESS</span><h2>요약 결과의 방향을 정하세요</h2></div><span className="badge">2 / 3</span></div><p className="stage-description">회의 목적에 맞게 문체와 분량을 선택한 뒤 생성을 시작합니다.</p><div className="process-grid"><div><label className="field-label" htmlFor="style">요약 스타일</label><select id="style" className="text-input" value={style} onChange={(event) => setStyle(event.target.value)}><option value="official">공식적이고 정중한 문체</option><option value="executive">경영진 보고용</option><option value="action">실행 항목 중심</option></select></div><div><label className="field-label" htmlFor="length">출력 분량</label><select id="length" className="text-input" value={length} onChange={(event) => setLength(event.target.value)}><option value="concise">간결하게 · 1~2쪽</option><option value="standard">표준 분량</option><option value="detailed">상세하게</option></select></div></div><div className="process-summary"><span className="summary-icon">✓</span><div><strong>{file?.name || "직접 입력한 전사문"}</strong><small>{types[meetingType as keyof typeof types]} · {engine === "openai" ? "OpenAI" : "Gemini"} · {model}</small></div></div>{error && <p className="error-message" role="alert">{error}</p>}<div className="stage-actions"><button className="secondary-button" onClick={() => setStage(1)}>← Input</button><button className="primary-button stage-button" onClick={summarize} disabled={busy}>{busy ? <><span className="spinner" /> 생성 중...</> : "회의록 생성하기 →"}</button></div></div>}
    {stage === 3 && <div className="stage-card card output-stage"><div className="stage-title"><div><span className="step">OUTPUT</span><h2>최종 회의록</h2></div>{result && <span className="result-state"><span className="status-dot" /> 생성 완료</span>}</div>{busy ? <div className="output-loading"><span className="spinner dark" /> 회의록을 생성하고 있습니다...</div> : result ? <article className="markdown-output" dangerouslySetInnerHTML={{ __html: rendered }} /> : <div className="empty-result"><div className="empty-mark">✦</div><h3>아직 결과가 없습니다</h3><p>Process 단계에서 회의록 생성을 시작하세요.</p></div>}{result && <div className="download-bar"><span>회의록 저장</span><button onClick={downloadMarkdown}>Markdown</button><button onClick={downloadDocx}>DOCX</button><button onClick={downloadPdf}>PDF</button></div>}{error && <p className="error-message" role="alert">{error}</p>}<div className="stage-actions"><button className="secondary-button" onClick={() => setStage(2)}>← Process</button><button className="secondary-button" onClick={() => { setResult(""); setStage(1); }}>새 회의록</button></div></div>}
  </section><footer><span>MEETING NOTES</span><span>개인 API 키로 안전하게 실행</span></footer></main>;
}
