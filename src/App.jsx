import { useState, useRef } from "react";
import html2pdf from "html2pdf.js";

// ─── Constants ───────────────────────────────────────────────────────────────

const QUESTION_TYPES = [
  { id: "fill",      label: "Fill in the Blanks",     defaultMarks: 1, hasOptions: false, hasDiagram: true  },
  { id: "mcq3",     label: "MCQ (3 options)",         defaultMarks: 1, hasOptions: 3,    hasDiagram: true  },
  { id: "mcq4",     label: "MCQ (4 options)",         defaultMarks: 1, hasOptions: 4,    hasDiagram: true  },
  { id: "tf",       label: "True or False",           defaultMarks: 1, hasOptions: false, hasDiagram: false },
  { id: "oneword",  label: "One Word Answer",         defaultMarks: 1, hasOptions: false, hasDiagram: false },
  { id: "short2",   label: "Short Answer (2 marks)",  defaultMarks: 2, hasOptions: false, hasDiagram: true  },
  { id: "short3",   label: "Short Answer (3 marks)",  defaultMarks: 3, hasOptions: false, hasDiagram: true  },
  { id: "long",     label: "Long Answer (5 marks)",   defaultMarks: 5, hasOptions: false, hasDiagram: true  },
  { id: "match",    label: "Match the Following",     defaultMarks: 5, hasOptions: false, hasDiagram: false },
  { id: "num",      label: "Numerical / Independent", defaultMarks: 4, hasOptions: false, hasDiagram: true  },
  { id: "diagram",  label: "Diagram Based",           defaultMarks: 3, hasOptions: false, hasDiagram: true  },
];

const DIAGRAM_OPTIONS = ["upload", "draw", "space"];
const ALPHA = ["A","B","C","D","E","F","G","H","I","J"];
const OPT_ALPHA = ["a","b","c","d"];

function uid() { return Math.random().toString(36).slice(2,9); }

function makeQuestion(typeId) {
  const type = QUESTION_TYPES.find(t => t.id === typeId);
  return {
    id: uid(),
    text: "",
    diagram: null, // { mode: 'upload'|'draw'|'space', dataUrl: null }
    options: type.hasOptions
      ? Array.from({ length: type.hasOptions }, () => "")
      : null,
    matchLeft: type.id === "match" ? ["","","","",""] : null,
    matchRight: type.id === "match" ? ["","","","",""] : null,
  };
}

function makeSection() {
  return {
    id: uid(),
    typeId: "fill",
    heading: "",
    questionCount: 5,
    marksPerQ: 1,
    questions: Array.from({ length: 5 }, () => makeQuestion("fill")),
  };
}

const DEFAULT_HEADER = {
  school: "Goodwill Public School",
  location: "Uttam Nagar",
  exam: "Half Early Examination",
  session: "2026–27",
  classVal: "VI",
  subject: "Computer",
  duration: "1.5 Hours",
  maxMarks: "30",
  date: "",
  instructions: "",
  logoUrl: null,
};

// ─── Tiny UI atoms ────────────────────────────────────────────────────────────

const Input = ({ label, value, onChange, placeholder, type = "text", small }) => (
  <div style={{ marginBottom: small ? 8 : 14 }}>
    {label && <label style={{ display:"block", fontSize:11, fontWeight:700, color:"#6b7280", marginBottom:3, textTransform:"uppercase", letterSpacing:"0.04em" }}>{label}</label>}
    <input
      type={type}
      value={value}
      onChange={e => onChange(e.target.value)}
      placeholder={placeholder}
      style={{
        width:"100%", boxSizing:"border-box",
        padding: small ? "7px 10px" : "9px 12px",
        borderRadius:7, border:"1.5px solid #e5e7eb",
        fontSize: small ? 13 : 14, outline:"none", background:"#fff",
        fontFamily:"inherit", color:"#111",
        transition:"border-color .15s"
      }}
      onFocus={e => e.target.style.borderColor="#6366f1"}
      onBlur={e => e.target.style.borderColor="#e5e7eb"}
    />
  </div>
);

const Textarea = ({ label, value, onChange, placeholder, rows=2 }) => (
  <div style={{ marginBottom:14 }}>
    {label && <label style={{ display:"block", fontSize:11, fontWeight:700, color:"#6b7280", marginBottom:3, textTransform:"uppercase", letterSpacing:"0.04em" }}>{label}</label>}
    <textarea
      value={value} onChange={e => onChange(e.target.value)}
      placeholder={placeholder} rows={rows}
      style={{
        width:"100%", boxSizing:"border-box",
        padding:"9px 12px", borderRadius:7,
        border:"1.5px solid #e5e7eb", fontSize:14, outline:"none",
        fontFamily:"inherit", resize:"vertical", color:"#111",
        transition:"border-color .15s"
      }}
      onFocus={e => e.target.style.borderColor="#6366f1"}
      onBlur={e => e.target.style.borderColor="#e5e7eb"}
    />
  </div>
);

const Btn = ({ onClick, children, variant="primary", small, disabled, full }) => {
  const styles = {
    primary:   { background:"#6366f1", color:"#fff", border:"none" },
    secondary: { background:"#f3f4f6", color:"#374151", border:"1.5px solid #e5e7eb" },
    danger:    { background:"#fee2e2", color:"#dc2626", border:"1.5px solid #fca5a5" },
    success:   { background:"#d1fae5", color:"#065f46", border:"1.5px solid #6ee7b7" },
    ghost:     { background:"transparent", color:"#6b7280", border:"1.5px solid transparent" },
  };
  return (
    <button onClick={onClick} disabled={disabled}
      style={{
        ...styles[variant],
        padding: small ? "5px 10px" : "9px 16px",
        borderRadius:7, fontSize: small ? 12 : 13, fontWeight:600,
        cursor: disabled ? "not-allowed" : "pointer",
        opacity: disabled ? 0.5 : 1,
        width: full ? "100%" : undefined,
        fontFamily:"inherit",
        transition:"opacity .15s"
      }}>
      {children}
    </button>
  );
};

const Card = ({ children, style }) => (
  <div style={{ background:"#fff", borderRadius:12, border:"1.5px solid #e5e7eb", padding:20, marginBottom:16, ...style }}>
    {children}
  </div>
);

const Badge = ({ children, color="#e0e7ff", textColor="#4338ca" }) => (
  <span style={{ background:color, color:textColor, fontSize:11, fontWeight:700, padding:"3px 8px", borderRadius:20, letterSpacing:"0.03em" }}>
    {children}
  </span>
);

// ─── Draw Canvas ──────────────────────────────────────────────────────────────

function DrawCanvas({ onSave, onCancel }) {
  const canvasRef = useRef();
  const drawing = useRef(false);
  const lastPos = useRef({ x:0, y:0 });

  function getPos(e, canvas) {
    const rect = canvas.getBoundingClientRect();
    const src = e.touches ? e.touches[0] : e;
    return { x: src.clientX - rect.left, y: src.clientY - rect.top };
  }

  function start(e) {
    e.preventDefault();
    drawing.current = true;
    lastPos.current = getPos(e, canvasRef.current);
  }

  function move(e) {
    e.preventDefault();
    if (!drawing.current) return;
    const ctx = canvasRef.current.getContext("2d");
    const pos = getPos(e, canvasRef.current);
    ctx.beginPath();
    ctx.moveTo(lastPos.current.x, lastPos.current.y);
    ctx.lineTo(pos.x, pos.y);
    ctx.strokeStyle = "#1e293b";
    ctx.lineWidth = 2;
    ctx.lineCap = "round";
    ctx.stroke();
    lastPos.current = pos;
  }

  function end(e) { e.preventDefault(); drawing.current = false; }

  function clear() {
    const ctx = canvasRef.current.getContext("2d");
    ctx.clearRect(0, 0, 600, 300);
  }

  function save() {
    onSave(canvasRef.current.toDataURL("image/png"));
  }

  return (
    <div style={{ background:"#f8fafc", borderRadius:10, padding:16, border:"1.5px solid #e5e7eb" }}>
      <div style={{ fontSize:12, fontWeight:600, color:"#6b7280", marginBottom:8 }}>Draw Diagram</div>
      <canvas ref={canvasRef} width={560} height={240}
        style={{ background:"#fff", borderRadius:7, border:"1px solid #d1d5db", display:"block", maxWidth:"100%", touchAction:"none", cursor:"crosshair" }}
        onMouseDown={start} onMouseMove={move} onMouseUp={end} onMouseLeave={end}
        onTouchStart={start} onTouchMove={move} onTouchEnd={end}
      />
      <div style={{ display:"flex", gap:8, marginTop:10 }}>
        <Btn onClick={clear} variant="secondary" small>Clear</Btn>
        <Btn onClick={save} variant="primary" small>Use this diagram</Btn>
        <Btn onClick={onCancel} variant="ghost" small>Cancel</Btn>
      </div>
    </div>
  );
}

// ─── Diagram Picker ───────────────────────────────────────────────────────────

function DiagramPicker({ diagram, onChange }) {
  const [mode, setMode] = useState(null);
  const fileRef = useRef();

  function handleUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => onChange({ mode:"upload", dataUrl: ev.target.result });
    reader.readAsDataURL(file);
    setMode(null);
  }

  if (diagram) return (
    <div style={{ marginTop:8 }}>
      {diagram.mode === "space"
        ? <div style={{ border:"1.5px dashed #d1d5db", borderRadius:7, padding:"12px 16px", fontSize:12, color:"#9ca3af", textAlign:"center" }}>Diagram space will appear in print</div>
        : <img src={diagram.dataUrl} alt="diagram" style={{ maxWidth:"100%", borderRadius:7, border:"1px solid #e5e7eb", display:"block" }} />
      }
      <div style={{ marginTop:6 }}>
        <Btn onClick={() => onChange(null)} variant="danger" small>Remove Diagram</Btn>
      </div>
    </div>
  );

  if (mode === "draw") return (
    <div style={{ marginTop:8 }}>
      <DrawCanvas onSave={dataUrl => { onChange({ mode:"draw", dataUrl }); setMode(null); }} onCancel={() => setMode(null)} />
    </div>
  );

  return (
    <div style={{ marginTop:8 }}>
      {!mode && (
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          <Btn onClick={() => fileRef.current.click()} variant="secondary" small>Upload Image</Btn>
          <Btn onClick={() => setMode("draw")} variant="secondary" small>Draw Here</Btn>
          <Btn onClick={() => onChange({ mode:"space", dataUrl:null })} variant="secondary" small>Leave Space</Btn>
        </div>
      )}
      <input ref={fileRef} type="file" accept="image/*" style={{ display:"none" }} onChange={handleUpload} />
    </div>
  );
}

// ─── Single Question Editor ───────────────────────────────────────────────────

function QuestionEditor({ q, index, typeId, onChange, showDiagram }) {
  const type = QUESTION_TYPES.find(t => t.id === typeId);
  const [showDiagPicker, setShowDiagPicker] = useState(!!q.diagram);

  function upd(patch) { onChange({ ...q, ...patch }); }

  return (
    <div style={{ background:"#fafafa", borderRadius:9, border:"1.5px solid #f0f0f0", padding:"14px 16px", marginBottom:10 }}>
      <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
        <Badge color="#ede9fe" textColor="#5b21b6">{ALPHA[index]}</Badge>
        <span style={{ fontSize:11, color:"#9ca3af", fontWeight:600 }}>{type.label}</span>
      </div>

      {typeId === "match" ? (
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
          <div>
            <div style={{ fontSize:11, fontWeight:700, color:"#6b7280", marginBottom:6 }}>Column A</div>
            {q.matchLeft.map((v,i) => (
              <input key={i} value={v} placeholder={`${i+1}.`}
                onChange={e => { const arr=[...q.matchLeft]; arr[i]=e.target.value; upd({matchLeft:arr}); }}
                style={{ display:"block", width:"100%", boxSizing:"border-box", marginBottom:5, padding:"7px 10px", borderRadius:6, border:"1.5px solid #e5e7eb", fontSize:13, outline:"none", fontFamily:"inherit" }}
                onFocus={e=>e.target.style.borderColor="#6366f1"} onBlur={e=>e.target.style.borderColor="#e5e7eb"} />
            ))}
          </div>
          <div>
            <div style={{ fontSize:11, fontWeight:700, color:"#6b7280", marginBottom:6 }}>Column B</div>
            {q.matchRight.map((v,i) => (
              <input key={i} value={v} placeholder={`${String.fromCharCode(97+i)}.`}
                onChange={e => { const arr=[...q.matchRight]; arr[i]=e.target.value; upd({matchRight:arr}); }}
                style={{ display:"block", width:"100%", boxSizing:"border-box", marginBottom:5, padding:"7px 10px", borderRadius:6, border:"1.5px solid #e5e7eb", fontSize:13, outline:"none", fontFamily:"inherit" }}
                onFocus={e=>e.target.style.borderColor="#6366f1"} onBlur={e=>e.target.style.borderColor="#e5e7eb"} />
            ))}
          </div>
        </div>
      ) : (
        <>
          <textarea value={q.text} onChange={e => upd({text:e.target.value})}
            placeholder={typeId==="fill" ? "Write statement with ___ for blank" : typeId==="tf" ? "Write a true/false statement" : "Write your question here"}
            rows={2}
            style={{ width:"100%", boxSizing:"border-box", padding:"8px 10px", borderRadius:6, border:"1.5px solid #e5e7eb", fontSize:13, fontFamily:"inherit", resize:"vertical", outline:"none" }}
            onFocus={e=>e.target.style.borderColor="#6366f1"} onBlur={e=>e.target.style.borderColor="#e5e7eb"} />

          {type.hasOptions && (
            <div style={{ marginTop:8 }}>
              {Array.from({length:type.hasOptions},(_,i)=>(
                <div key={i} style={{ display:"flex", gap:6, marginBottom:5, alignItems:"center" }}>
                  <span style={{ fontSize:12, color:"#9ca3af", fontWeight:700, width:20, flexShrink:0 }}>{OPT_ALPHA[i]})</span>
                  <input value={q.options[i]} onChange={e=>{ const arr=[...q.options]; arr[i]=e.target.value; upd({options:arr}); }}
                    placeholder={`Option ${OPT_ALPHA[i].toUpperCase()}`}
                    style={{ flex:1, padding:"6px 10px", borderRadius:6, border:"1.5px solid #e5e7eb", fontSize:13, outline:"none", fontFamily:"inherit" }}
                    onFocus={e=>e.target.style.borderColor="#6366f1"} onBlur={e=>e.target.style.borderColor="#e5e7eb"} />
                </div>
              ))}
            </div>
          )}
        </>
      )}

      {showDiagram && type.hasDiagram && (
        <div style={{ marginTop:8 }}>
          {!showDiagPicker && !q.diagram
            ? <button onClick={()=>setShowDiagPicker(true)} style={{ fontSize:11, color:"#6366f1", background:"none", border:"none", cursor:"pointer", padding:0, fontWeight:600 }}>+ Add Diagram</button>
            : <DiagramPicker diagram={q.diagram} onChange={d=>{ upd({diagram:d}); if(!d) setShowDiagPicker(false); }} />
          }
        </div>
      )}
    </div>
  );
}

// ─── Section Editor ───────────────────────────────────────────────────────────

function SectionEditor({ section, index, total, onChange, onDelete, onMoveUp, onMoveDown }) {
  const [open, setOpen] = useState(true);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiTopic, setAiTopic] = useState("");
  const [aiUses, setAiUses] = useState(() => {
    try {
      const d = JSON.parse(localStorage.getItem("qp_ai_usage") || "{}");
      return d.date === new Date().toISOString().slice(0,10) ? (Number(d.count) || 0) : 0;
    } catch { return 0; }
  });
  const AI_DAILY_LIMIT = 10;
  const type = QUESTION_TYPES.find(t => t.id === section.typeId);
  const sectionMarks = section.questionCount * section.marksPerQ;

  function upd(patch) { onChange({ ...section, ...patch }); }

  function changeType(newTypeId) {
    const newType = QUESTION_TYPES.find(t => t.id === newTypeId);
    upd({
      typeId: newTypeId,
      marksPerQ: newType.defaultMarks,
      questions: Array.from({ length: section.questionCount }, () => makeQuestion(newTypeId)),
    });
  }

  function changeCount(n) {
    const count = Math.max(1, Math.min(20, parseInt(n)||1));
    const qs = [...section.questions];
    while (qs.length < count) qs.push(makeQuestion(section.typeId));
    upd({ questionCount: count, questions: qs.slice(0, count) });
  }

  function updateQuestion(i, q) {
    const qs = [...section.questions];
    qs[i] = q;
    upd({ questions: qs });
  }

  async function generateAI() {
    if (!aiTopic.trim()) return;
    if (aiUses >= AI_DAILY_LIMIT) {
      alert(`You have reached the ${AI_DAILY_LIMIT}-generation daily limit on this device.`);
      return;
    }
    setAiLoading(true);
    try {
      const typeLabel = type.label;
      const count = section.questionCount;
      const prompt = `Generate ${count} "${typeLabel}" questions for a school exam on topic: "${aiTopic}".
Use clear, age-appropriate school language. Do not repeat questions. Return ONLY valid JSON array, no other text, no markdown:
${section.typeId === "mcq3" || section.typeId === "mcq4"
  ? `[{"text":"Question?","options":["opt1","opt2","opt3"${section.typeId==="mcq4"?',"opt4"':''}]}]`
  : section.typeId === "match"
  ? `[{"matchLeft":["term1","term2","term3","term4","term5"],"matchRight":["def1","def2","def3","def4","def5"]}]`
  : `[{"text":"Question or statement here"}]`
}
Generate exactly ${count} items.`;

      const res = await fetch("/api/generate", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body: JSON.stringify({ prompt })
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Request failed (${res.status})`);
      }
      const data = await res.json();
      if (!data.content) throw new Error(data.error || "AI returned no content.");
      const raw = data.content || "";
      const clean = raw.replace(/```json|```/g,"").trim();
      const start = clean.indexOf("[");
      const end = clean.lastIndexOf("]");
      const jsonText = start >= 0 && end > start ? clean.slice(start, end + 1) : clean;
      const parsed = JSON.parse(jsonText);
      if (!Array.isArray(parsed) || parsed.length === 0) throw new Error("AI returned an invalid question list.");
      const qs = parsed.map(item => ({
        id: uid(),
        text: item.text || "",
        diagram: null,
        options: item.options || (type.hasOptions ? Array.from({length:type.hasOptions},()=>"") : null),
        matchLeft: item.matchLeft || (section.typeId==="match" ? ["","","","",""] : null),
        matchRight: item.matchRight || (section.typeId==="match" ? ["","","","",""] : null),
      }));
      upd({ questions: qs.slice(0, section.questionCount) });
      const nextUses = aiUses + 1;
      setAiUses(nextUses);
      try { localStorage.setItem("qp_ai_usage", JSON.stringify({ date: new Date().toISOString().slice(0,10), count: nextUses })); } catch {}
    } catch(e) { alert(`AI generation failed. ${e?.message || "Please fill manually."}`); }
    setAiLoading(false);
  }

  return (
    <Card>
      {/* Section header */}
      <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom: open ? 16 : 0 }}>
        <div style={{ display:"flex", flexDirection:"column", gap:2 }}>
          <button onClick={onMoveUp} disabled={index===0} style={{ background:"none", border:"none", cursor:"pointer", fontSize:12, color: index===0?"#d1d5db":"#6b7280", padding:"1px 4px" }}>▲</button>
          <button onClick={onMoveDown} disabled={index===total-1} style={{ background:"none", border:"none", cursor:"pointer", fontSize:12, color: index===total-1?"#d1d5db":"#6b7280", padding:"1px 4px" }}>▼</button>
        </div>
        <Badge color="#fef3c7" textColor="#92400e">Q{index+1}</Badge>
        <div style={{ flex:1, fontSize:14, fontWeight:700, color:"#111" }}>{section.heading || type.label}</div>
        <Badge color="#f0fdf4" textColor="#166534">{sectionMarks} marks</Badge>
        <button onClick={()=>setOpen(o=>!o)} style={{ background:"none", border:"none", cursor:"pointer", fontSize:16, color:"#6b7280", padding:"0 4px" }}>
          {open ? "▾" : "▸"}
        </button>
        <button onClick={onDelete} style={{ background:"none", border:"none", cursor:"pointer", fontSize:16, color:"#f87171", padding:"0 4px" }}>✕</button>
      </div>

      {open && (
        <>
          {/* Section config */}
          <div className="qp-section-grid" style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:10, marginBottom:14 }}>
            <div>
              <label style={{ display:"block", fontSize:11, fontWeight:700, color:"#6b7280", marginBottom:3, textTransform:"uppercase", letterSpacing:"0.04em" }}>Type</label>
              <select value={section.typeId} onChange={e=>changeType(e.target.value)}
                style={{ width:"100%", padding:"8px 10px", borderRadius:7, border:"1.5px solid #e5e7eb", fontSize:13, outline:"none", fontFamily:"inherit", background:"#fff" }}>
                {QUESTION_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
              </select>
            </div>
            <div>
              <label style={{ display:"block", fontSize:11, fontWeight:700, color:"#6b7280", marginBottom:3, textTransform:"uppercase", letterSpacing:"0.04em" }}>No. of Questions</label>
              <input type="number" min={1} max={20} value={section.questionCount}
                onChange={e=>changeCount(e.target.value)}
                style={{ width:"100%", boxSizing:"border-box", padding:"8px 10px", borderRadius:7, border:"1.5px solid #e5e7eb", fontSize:13, outline:"none", fontFamily:"inherit" }} />
            </div>
            <div>
              <label style={{ display:"block", fontSize:11, fontWeight:700, color:"#6b7280", marginBottom:3, textTransform:"uppercase", letterSpacing:"0.04em" }}>Marks/Question</label>
              <input type="number" min={1} max={20} value={section.marksPerQ}
                onChange={e=>upd({marksPerQ:Math.max(1,parseInt(e.target.value)||1)})}
                style={{ width:"100%", boxSizing:"border-box", padding:"8px 10px", borderRadius:7, border:"1.5px solid #e5e7eb", fontSize:13, outline:"none", fontFamily:"inherit" }} />
            </div>
          </div>

          <div style={{ marginBottom:14 }}>
            <label style={{ display:"block", fontSize:11, fontWeight:700, color:"#6b7280", marginBottom:3, textTransform:"uppercase", letterSpacing:"0.04em" }}>Custom Heading (optional)</label>
            <input value={section.heading} onChange={e=>upd({heading:e.target.value})}
              placeholder={`e.g. "Section A" or leave blank to use type name`}
              style={{ width:"100%", boxSizing:"border-box", padding:"8px 10px", borderRadius:7, border:"1.5px solid #e5e7eb", fontSize:13, outline:"none", fontFamily:"inherit" }} />
          </div>

          {/* AI generator */}
          <div style={{ background:"#eef2ff", borderRadius:8, padding:"10px 12px", marginBottom:14, display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
            <span style={{ fontSize:11, fontWeight:700, color:"#4338ca" }}>AI Generate</span>
            <span style={{ fontSize:10, color:"#6366f1" }}>{AI_DAILY_LIMIT - aiUses} left today</span>
            <input value={aiTopic} onChange={e=>setAiTopic(e.target.value)}
              placeholder="Topic e.g. Photosynthesis"
              style={{ flex:1, minWidth:120, padding:"6px 10px", borderRadius:6, border:"1.5px solid #c7d2fe", fontSize:12, outline:"none", fontFamily:"inherit" }}
              onFocus={e=>e.target.style.borderColor="#6366f1"} onBlur={e=>e.target.style.borderColor="#c7d2fe"} />
            <Btn onClick={generateAI} disabled={aiLoading} small>{aiLoading ? "..." : "Generate"}</Btn>
          </div>

          {/* Questions */}
          <div>
            {section.questions.map((q,i) => (
              <QuestionEditor key={q.id} q={q} index={i} typeId={section.typeId}
                onChange={nq => updateQuestion(i, nq)}
                showDiagram={true} />
            ))}
          </div>
        </>
      )}
    </Card>
  );
}

// ─── Preview ──────────────────────────────────────────────────────────────────

function buildPreviewHTML(header, sections) {
  const { school, location, exam, session, classVal, subject, duration, maxMarks, date, instructions, logoUrl } = header;

  const sectionHTML = sections.map((sec, si) => {
    const type = QUESTION_TYPES.find(t => t.id === sec.typeId);
    const heading = sec.heading || type.label;
    const perQ = sec.marksPerQ;
    const total = sec.questionCount * perQ;
    const marksStr = perQ === 1 ? `(${total})` : `(${perQ}x${sec.questionCount}=${total})`;

    const qHTML = sec.questions.map((q, qi) => {
      let inner = "";

      if (sec.typeId === "match") {
        inner = `<table style="width:100%;border-collapse:collapse;margin-top:6px;">
          <tr><th style="text-align:left;padding:3px 8px;border:1px solid #ccc;width:50%">Column A</th><th style="text-align:left;padding:3px 8px;border:1px solid #ccc">Column B</th></tr>
          ${(q.matchLeft||[]).map((l,i)=>`<tr><td style="padding:3px 8px;border:1px solid #ccc">${i+1}. ${l||"_______"}</td><td style="padding:3px 8px;border:1px solid #ccc">${String.fromCharCode(97+i)}. ${(q.matchRight||[])[i]||"_______"}</td></tr>`).join("")}
        </table>`;
      } else {
        inner = `<div>${q.text || "___________________________________"}</div>`;
        if (sec.typeId === "fill") {
          inner = `<div>${q.text || "___________________________________ _______________"}</div>`;
        }
        if (sec.typeId === "tf") {
          inner = `<div>${q.text || "___________________________________"} <span style="margin-left:24px">__________</span></div>`;
        }
        if (q.options) {
          inner += `<div style="margin-left:20px;margin-top:4px">${q.options.map((o,oi)=>`<div style="margin-bottom:2px">${OPT_ALPHA[oi]}) ${o||"_______"}</div>`).join("")}</div>`;
        }
        if (q.diagram) {
          if (q.diagram.mode === "space") {
            inner += `<div style="border:1px dashed #aaa;height:80px;margin-top:8px;display:flex;align-items:center;justify-content:center;color:#aaa;font-size:12px">Diagram</div>`;
          } else if (q.diagram.dataUrl) {
            inner += `<div style="margin-top:8px"><img src="${q.diagram.dataUrl}" style="max-width:100%;max-height:140px;display:block;border:1px solid #ddd" /></div>`;
          }
        }
        if (["short2","short3","long","num","diagram"].includes(sec.typeId)) {
          const lines = sec.typeId === "long" ? 6 : 3;
          inner += Array.from({length:lines},()=>`<div style="border-bottom:1px solid #ccc;height:22px;margin-top:4px"></div>`).join("");
        }
      }

      return `<div style="margin-bottom:12px">
        <span style="font-weight:600">${ALPHA[qi]}.</span> ${inner}
      </div>`;
    }).join("");

    return `<div style="margin-bottom:22px">
      <div style="font-weight:700;margin-bottom:8px">Q${si+1}. ${heading}: ${marksStr}</div>
      ${qHTML}
    </div>`;
  }).join("");

  const totalMarks = sections.reduce((s,sec)=>s + sec.questionCount*sec.marksPerQ, 0);

  return `<div style="font-family:'Times New Roman',serif;max-width:680px;margin:0 auto;padding:32px 28px;color:#000;background:#fff;font-size:14px;line-height:1.6">
    ${logoUrl ? `<div style="text-align:center;margin-bottom:8px"><img src="${logoUrl}" style="height:60px" /></div>` : ""}
    <div style="text-align:center;margin-bottom:6px">
      <div style="font-size:19px;font-weight:bold;text-transform:uppercase;letter-spacing:0.03em">${school}</div>
      <div style="font-size:13px;text-transform:uppercase">${location}</div>
      <div style="font-size:13px;text-transform:uppercase;margin-top:3px">${exam} ${session}</div>
      <div style="font-size:13px;text-transform:uppercase">CLASS ${classVal}</div>
      ${date ? `<div style="font-size:12px;margin-top:2px">Date: ${date}</div>` : ""}
    </div>
    <div style="display:flex;justify-content:space-between;border-top:1.5px solid #000;border-bottom:1.5px solid #000;padding:6px 0;margin:10px 0 16px;font-size:13px">
      <span>Time: ${duration}</span>
      <span>Subject: ${subject}</span>
      <span>M.M.: ${maxMarks} ${totalMarks != parseInt(maxMarks) ? `<span style="color:#dc2626;font-size:11px">(current: ${totalMarks})</span>` : ""}</span>
    </div>
    ${instructions ? `<div style="font-size:12px;margin-bottom:14px;padding:8px 12px;border:1px solid #e5e7eb;border-radius:4px"><strong>Instructions:</strong> ${instructions}</div>` : ""}
    ${sectionHTML}
    <div style="text-align:center;margin-top:20px;color:#888;font-size:12px">***************************************</div>
    <div style="text-align:center;font-size:12px;margin-top:6px">1</div>
  </div>`;
}

// ─── Marks Summary Bar ────────────────────────────────────────────────────────

function MarksSummary({ sections, maxMarks }) {
  const total = sections.reduce((s,sec)=>s + sec.questionCount*sec.marksPerQ, 0);
  const max = parseInt(maxMarks)||0;
  const diff = max - total;
  const pct = max ? Math.min(100, (total/max)*100) : 0;
  const over = total > max;

  return (
    <div style={{ background: over ? "#fef2f2" : "#f0fdf4", border:`1.5px solid ${over?"#fca5a5":"#86efac"}`, borderRadius:10, padding:"12px 16px", marginBottom:16 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}>
        <span style={{ fontSize:12, fontWeight:700, color: over?"#dc2626":"#166534" }}>
          {over ? `${total-max} marks over limit` : diff === 0 ? "Paper complete" : `${diff} marks remaining`}
        </span>
        <span style={{ fontSize:13, fontWeight:800, color: over?"#dc2626":"#166534" }}>{total} / {max}</span>
      </div>
      <div style={{ background:"#e5e7eb", borderRadius:99, height:6, overflow:"hidden" }}>
        <div style={{ background: over?"#ef4444":"#22c55e", width:`${pct}%`, height:"100%", borderRadius:99, transition:"width .3s" }} />
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function App() {
  const [view, setView] = useState("build"); // build | preview
  const [header, setHeader] = useState(DEFAULT_HEADER);
  const [sections, setSections] = useState([makeSection()]);
  const [savedDrafts, setSavedDrafts] = useState(() => {
    try { return JSON.parse(localStorage.getItem("qp_drafts")||"[]"); } catch { return []; }
  });
  const [draftName, setDraftName] = useState("");
  const [showDraftPanel, setShowDraftPanel] = useState(false);
  const logoRef = useRef();
  const previewRef = useRef();

  function addSection() { setSections(s=>[...s, makeSection()]); }

  function updateSection(i, sec) { setSections(s=>s.map((x,idx)=>idx===i?sec:x)); }

  function deleteSection(i) { setSections(s=>s.filter((_,idx)=>idx!==i)); }

  function moveSection(i, dir) {
    setSections(s=>{
      const arr=[...s];
      const j=i+dir;
      if(j<0||j>=arr.length) return arr;
      [arr[i],arr[j]]=[arr[j],arr[i]];
      return arr;
    });
  }

  function saveDraft() {
    const name = draftName.trim() || `Draft ${new Date().toLocaleDateString("en-IN")}`;
    const draft = { name, ts: Date.now(), header, sections };
    const drafts = [draft, ...savedDrafts.filter(d=>d.name!==name)].slice(0,10);
    setSavedDrafts(drafts);
    try { localStorage.setItem("qp_drafts", JSON.stringify(drafts)); } catch{}
    setDraftName("");
    alert(`Saved: "${name}"`);
  }

  function loadDraft(draft) {
    setHeader(draft.header);
    setSections(draft.sections);
    setShowDraftPanel(false);
    setView("build");
  }

  function deleteDraft(ts) {
    const drafts = savedDrafts.filter(d=>d.ts!==ts);
    setSavedDrafts(drafts);
    try { localStorage.setItem("qp_drafts", JSON.stringify(drafts)); } catch{}
  }

  function handleLogo(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => setHeader(h=>({...h, logoUrl: ev.target.result}));
    reader.readAsDataURL(file);
  }

  function newPaper() {
    if (!window.confirm("Start a new paper? Unsaved changes on this screen will be lost.")) return;
    setHeader(DEFAULT_HEADER);
    setSections([makeSection()]);
    setView("build");
    setShowDraftPanel(false);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  const previewHTML = buildPreviewHTML(header, sections);

  async function downloadPDF() {
    if (!previewRef.current) return;
    const filename = `${(header.school || "Question-Paper").replace(/[^a-z0-9]+/gi, "-")}-${(header.classVal || "").replace(/[^a-z0-9]+/gi, "-")}-${(header.subject || "").replace(/[^a-z0-9]+/gi, "-")}.pdf`.replace(/-+/g, "-");
    await html2pdf().set({
      margin: 0,
      filename,
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff" },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
      pagebreak: { mode: ["css", "legacy"] }
    }).from(previewRef.current).save();
  }

  return (
    <div style={{ minHeight:"100vh", background:"#f5f5f7", fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif" }}>

      {/* Top bar */}
      <div style={{ background:"#18181b", color:"#fff", padding:"0 20px", height:52, display:"flex", alignItems:"center", gap:12, position:"sticky", top:0, zIndex:100, boxShadow:"0 2px 8px rgba(0,0,0,.2)" }}>
        <div style={{ fontSize:15, fontWeight:800, letterSpacing:"-0.02em" }}>QP Builder</div>
        <div style={{ fontSize:11, color:"#71717a", marginRight:"auto" }}>Goodwill Public School</div>
        <button onClick={()=>setView("build")} style={{ background: view==="build"?"#6366f1":"transparent", color: view==="build"?"#fff":"#a1a1aa", border:"none", padding:"5px 12px", borderRadius:6, fontSize:12, fontWeight:600, cursor:"pointer" }}>Build</button>
        <button onClick={()=>setView("preview")} style={{ background: view==="preview"?"#6366f1":"transparent", color: view==="preview"?"#fff":"#a1a1aa", border:"none", padding:"5px 12px", borderRadius:6, fontSize:12, fontWeight:600, cursor:"pointer" }}>Preview</button>
        <button onClick={newPaper} style={{ background:"transparent", color:"#a1a1aa", border:"1px solid #3f3f46", padding:"5px 12px", borderRadius:6, fontSize:12, fontWeight:600, cursor:"pointer" }}>New Paper</button>
        <button onClick={()=>setShowDraftPanel(d=>!d)} style={{ background:"transparent", color:"#a1a1aa", border:"1px solid #3f3f46", padding:"5px 12px", borderRadius:6, fontSize:12, fontWeight:600, cursor:"pointer" }}>
          Drafts {savedDrafts.length>0 && `(${savedDrafts.length})`}
        </button>
      </div>

      {/* Draft Panel */}
      {showDraftPanel && (
        <div style={{ background:"#fff", borderBottom:"1.5px solid #e5e7eb", padding:"16px 20px" }}>
          <div style={{ maxWidth:720, margin:"0 auto" }}>
            <div style={{ fontSize:13, fontWeight:700, marginBottom:10 }}>Saved Drafts</div>
            <div style={{ display:"flex", gap:8, marginBottom:12 }}>
              <input value={draftName} onChange={e=>setDraftName(e.target.value)}
                placeholder="Draft name (optional)"
                style={{ flex:1, padding:"7px 10px", borderRadius:6, border:"1.5px solid #e5e7eb", fontSize:13, outline:"none", fontFamily:"inherit" }} />
              <Btn onClick={saveDraft} small>Save Current</Btn>
            </div>
            {savedDrafts.length === 0
              ? <div style={{ fontSize:12, color:"#9ca3af" }}>No drafts saved yet.</div>
              : savedDrafts.map(d=>(
                <div key={d.ts} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6, padding:"8px 10px", background:"#f9fafb", borderRadius:7 }}>
                  <div style={{ flex:1 }}>
                    <div style={{ fontSize:13, fontWeight:600 }}>{d.name}</div>
                    <div style={{ fontSize:11, color:"#9ca3af" }}>{new Date(d.ts).toLocaleString("en-IN")}</div>
                  </div>
                  <Btn onClick={()=>loadDraft(d)} small variant="secondary">Load</Btn>
                  <Btn onClick={()=>deleteDraft(d.ts)} small variant="danger">Delete</Btn>
                </div>
              ))
            }
          </div>
        </div>
      )}

      <div style={{ maxWidth:720, margin:"0 auto", padding:"20px 16px" }}>

        {view === "build" && (
          <>
            {/* Header Card */}
            <Card>
              <div style={{ fontSize:14, fontWeight:800, color:"#18181b", marginBottom:16 }}>Paper Details</div>
              <div className="qp-header-grid" style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
                <Input label="School Name" value={header.school} onChange={v=>setHeader(h=>({...h,school:v}))} />
                <Input label="Location" value={header.location} onChange={v=>setHeader(h=>({...h,location:v}))} />
                <Input label="Examination" value={header.exam} onChange={v=>setHeader(h=>({...h,exam:v}))} />
                <Input label="Session" value={header.session} onChange={v=>setHeader(h=>({...h,session:v}))} />
                <Input label="Class" value={header.classVal} onChange={v=>setHeader(h=>({...h,classVal:v}))} />
                <Input label="Subject" value={header.subject} onChange={v=>setHeader(h=>({...h,subject:v}))} />
                <Input label="Duration" value={header.duration} onChange={v=>setHeader(h=>({...h,duration:v}))} />
                <Input label="Max Marks" value={header.maxMarks} onChange={v=>setHeader(h=>({...h,maxMarks:v}))} />
                <Input label="Exam Date (optional)" value={header.date} onChange={v=>setHeader(h=>({...h,date:v}))} />
              </div>
              <Textarea label="Instructions (optional)" value={header.instructions}
                onChange={v=>setHeader(h=>({...h,instructions:v}))}
                placeholder="e.g. Attempt all questions. No calculators allowed." />
              <div>
                <label style={{ display:"block", fontSize:11, fontWeight:700, color:"#6b7280", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.04em" }}>School Logo (optional)</label>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  {header.logoUrl && <img src={header.logoUrl} style={{ height:40, borderRadius:4, border:"1px solid #e5e7eb" }} />}
                  <Btn onClick={()=>logoRef.current.click()} variant="secondary" small>Upload Logo</Btn>
                  {header.logoUrl && <Btn onClick={()=>setHeader(h=>({...h,logoUrl:null}))} variant="danger" small>Remove</Btn>}
                </div>
                <input ref={logoRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleLogo} />
              </div>
            </Card>

            {/* Marks summary */}
            <MarksSummary sections={sections} maxMarks={header.maxMarks} />

            {/* Sections */}
            {sections.map((sec,i)=>(
              <SectionEditor key={sec.id} section={sec} index={i} total={sections.length}
                onChange={s=>updateSection(i,s)}
                onDelete={()=>deleteSection(i)}
                onMoveUp={()=>moveSection(i,-1)}
                onMoveDown={()=>moveSection(i,1)} />
            ))}

            <button onClick={addSection}
              style={{ width:"100%", padding:"12px", background:"#fff", border:"2px dashed #c7d2fe", borderRadius:10, fontSize:13, fontWeight:700, color:"#6366f1", cursor:"pointer", marginBottom:20 }}>
              + Add Section
            </button>

            <div style={{ display:"flex", gap:8 }}>
              <Btn onClick={()=>setView("preview")} full>Preview Paper</Btn>
            </div>
          </>
        )}

        {view === "preview" && (
          <>
            <div style={{ display:"flex", gap:8, marginBottom:16, flexWrap:"wrap" }}>
              <Btn onClick={()=>setView("build")} variant="secondary">Edit</Btn>
              <Btn onClick={downloadPDF}>Download PDF</Btn>
              <Btn onClick={()=>window.print()} variant="secondary">Print</Btn>
            </div>
            <Card style={{ padding:0, overflow:"hidden" }}>
              <div ref={previewRef} dangerouslySetInnerHTML={{ __html: previewHTML }} />
            </Card>
          </>
        )}
      </div>

      <style>{`
        @media (max-width: 640px) {
          body { overflow-x: hidden; }
          button, input, textarea, select { font-size: 16px !important; }
          [style*="grid-template-columns:1fr 1fr"] { grid-template-columns: 1fr !important; }
          [style*="grid-template-columns:1fr 1fr 1fr"] { grid-template-columns: 1fr !important; }
          [style*="justify-content:space-between"] { gap: 8px; }
          .qp-mobile-scroll { overflow-x: auto; }
        }
        @media print {
          body > * { display: none !important; }
          body > div > div:last-child { display: block !important; }
          * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
        * { box-sizing: border-box; }
        input[type=number]::-webkit-inner-spin-button { opacity: 0.5; }
      `}</style>
    </div>
  );
}
