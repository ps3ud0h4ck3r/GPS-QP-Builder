import { useEffect, useMemo, useRef, useState } from "react";
import gpsLogo from "./gps-logo.png";
import html2pdf from "html2pdf.js";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType } from "docx";

const QUESTION_TYPES = [
  { id: "fill", label: "Fill in the Blanks", defaultMarks: 1, hasOptions: false, hasDiagram: true },
  { id: "mcq3", label: "MCQ (3 options)", defaultMarks: 1, hasOptions: 3, hasDiagram: true },
  { id: "mcq4", label: "MCQ (4 options)", defaultMarks: 1, hasOptions: 4, hasDiagram: true },
  { id: "tf", label: "True or False", defaultMarks: 1, hasOptions: false, hasDiagram: false },
  { id: "oneword", label: "One Word Answer", defaultMarks: 1, hasOptions: false, hasDiagram: false },
  { id: "short2", label: "Short Answer (2 marks)", defaultMarks: 2, hasOptions: false, hasDiagram: true },
  { id: "short3", label: "Short Answer (3 marks)", defaultMarks: 3, hasOptions: false, hasDiagram: true },
  { id: "long", label: "Long Answer (5 marks)", defaultMarks: 5, hasOptions: false, hasDiagram: true },
  { id: "match", label: "Match the Following", defaultMarks: 5, hasOptions: false, hasDiagram: false },
  { id: "num", label: "Numerical / Independent", defaultMarks: 4, hasOptions: false, hasDiagram: true },
  { id: "diagram", label: "Diagram Based", defaultMarks: 3, hasOptions: false, hasDiagram: true },
];

const ALPHA = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T"];
const OPT_ALPHA = ["a","b","c","d"];
const GREEK = ["α","β","γ","δ","θ","λ","μ","π","σ","φ","Δ","Ω","Σ","Φ","Γ","√","∞","≤","≥","≠","±"];
const PRINT_FONTS = { small: 13, medium: 15, large: 17 };
const AI_DAILY_LIMIT = 10;

function isJuniorClass(value="") {
  const v = String(value).trim().toUpperCase().replace(/\s+/g, " ");
  return /^(?:1|I|CLASS\s*1|CLASS\s*I|CLASS\s*ONE)$/.test(v) || /^(?:2|II|CLASS\s*2|CLASS\s*II|CLASS\s*TWO)$/.test(v);
}

const DEFAULT_HEADER = {
  school: "Goodwill Public School", location: "Uttam Nagar", exam: "Half Early Examination", session: "2026–27",
  classVal: "VI", subject: "Computer", duration: "1.5 Hours", maxMarks: "30", date: "", instructions: "", logoUrl: null, language: "English", numbering: "alpha", showStudentInfo: true,
};

function uid() { return Math.random().toString(36).slice(2, 10); }
function makeQuestion(typeId) {
  const type = QUESTION_TYPES.find(t => t.id === typeId);
  return { id: uid(), text: "", diagram: null, answerLines: 0, answerLineHeight: 40, options: type.hasOptions ? Array.from({length:type.hasOptions}, () => "") : null,
    matchLeft: typeId === "match" ? ["","","","",""] : null, matchRight: typeId === "match" ? ["","","","",""] : null };
}
function makeSection() {
  return { id: uid(), typeId: "fill", heading: "", questionCount: 1, marksPerQ: 1, questions: [makeQuestion("fill")] };
}
function totalMarks(sections) { return sections.reduce((s, sec) => s + (Number(sec.questionCount)||0) * (Number(sec.marksPerQ)||0), 0); }
function escapeHtml(v="") { return String(v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\"/g,"&quot;"); }
function plainFromHtml(html="") {
  const doc = new DOMParser().parseFromString(`<div>${html}</div>`, "text/html");
  return doc.body.textContent || "";
}
function hasHtml(v="") { return /<([a-z][\s\S]*?)>/i.test(v); }

const Input = ({ label, value, onChange, placeholder, type="text", small, disabled }) => (
  <div style={{marginBottom:small?8:14}}>
    {label && <label className="field-label">{label}</label>}
    <input disabled={disabled} type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} className="field-input" style={{padding:small?"7px 10px":"9px 12px",fontSize:small?13:14}} />
  </div>
);
const Textarea = ({ label, value, onChange, placeholder, rows=2 }) => (
  <div style={{marginBottom:14}}>
    {label && <label className="field-label">{label}</label>}
    <textarea value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} rows={rows} className="field-input field-textarea" />
  </div>
);
const Btn = ({ onClick, children, variant="primary", small, disabled, full, title }) => <button title={title} onClick={onClick} disabled={disabled} className={`btn btn-${variant} ${small?"btn-small":""} ${full?"btn-full":""}`}>{children}</button>;
const Card = ({ children, style, className="" }) => <div className={`card ${className}`} style={style}>{children}</div>;
const Badge = ({ children, color="var(--badge-bg)", textColor="var(--badge-text)" }) => <span className="badge" style={{background:color,color:textColor}}>{children}</span>;

function RichTextEditor({ value, onChange, placeholder }) {
  const ref = useRef(null);
  const lastValue = useRef(value);
  useEffect(() => {
    if (!ref.current || lastValue.current === value) return;
    ref.current.innerHTML = hasHtml(value) ? value : escapeHtml(value).replace(/\n/g,"<br>");
    lastValue.current = value;
  }, [value]);
  useEffect(() => {
    if (ref.current && ref.current.innerHTML === "") ref.current.dataset.placeholder = placeholder || "";
  }, [placeholder]);
  function emit() {
    if (!ref.current) return;
    const html = ref.current.innerHTML;
    lastValue.current = html;
    onChange(html);
  }
  function command(cmd, arg=null) { ref.current?.focus(); document.execCommand(cmd, false, arg); emit(); }
  function insert(text) { ref.current?.focus(); document.execCommand("insertText", false, text); emit(); }
  function insertTable() {
    ref.current?.focus();
    const html = `<table class="question-table"><tbody>${Array.from({length:3},(_,r)=>`<tr>${Array.from({length:3},(_,c)=>`<td>${r===0?`Heading ${c+1}`:""}</td>`).join("")}</tr>`).join("")}</tbody></table><br>`;
    document.execCommand("insertHTML", false, html); emit();
  }
  return <div className="rich-wrap">
    <div className="rich-toolbar">
      <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("bold")}><b>B</b></button>
      <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("italic")}><i>I</i></button><button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("underline")}><u>U</u></button><button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("justifyLeft")}>≡</button><button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("justifyCenter")}>☰</button><button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("justifyRight")}>≡</button><button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("insertUnorderedList")}>• List</button><button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("insertOrderedList")}>1. List</button>
      <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("superscript")}>x<sup>2</sup></button>
      <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("subscript")}>x<sub>2</sub></button>
      <select onChange={e=>{ if(e.target.value) insert(e.target.value); e.target.value=""; }} defaultValue="" aria-label="Greek symbols">
        <option value="">Ω Greek / symbols</option>{GREEK.map(x=><option key={x} value={x}>{x}</option>)}
      </select>
      <button type="button" onMouseDown={e=>e.preventDefault()} onClick={insertTable}>▦ Table</button>
    </div>
    <div ref={ref} contentEditable suppressContentEditableWarning className="rich-editor" data-placeholder={placeholder||""} onInput={emit} onBlur={emit} />
  </div>;
}

function DrawCanvas({ onSave, onCancel }) {
  const canvasRef=useRef(); const drawing=useRef(false); const last=useRef({x:0,y:0});
  function pos(e){const r=canvasRef.current.getBoundingClientRect(),s=e.touches?e.touches[0]:e;return{x:s.clientX-r.left,y:s.clientY-r.top};}
  function start(e){e.preventDefault();drawing.current=true;last.current=pos(e);}
  function move(e){e.preventDefault();if(!drawing.current)return;const c=canvasRef.current,ctx=c.getContext("2d"),p=pos(e);ctx.beginPath();ctx.moveTo(last.current.x,last.current.y);ctx.lineTo(p.x,p.y);ctx.strokeStyle="#1e293b";ctx.lineWidth=2;ctx.lineCap="round";ctx.stroke();last.current=p;}
  function clear(){canvasRef.current.getContext("2d").clearRect(0,0,600,300);}
  return <div className="draw-box"><div className="small-muted">Draw Diagram</div><canvas ref={canvasRef} width="560" height="240" className="draw-canvas" onMouseDown={start} onMouseMove={move} onMouseUp={()=>drawing.current=false} onMouseLeave={()=>drawing.current=false} onTouchStart={start} onTouchMove={move} onTouchEnd={()=>drawing.current=false}/><div className="toolbar-row"><Btn onClick={clear} variant="secondary" small>Clear</Btn><Btn onClick={()=>onSave(canvasRef.current.toDataURL("image/png"))} small>Use this diagram</Btn><Btn onClick={onCancel} variant="ghost" small>Cancel</Btn></div></div>;
}
function DiagramPicker({ diagram, onChange }) {
  const [mode,setMode]=useState(null); const fileRef=useRef();
  function upload(e){const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>onChange({mode:"upload",dataUrl:ev.target.result});r.readAsDataURL(f);setMode(null);}
  if(diagram)return <div className="diagram-preview">{diagram.mode==="space"?<div className="diagram-space">Diagram space will appear in print</div>:<img src={diagram.dataUrl} alt="diagram"/>}<div className="toolbar-row"><Btn onClick={()=>onChange(null)} variant="danger" small>Remove Diagram</Btn></div></div>;
  if(mode==="draw")return <DrawCanvas onSave={dataUrl=>{onChange({mode:"draw",dataUrl});setMode(null)}} onCancel={()=>setMode(null)}/>;
  return <div className="toolbar-row"><Btn onClick={()=>fileRef.current.click()} variant="secondary" small>Upload Image</Btn><Btn onClick={()=>setMode("draw")} variant="secondary" small>Draw Here</Btn><Btn onClick={()=>onChange({mode:"space",dataUrl:null})} variant="secondary" small>Leave Space</Btn><input ref={fileRef} type="file" accept="image/*" hidden onChange={upload}/></div>;
}

function QuestionEditor({ q,index,typeId,onChange,showDiagram,onMoveUp,onMoveDown,onDelete,onReorder,onSaveBank,numbering,junior }) {
  const type=QUESTION_TYPES.find(t=>t.id===typeId); const [showDiag,setShowDiag]=useState(!!q.diagram);
  const upd=p=>onChange({...q,...p});
  return <div className={`question-editor ${q.pageBreakAfter?"question-break-after":""}`} data-tour="tour-question" draggable onDragStart={e=>e.dataTransfer.setData("text/plain",q.id)} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const source=e.dataTransfer.getData("text/plain");if(source&&source!==q.id)onReorder(source,q.id)}}>
    <div className="question-editor-head" data-tour="tour-question-tools"><span className="drag-handle" title="Drag to reorder">⠿</span><Badge color="#ede9fe" textColor="#5b21b6">{ALPHA[index]}</Badge><span className="small-muted">{type.label}</span><span style={{marginLeft:"auto"}}><Btn onClick={onMoveUp} variant="ghost" small title="Move up">↑</Btn><Btn onClick={onMoveDown} variant="ghost" small title="Move down">↓</Btn><Btn onClick={onSaveBank} variant="ghost" small title="Save to Question Bank">▦</Btn><Btn onClick={onDelete} variant="danger" small title="Delete question">✕</Btn></span></div>
    {typeId==="match" ? <div className="two-col"><div><div className="field-label">Column A</div>{q.matchLeft.map((v,i)=><input key={i} value={v} placeholder={`${i+1}.`} onChange={e=>{const a=[...q.matchLeft];a[i]=e.target.value;upd({matchLeft:a})}} className="field-input compact"/>)}</div><div><div className="field-label">Column B</div>{q.matchRight.map((v,i)=><input key={i} value={v} placeholder={`${String.fromCharCode(97+i)}.`} onChange={e=>{const a=[...q.matchRight];a[i]=e.target.value;upd({matchRight:a})}} className="field-input compact"/>)}</div></div> : <>
      <RichTextEditor value={q.text} onChange={text=>upd({text})} placeholder={typeId==="fill"?"Write statement with ___ for blank":typeId==="tf"?"Write a true/false statement":"Write your question here"}/>
      {type.hasOptions && <div style={{marginTop:8}}>{Array.from({length:type.hasOptions},(_,i)=><div className="option-row" key={i}><span>{OPT_ALPHA[i]})</span><input value={q.options[i]} onChange={e=>{const a=[...q.options];a[i]=e.target.value;upd({options:a})}} placeholder={`Option ${OPT_ALPHA[i].toUpperCase()}`} className="field-input compact"/></div>)}</div>}
    </>}
    {showDiagram && type.hasDiagram && <div style={{marginTop:8}}>{!showDiag&&!q.diagram?<button className="link-btn" onClick={()=>setShowDiag(true)}>+ Add Diagram</button>:<DiagramPicker diagram={q.diagram} onChange={d=>{upd({diagram:d});if(!d)setShowDiag(false)}}/>}</div>}
    {junior&&<div className="junior-answer-controls">
      <label className="check-row"><input type="checkbox" checked={Number(q.answerLines||0)>0} onChange={e=>upd({answerLines:e.target.checked?3:0})}/> Add Lines for Student Answer</label>
      {Number(q.answerLines||0)>0&&<><label className="field-label" style={{display:"inline-flex",alignItems:"center",gap:8,marginLeft:24}}>Number of Lines <select value={q.answerLines} onChange={e=>upd({answerLines:Math.max(1,Math.min(12,parseInt(e.target.value)||1))})} className="field-input compact" style={{width:90,marginBottom:0}}>{Array.from({length:12},(_,i)=><option key={i+1} value={i+1}>{i+1}</option>)}</select></label><label className="field-label" style={{display:"inline-flex",alignItems:"center",gap:8,marginLeft:24}}>Line Height (px) <input type="number" min="1" step="1" value={q.answerLineHeight ?? 40} onChange={e=>upd({answerLineHeight:e.target.value===""?"":Math.max(1,parseInt(e.target.value)||1)})} className="field-input compact" style={{width:100,marginBottom:0}}/></label></>}
    </div>}
    <label className="check-row"><input type="checkbox" checked={!!q.pageBreakAfter} onChange={e=>upd({pageBreakAfter:e.target.checked})}/> Start a new A4 page after this question</label>
  </div>;
}

function SectionEditor({ section,index,total,onChange,onDelete,onMoveUp,onMoveDown,onDuplicate,onSaveBank,maxAllowed,notify,language,numbering,junior }) {
  const [open,setOpen]=useState(true); const [aiLoading,setAiLoading]=useState(false); const [aiTopic,setAiTopic]=useState("");
  const [aiUses,setAiUses]=useState(()=>{try{const d=JSON.parse(localStorage.getItem("qp_ai_usage")||"{}");return d.date===new Date().toISOString().slice(0,10)?Number(d.count)||0:0}catch{return 0}});
  const type=QUESTION_TYPES.find(t=>t.id===section.typeId); const sectionMarks=section.questionCount*section.marksPerQ;
  const upd=p=>onChange({...section,...p});
  function changeType(id){const t=QUESTION_TYPES.find(x=>x.id===id);const maxCount=Math.max(1,Math.floor(maxAllowed/t.defaultMarks)||1);const count=Math.min(section.questionCount,maxCount);upd({typeId:id,marksPerQ:t.defaultMarks,questionCount:count,questions:Array.from({length:count},()=>makeQuestion(id))});}
  function changeCount(n){if(String(n)===""){upd({questionCount:""});return;}const desired=Math.max(1,Math.min(20,parseInt(n)||1));const maxCount=Math.max(1,Math.floor(maxAllowed/(Number(section.marksPerQ)||1))||1);if(desired*(Number(section.marksPerQ)||0)>maxAllowed){notify(`This section can have at most ${maxCount} question${maxCount===1?"":"s"}. Maximum paper marks are ${maxAllowed} here.`);return;}const qs=[...section.questions];while(qs.length<desired)qs.push(makeQuestion(section.typeId));upd({questionCount:desired,questions:qs.slice(0,desired)});}
  function changeMarks(n){if(String(n)===""){upd({marksPerQ:""});return;}const m=Math.max(1,Math.min(20,parseInt(n)||1));if((Number(section.questionCount)||0)*m>maxAllowed){notify(`These marks would exceed the remaining ${maxAllowed} marks. Reduce questions or marks per question.`);return;}upd({marksPerQ:m});}
  function updateQuestion(i,q){const qs=[...section.questions];qs[i]=q;upd({questions:qs});}
  function reorderQuestion(sourceId,targetId){const qs=[...section.questions],from=qs.findIndex(x=>x.id===sourceId),to=qs.findIndex(x=>x.id===targetId);if(from<0||to<0||from===to)return;const [moving]=qs.splice(from,1);qs.splice(to,0,moving);upd({questions:qs});}
  async function generateAI(){if(!aiTopic.trim())return;if(aiUses>=AI_DAILY_LIMIT){notify(`You have reached the ${AI_DAILY_LIMIT}-generation daily limit on this device.`);return;}setAiLoading(true);try{const count=section.questionCount;const prompt=`Generate ${count} "${type.label}" questions for a school exam on topic: "${aiTopic}". Use clear, age-appropriate school language. Write the questions in ${language}. Do not repeat questions. Return ONLY valid JSON array, no markdown. ${section.typeId.startsWith("mcq")?`[{"text":"Question?","options":["opt1","opt2","opt3"${section.typeId==="mcq4"?',"opt4"':''}]}]`:section.typeId==="match"?'[{"matchLeft":["term1","term2","term3","term4","term5"],"matchRight":["def1","def2","def3","def4","def5"]}]':'[{"text":"Question or statement here"}]'} Generate exactly ${count} items.`;const res=await fetch("/api/generate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({prompt})});const data=await res.json().catch(()=>({}));if(!res.ok)throw new Error(data.error||`Request failed (${res.status})`);const clean=(data.content||"").replace(/```json|```/g,"").trim();const a=clean.indexOf("["),b=clean.lastIndexOf("]");const parsed=JSON.parse(a>=0&&b>a?clean.slice(a,b+1):clean);if(!Array.isArray(parsed))throw new Error("AI returned an invalid list.");const qs=parsed.slice(0,count).map(item=>({id:uid(),text:item.text||"",diagram:null,answerLines:0,answerLineHeight:40,options:item.options|| (type.hasOptions?Array.from({length:type.hasOptions},()=>""):null),matchLeft:item.matchLeft|| (section.typeId==="match"?["","","","",""]:null),matchRight:item.matchRight|| (section.typeId==="match"?["","","","",""]:null)}));upd({questions:qs});const next=aiUses+1;setAiUses(next);localStorage.setItem("qp_ai_usage",JSON.stringify({date:new Date().toISOString().slice(0,10),count:next}));}catch(e){notify(`AI generation failed. ${e.message||"Please fill manually."}`)}setAiLoading(false);}
  return <Card className="section-card" style={{marginBottom:16}}>
    <div className="section-head" data-tour="tour-section-settings"><div className="drag-handle section-drag" draggable onDragStart={e=>e.dataTransfer.setData("section-id",section.id)} title="Drag to reorder section">⠿</div><div className="section-arrows"><button onClick={onMoveUp} disabled={index===0}>▲</button><button onClick={onMoveDown} disabled={index===total-1}>▼</button></div><Badge color="#fef3c7" textColor="#92400e">Q{index+1}</Badge><div className="section-title">{section.heading||type.label}</div><Badge color="#f0fdf4" textColor="#166534">{sectionMarks} marks</Badge><Btn onClick={onDuplicate} variant="secondary" small title="Duplicate this section">Duplicate</Btn><button className="icon-btn" onClick={()=>setOpen(o=>!o)}>{open?"▾":"▸"}</button><button className="icon-btn danger-text" onClick={onDelete}>✕</button></div>
    {open&&<>
      <div className="qp-section-grid section-config"><div><label className="field-label">Type</label><select value={section.typeId} onChange={e=>changeType(e.target.value)} className="field-input"><option value="fill">Fill in the Blanks</option>{QUESTION_TYPES.slice(1).map(t=><option key={t.id} value={t.id}>{t.label}</option>)}</select></div><div><label className="field-label">No. of Questions</label><input type="number" min="1" max="20" value={section.questionCount} onChange={e=>changeCount(e.target.value)} className="field-input"/></div><div><label className="field-label">Marks/Question</label><input type="number" min="1" max="20" value={section.marksPerQ} onChange={e=>changeMarks(e.target.value)} className="field-input"/></div><div><label className="field-label">Page Break</label><select value={section.pageBreak||"auto"} onChange={e=>upd({pageBreak:e.target.value})} className="field-input"><option value="auto">Automatic</option><option value="before">Start section on new A4 page</option></select></div></div>
      <div style={{marginBottom:14}}><label className="field-label">Custom Heading (optional)</label><input value={section.heading} onChange={e=>upd({heading:e.target.value})} placeholder='e.g. "Section A" or leave blank to use type name' className="field-input"/></div>
      <div>{section.questions.map((q,i)=><QuestionEditor key={q.id} q={q} index={i} typeId={section.typeId} onChange={nq=>updateQuestion(i,nq)} onReorder={reorderQuestion} onSaveBank={()=>onSaveBank?.(q,section)} showDiagram numbering={numbering} junior={junior} onMoveUp={()=>{if(i>0){const qs=[...section.questions];[qs[i-1],qs[i]]=[qs[i],qs[i-1]];upd({questions:qs})}}} onMoveDown={()=>{if(i<section.questions.length-1){const qs=[...section.questions];[qs[i+1],qs[i]]=[qs[i],qs[i+1]];upd({questions:qs})}}} onDelete={()=>{if(section.questions.length===1){notify("A section must contain at least one question.");return;}upd({questions:section.questions.filter(x=>x.id!==q.id),questionCount:section.questionCount-1})}}/>)}</div>
    </>}
  </Card>;
}

function formatInstructions(text){if(!text)return"";const normalized=String(text).replace(/\r\n/g,"\n").trim();if(!normalized)return"";const lines=normalized.replace(/\s+(?=\d+\.\s)/g,"\n").split(/\n+/).map(x=>x.trim()).filter(Boolean);return lines.map(x=>`<div style="margin:2px 0">${escapeHtml(x)}</div>`).join("");}
function renderQuestionText(text){if(!text)return"___________________________________";return hasHtml(text)?text:escapeHtml(text).replace(/\n/g,"<br>");}
function questionLabel(index, style="alpha") {
  if(style === "numeric") return `${index+1}.`;
  if(style === "roman") { const vals=["i","ii","iii","iv","v","vi","vii","viii","ix","x","xi","xii","xiii","xiv","xv","xvi","xvii","xviii","xix","xx"]; return `(${vals[index]||index+1})`; }
  return `${ALPHA[index]||index+1}.`;
}
function buildPreviewHTML(header,sections,fontSize=15,zoom=1){
  const {school,location,exam,session,classVal,subject,duration,maxMarks,date,instructions,logoUrl,numbering,language,showStudentInfo}=header;
  const junior=isJuniorClass(classVal);
  const labels=language==="Hindi"?{time:"समय",subject:"विषय",marks:"पूर्णांक",instructions:"निर्देश"}:language==="Sanskrit"?{time:"कालः",subject:"विषयः",marks:"पूर्णाङ्काः",instructions:"निर्देशाः"}:{time:"Time",subject:"Subject",marks:"M.M.",instructions:"Instructions"};
  const sectionHTML=sections.map((sec,si)=>{
    const type=QUESTION_TYPES.find(t=>t.id===sec.typeId),heading=sec.heading||type.label,total=sec.questionCount*sec.marksPerQ;
    const marksStr=sec.marksPerQ===1?`(${total})`:`(${sec.marksPerQ} × ${sec.questionCount} = ${total})`;
    const qHTML=sec.questions.map((q,qi)=>{
      let inner="";
      if(sec.typeId==="match") inner=`<table class="preview-table"><tr><th>Column A</th><th>Column B</th></tr>${(q.matchLeft||[]).map((l,i)=>`<tr><td>${i+1}. ${escapeHtml(l||"_______")}</td><td>${String.fromCharCode(97+i)}. ${escapeHtml((q.matchRight||[])[i]||"_______")}</td></tr>`).join("")}</table>`;
      else { inner=`<div>${renderQuestionText(q.text)}</div>`; if(sec.typeId==="tf") inner+=` <span style="margin-left:24px">__________</span>`; if(q.options) inner+=`<div class="preview-options">${q.options.map((o,oi)=>`<div>${OPT_ALPHA[oi]}) ${escapeHtml(o||"_______")}</div>`).join("")}</div>`; if(q.diagram) inner+=q.diagram.mode==="space"?`<div class="preview-diagram-space">Diagram</div>`:q.diagram.dataUrl?`<div class="preview-diagram"><img src="${q.diagram.dataUrl}"/></div>`:""; }
      const answerLines=junior&&Number(q.answerLines||0)>0?`<div class="answer-lines" style="margin-top:10px;">${Array.from({length:Math.min(12,Number(q.answerLines)||0)},()=>`<div style="height:${Math.max(1,Number(q.answerLineHeight)||60)}px;border-bottom:1px solid #222;margin:0 0 2px;"></div>`).join("")}</div>`:"";
      return `<div class="preview-question ${q.pageBreakAfter?"question-break-after":""}"><span class="question-label">${questionLabel(qi,numbering)}</span><div class="question-body">${inner}${answerLines}</div></div>`;
    }).join("");
    return `<div class="preview-section ${sec.pageBreak==="before"&&si>0?"section-break-before":""}"><div class="preview-section-head"><span>Q${si+1}. ${escapeHtml(heading)}</span><span>${marksStr}</span></div>${qHTML}</div>`;
  }).join("");
  const studentInfo=junior&&showStudentInfo!==false?`<div class="student-info-row" style="display:flex;justify-content:space-between;gap:24px;margin:8px 0 14px;font-size:var(--paper-font);"><span>Name: <span style="display:inline-block;min-width:240px;border-bottom:1px solid #222;"></span></span><span>Roll No.: <span style="display:inline-block;min-width:120px;border-bottom:1px solid #222;"></span></span></div>`:"";
  return `<div class="paper" style="--paper-font:${fontSize}px;transform:scale(${zoom});transform-origin:top center"><div class="paper-header">${logoUrl?`<img class="paper-logo" src="${logoUrl}"/>`:""}<div class="school-name">${escapeHtml(school)}</div><div class="school-location">${escapeHtml(location)}</div><div class="exam-line">${escapeHtml(exam)} ${escapeHtml(session)}</div><div class="class-line">CLASS ${escapeHtml(classVal)}</div>${date?`<div class="date-line">Date: ${escapeHtml(date)}</div>`:""}</div>${studentInfo}<div class="meta-row"><span>${labels.time}: ${escapeHtml(duration)}</span><span>${labels.subject}: ${escapeHtml(subject)}</span><span>${labels.marks}: ${escapeHtml(maxMarks)}</span></div>${instructions?`<div class="instructions"><strong>${labels.instructions}:</strong>${formatInstructions(instructions)}</div>`:""}${sectionHTML}<div class="paper-end">***************************************<br/><span>1</span></div></div>`;
}

function Blueprint({sections,maxMarks}){
  const max=parseInt(maxMarks)||0,total=totalMarks(sections);
  return <Card className="blueprint-card"><div className="blueprint-head"><div><div className="card-title" style={{marginBottom:3}}>Automatic Blueprint / Marks Distribution</div><div className="small-muted">Live distribution based on the current paper structure.</div></div><Badge color="var(--gold-soft)" textColor="var(--burgundy)">{total}/{max} marks</Badge></div><div className="blueprint-table-wrap"><table className="blueprint-table"><thead><tr><th>Section</th><th>Type</th><th>Questions</th><th>Marks/Q</th><th>Total</th><th>Share</th></tr></thead><tbody>{sections.map((s,i)=>{const t=QUESTION_TYPES.find(x=>x.id===s.typeId),m=s.questionCount*s.marksPerQ,pct=max?Math.round(m/max*100):0;return <tr key={s.id}><td>Q{i+1}</td><td>{s.heading||t.label}</td><td>{s.questionCount}</td><td>{s.marksPerQ}</td><td><b>{m}</b></td><td>{pct}%</td></tr>})}</tbody><tfoot><tr><td colSpan="4">Total</td><td>{total}</td><td>{max?Math.round(total/max*100):0}%</td></tr></tfoot></table></div></Card>;
}

function MarksSummary({sections,maxMarks}){const total=totalMarks(sections),max=parseInt(maxMarks)||0,diff=max-total,over=total>max,pct=max?Math.min(100,total/max*100):0;return <div className={`marks-summary ${over?"over":""}`}><div><span>{over?`${total-max} marks over limit`:diff===0?"Paper complete":`${diff} marks remaining`}</span><b>{total} / {max}</b></div><div className="progress"><div style={{width:`${pct}%`}}/></div></div>}

function htmlToDocxRuns(html=""){
  const doc=new DOMParser().parseFromString(`<div>${html}</div>`,"text/html");
  const out=[];
  function walk(node,styles={}){node.childNodes.forEach(n=>{if(n.nodeType===Node.TEXT_NODE){if(n.textContent)out.push(new TextRun({text:n.textContent,bold:!!styles.bold,italics:!!styles.italics,superScript:!!styles.sup,subScript:!!styles.sub,underline:!!styles.underline}));}else if(n.nodeType===Node.ELEMENT_NODE){const tag=n.tagName.toLowerCase();if(tag==="br"){out.push(new TextRun({text:"\n"}));return;}walk(n,{bold:styles.bold||tag==="b"||tag==="strong",italics:styles.italics||tag==="i"||tag==="em",sup:styles.sup||tag==="sup",sub:styles.sub||tag==="sub",underline:styles.underline||tag==="u"});}})}
  walk(doc.body); return out.length?out:[new TextRun("")];
}
function makeDocParagraph(html,opts={}){return new Paragraph({children:htmlToDocxRuns(html),spacing:{after:opts.after??100},alignment:opts.alignment});}
async function createDocxBlob(header,sections){const children=[];if(header.logoUrl){/* DOCX cannot reliably embed data URLs in all browsers without extra conversion, so the school name remains as the text header. */}children.push(new Paragraph({children:[new TextRun({text:header.school||"Question Paper",bold:true,size:30})],alignment:AlignmentType.CENTER,spacing:{after:50}}));if(header.location)children.push(new Paragraph({children:[new TextRun({text:header.location,size:20})],alignment:AlignmentType.CENTER,spacing:{after:50}}));children.push(new Paragraph({children:[new TextRun({text:`${header.exam||""} ${header.session||""}`,bold:true,size:22})],alignment:AlignmentType.CENTER,spacing:{after:50}}));children.push(new Paragraph({children:[new TextRun({text:`CLASS ${header.classVal||""}`,bold:true,size:20})],alignment:AlignmentType.CENTER,spacing:{after:120}}));if(isJuniorClass(header.classVal)&&header.showStudentInfo!==false) children.push(new Paragraph({children:[new TextRun({text:"Name: ________________________________    Roll No.: ____________",size:20})],spacing:{after:140}}));children.push(new Paragraph({children:[new TextRun({text:`Time: ${header.duration||""}     Subject: ${header.subject||""}     M.M.: ${header.maxMarks||""}`,bold:true,size:20})],spacing:{after:180}}));if(header.instructions){children.push(new Paragraph({children:[new TextRun({text:"Instructions:",bold:true,size:20})],spacing:{after:40}}));String(header.instructions).replace(/\r\n/g,"\n").replace(/\s+(?=\d+\.\s)/g,"\n").split(/\n+/).filter(Boolean).forEach(x=>children.push(new Paragraph({children:[new TextRun({text:x.trim(),size:20})],spacing:{after:40}})))}sections.forEach((sec,si)=>{const type=QUESTION_TYPES.find(t=>t.id===sec.typeId),total=sec.questionCount*sec.marksPerQ,marks=sec.marksPerQ===1?`(${total})`:`(${sec.marksPerQ}x${sec.questionCount}=${total})`;children.push(new Paragraph({children:[new TextRun({text:`Q${si+1}. ${sec.heading||type.label}`,bold:true,size:22}),new TextRun({text:`\t${marks}`,bold:true,size:22})],pageBreakBefore:si>0 && sec.pageBreak==="before",spacing:{before:160,after:100}}));sec.questions.forEach((q,qi)=>{children.push(new Paragraph({children:[new TextRun({text:`${questionLabel(qi,header.numbering||"alpha")} `,bold:true}),...htmlToDocxRuns(q.text||"___________________________________")],pageBreakBefore:si>0 && !!sec.questions[qi-1]?.pageBreakAfter,spacing:{after:80}}));if(Number(q.answerLines||0)>0&&isJuniorClass(header.classVal)){for(let li=0;li<Math.min(12,Number(q.answerLines)||0);li++)children.push(new Paragraph({children:[new TextRun({text:"______________________________________________",size:20})],spacing:{after:90}}));}if(q.options)q.options.forEach((o,i)=>children.push(new Paragraph({children:[new TextRun({text:`${OPT_ALPHA[i]}) ${plainFromHtml(o||"_______")}`})],indent:{left:360},spacing:{after:40}})));if(sec.typeId==="match"){children.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[new TableRow({children:[new TableCell({children:[new Paragraph({text:"Column A"})]}),new TableCell({children:[new Paragraph({text:"Column B"})]})]}),...(q.matchLeft||[]).map((l,i)=>new TableRow({children:[new TableCell({children:[new Paragraph({text:`${i+1}. ${l||"_______"}`})]}),new TableCell({children:[new Paragraph({text:`${String.fromCharCode(97+i)}. ${(q.matchRight||[])[i]||"_______"}`})]})]}))]}));}})});const doc=new Document({sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:720,right:720,bottom:720,left:720}}},children}]});return Packer.toBlob(doc);}


const PAPER_TEMPLATES = [
  {id:"unit-test",name:"Unit Test",description:"Compact 20-mark classroom test",header:{exam:"Unit Test",duration:"45 Minutes",maxMarks:"20"},pattern:[["short2",5,2],["fill",5,1]]},
  {id:"periodic",name:"Periodic Test",description:"Balanced 30-mark periodic assessment",header:{exam:"Periodic Test",duration:"1 Hour",maxMarks:"30"},pattern:[["mcq4",5,1],["short2",5,2],["short3",3,3]]},
  {id:"half-yearly",name:"Half Yearly",description:"Full 50-mark examination structure",header:{exam:"Half Yearly Examination",duration:"2 Hours",maxMarks:"50"},pattern:[["mcq4",10,1],["short2",5,2],["short3",5,3],["long",3,5]]},
  {id:"annual",name:"Annual Examination",description:"Full 80-mark annual paper",header:{exam:"Annual Examination",duration:"3 Hours",maxMarks:"80"},pattern:[["mcq4",10,1],["fill",10,1],["short2",5,2],["short3",5,3],["long",5,5]]}
];

function clonePaper(value){return JSON.parse(JSON.stringify(value));}
function templateToPaper(template,baseHeader){
  const header={...DEFAULT_HEADER,...baseHeader,...(template.header||{})};
  const sections=(template.pattern||[]).map(([typeId,count,marks])=>{
    const type=QUESTION_TYPES.find(x=>x.id===typeId)||QUESTION_TYPES[0];
    return {id:uid(),typeId,heading:"",questionCount:count,marksPerQ:marks,questions:Array.from({length:count},()=>makeQuestion(type.id))};
  });
  return {header,sections};
}
function paginatePaperSections(sections,maxUnits=8){
  const pages=[]; let page=[]; let units=0;
  const pushPage=()=>{if(page.length)pages.push(page);page=[];units=0};
  sections.forEach(sec=>{
    const questions=sec.questions||[];
    if(sec.pageBreak==="before"&&page.length)pushPage();
    if(!questions.length){if(units+1>maxUnits&&page.length)pushPage();page.push(sec);units+=1;return;}
    let start=0;
    while(start<questions.length){
      const capacity=Math.max(1,maxUnits-units-1);
      const take=Math.min(capacity,questions.length-start);
      const part={...clonePaper(sec),id:uid(),questions:questions.slice(start,start+take),questionCount:take};
      page.push(part);units+=1+take;start+=take;
      if(start<questions.length||questions[start-1]?.pageBreakAfter)pushPage();
    }
  });
  pushPage();
  return pages.length?pages:[[]];
}

function BuilderApp({user,initialWorkspace,onWorkspaceChange,onLogout,onAccount,onAdmin}){
  const ws=initialWorkspace||{}; const [view,setView]=useState("build"); const [tourStep,setTourStep]=useState(null); const [header,setHeader]=useState({...DEFAULT_HEADER,...(ws.current?.header||{})}); const [sections,setSections]=useState(ws.current?.sections?.length?ws.current.sections:[makeSection()]); const [savedDrafts,setSavedDrafts]=useState(ws.drafts||[]); const [templates,setTemplates]=useState(ws.templates?.length?ws.templates:[]); const [draftName,setDraftName]=useState(""); const [showDraftPanel,setShowDraftPanel]=useState(false); const [showTemplates,setShowTemplates]=useState(false); const [showWizard,setShowWizard]=useState(false); const [wizardStep,setWizardStep]=useState(0); const [wizard,setWizard]=useState({classVal:DEFAULT_HEADER.classVal,subject:DEFAULT_HEADER.subject,exam:"Half Yearly Examination",maxMarks:"30",duration:"1.5 Hours",pattern:"balanced"}); const [recovery,setRecovery]=useState(null); const [lastSavedAt,setLastSavedAt]=useState(Date.now()); const [dark,setDark]=useState(!!ws.settings?.dark); const [printFont,setPrintFont]=useState(ws.settings?.printFont||"medium"); const [zoom,setZoom]=useState(ws.settings?.zoom||1); const [notice,setNotice]=useState(""); const [previewPage,setPreviewPage]=useState(0); const logoRef=useRef(); const previewRef=useRef();
  useEffect(()=>{document.documentElement.classList.toggle("dark",dark)},[dark]);
  useEffect(()=>{const h=e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="k"){e.preventDefault();setCommandOpen(v=>!v)}if(e.key==="Escape"){setCommandOpen(false);setDesignOpen(false);setBankOpen(false)}};window.addEventListener("keydown",h);return()=>window.removeEventListener("keydown",h)},[]);

  useEffect(()=>{
    const key="gps_qp_autosave_"+String(user?.id||user?.username||"default");
    try{
      const raw=localStorage.getItem(key); const local=raw?JSON.parse(raw):null;
      const serverTs=Number(ws.current?.updatedAt||0);
      if(local&&Number(local.ts||0)>serverTs+1500) setRecovery(local);
    }catch{}
  },[]);
  useEffect(()=>{
    const snapshot={ts:Date.now(),header,sections}; setLastSavedAt(snapshot.ts);
    const key="gps_qp_autosave_"+String(user?.id||user?.username||"default");
    try{localStorage.setItem(key,JSON.stringify(snapshot));}catch{}
    const t=setTimeout(()=>onWorkspaceChange({current:{header,sections,updatedAt:snapshot.ts},drafts:savedDrafts,templates,settings:{dark,printFont,zoom}}),700);
    return()=>clearTimeout(t)
  },[header,sections,savedDrafts,templates,dark,printFont,zoom]);
  function notify(msg){setNotice(msg);window.clearTimeout(notify.t);notify.t=window.setTimeout(()=>setNotice(""),3500)}
  const tourSteps=[
    {target:"tour-create",title:"Create Paper",text:"This is your main workspace. Start here whenever you want to create or edit a question paper."},
    {target:"tour-paper-details",title:"Paper Details",text:"Enter the paper's basic information here, such as class, subject, total marks, title, date, duration and language."},
    {target:"tour-student-info",title:"Student Information",text:"For Class 1 and 2, use this switch to show or hide Student Name and Roll Number on the printed paper."},
    {target:"tour-instructions",title:"Instructions",text:"Add exam instructions here. Keep them short and clear, for example: Attempt all questions."},
    {target:"tour-marks",title:"Marks Tracker",text:"This live tracker shows how many marks your paper currently uses and how many remain."},
    {target:"tour-add-questions",title:"Add Questions",text:"This is where you build the actual question paper. Add a question section, then configure its type, count and marks."},
    {target:"tour-section-settings",title:"Section Settings",text:"Choose the question type, number of questions, marks per question and page-break behavior."},
    {target:"tour-question",title:"Question Editor",text:"Type the question here. Depending on the question type, you can also add options, matching columns, diagrams and answer lines."},
    {target:"tour-question-tools",title:"Question Tools",text:"Use these controls to reorder, move or delete a question. You can also drag questions into a different order."},
    {target:"tour-preview",title:"Live Preview",text:"Watch the paper update here while you build it. This preview shows the actual printed-paper structure."},
    {target:"tour-full-preview",title:"Full Preview",text:"Open the complete A4 paper in a dedicated preview screen before exporting or printing."},
    {target:"tour-exports",title:"Export",text:"When everything looks correct, export the paper as PDF or DOCX, or use Full Preview to print it."},
    {target:"tour-help",title:"You Are Ready",text:"That's the complete workflow. You can open this guide again any time from Need Help."}
  ];
  function startTour(){setTourStep(0)}
  function finishTour(){setTourStep(null)}
  useEffect(()=>{if(tourStep===null)return;const step=tourSteps[tourStep];const el=document.querySelector('[data-tour="'+step.target+'"]');if(el){el.scrollIntoView({behavior:"smooth",block:"center"});el.classList.add("tour-target-active");setTimeout(()=>el.classList.remove("tour-target-active"),650)}},[tourStep]);
  const max=parseInt(header.maxMarks)||0; const current=totalMarks(sections);
  function updateMaxMarks(v){if(String(v)===""){setHeader(h=>({...h,maxMarks:""}));return;}const n=Math.max(1,parseInt(v)||1);if(n<current){notify(`Max marks cannot be set below the current ${current} marks. Remove marks first.`);return;}setHeader(h=>({...h,maxMarks:String(n)}));}
  function addSection(){if(current>=max){notify(`You have used all ${max} marks. Increase Max Marks or reduce another section first.`);return;}setSections(s=>[...s,makeSection()]);}
  function updateSection(i,sec){setSections(s=>s.map((x,idx)=>idx===i?sec:x));}
  function deleteSection(i){if(sections.length===1){notify("Keep at least one section.");return;}setSections(s=>s.filter((_,idx)=>idx!==i));}
  function moveSection(i,j){setSections(s=>{const a=[...s];if(j<0||j>=a.length)return a;[a[i],a[j]]=[a[j],a[i]];return a})}
  function dragSection(id,targetId){if(id===targetId)return;setSections(s=>{const a=[...s],from=a.findIndex(x=>x.id===id),to=a.findIndex(x=>x.id===targetId);if(from<0||to<0)return a;const [item]=a.splice(from,1);a.splice(to,0,item);return a})}
  function duplicateSection(i){const src=sections[i];const needed=src.questionCount*src.marksPerQ;if(current+needed>max){notify(`Duplicating this section needs ${needed} more marks, but only ${Math.max(0,max-current)} are available.`);return;}const copy=JSON.parse(JSON.stringify(src));copy.id=uid();copy.heading=src.heading?`${src.heading} Copy`:"";copy.questions=copy.questions.map(q=>({...q,id:uid()}));setSections(s=>[...s.slice(0,i+1),copy,...s.slice(i+1)]);}
  function saveDraft(){const name=draftName.trim()||`Draft ${new Date().toLocaleDateString("en-IN")}`,draft={name,ts:Date.now(),header,sections};const drafts=[draft,...savedDrafts.filter(d=>d.name!==name)].slice(0,10);setSavedDrafts(drafts);setDraftName("");notify(`Saved: "${name}"`)}
  function loadDraft(d){setHeader(d.header);setSections(d.sections);setShowDraftPanel(false);setView("build")}
  function deleteDraft(ts){const d=savedDrafts.filter(x=>x.ts!==ts);setSavedDrafts(d)}

  function saveTemplate(){
    const name=(draftName.trim()||header.exam||"My Template").trim();
    const template={id:uid(),name,description:"Saved paper structure",ts:Date.now(),header:clonePaper(header),sections:clonePaper(sections)};
    setTemplates(t=>[template,...t.filter(x=>x.name!==name)].slice(0,20)); setDraftName(""); setShowTemplates(true); notify("Template saved: "+name);
  }
  function applyTemplate(template){
    if(template.sections){
      setHeader({...clonePaper(template.header||header),school:header.school,location:header.location,session:header.session,classVal:header.classVal,subject:header.subject});
      setSections(clonePaper(template.sections).map(sec=>({...sec,id:uid(),questions:(sec.questions||[]).map(q=>({...q,id:uid()}))})));
    }else{
      const paper=templateToPaper(template,{...header,school:header.school,location:header.location,session:header.session,subject:header.subject,classVal:header.classVal});
      setHeader(paper.header); setSections(paper.sections);
    }
    setShowTemplates(false); setView("build"); setPreviewPage(0); notify("Template applied.");
  }
  function deleteTemplate(id){setTemplates(t=>t.filter(x=>x.id!==id))}
  function duplicateDraft(d){
    const copy={...clonePaper(d),id:uid(),name:(d.name||"Paper")+" Copy",ts:Date.now()};
    setSavedDrafts(x=>[copy,...x].slice(0,20)); notify("Paper duplicated.");
  }
  function recoverLocal(){if(!recovery)return;setHeader({...DEFAULT_HEADER,...recovery.header});setSections(recovery.sections?.length?recovery.sections:[makeSection()]);setRecovery(null);notify("Recovered your latest local draft.");}
  function discardRecovery(){setRecovery(null)}
  function startWizard(){setWizardStep(0);setWizard({...wizard,classVal:header.classVal,subject:header.subject});setShowWizard(true)}
  function finishWizard(){
    const presets={
      quick:[["mcq4",5,1],["short2",5,2],["short3",3,3]],
      balanced:[["mcq4",5,1],["fill",5,1],["short2",5,2],["short3",3,3]],
      detailed:[["mcq4",10,1],["short2",5,2],["short3",5,3],["long",3,5]]
    };
    const pattern=presets[wizard.pattern]||presets.balanced;
    let remaining=Math.max(1,parseInt(wizard.maxMarks)||30), selected=[], used=0;
    for(const [type,count,marks] of pattern){const allowed=Math.max(0,remaining-used);const take=Math.min(count,Math.floor(allowed/marks));if(take>0){selected.push({id:uid(),typeId:type,heading:"",questionCount:take,marksPerQ:marks,questions:Array.from({length:take},()=>makeQuestion(type))});used+=take*marks;}}
    if(!selected.length) selected=[makeSection()];
    const nextHeader={...header,classVal:wizard.classVal||header.classVal,subject:wizard.subject||header.subject,exam:wizard.exam||header.exam,maxMarks:String(wizard.maxMarks||30),duration:wizard.duration||header.duration,showStudentInfo:isJuniorClass(wizard.classVal)};
    setHeader(nextHeader);setSections(selected);setShowWizard(false);setView("build");setPreviewPage(0);notify("Paper Wizard created your paper structure.");
  }

  function handleLogo(e){const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>setHeader(h=>({...h,logoUrl:ev.target.result}));r.readAsDataURL(f)}
  function newPaper(){if(!confirm("Start a new paper? Unsaved changes on this screen will be lost."))return;setHeader(DEFAULT_HEADER);setSections([makeSection()]);setView("build");setShowDraftPanel(false);window.scrollTo({top:0,behavior:"smooth"})}
  const previewHTML=useMemo(()=>buildPreviewHTML(header,sections,PRINT_FONTS[printFont],zoom),[header,sections,printFont,zoom]);
  const previewPages=useMemo(()=>paginatePaperSections(sections,8).map(page=>buildPreviewHTML(header,page,PRINT_FONTS[printFont],1)),[header,sections,printFont]);
  useEffect(()=>{setPreviewPage(p=>Math.min(p,Math.max(0,previewPages.length-1)))},[previewPages.length]);
  function filename(ext){const clean=v=>String(v||"Unknown").trim().replace(/[^a-z0-9]+/gi,"-").replace(/^-+|-+$/g,"");return `GPS-${clean(header.classVal||"Class")}-${clean(header.subject||"Subject")}.${ext}`}
  async function pdfBlob(){const holder=document.createElement("div");holder.style.position="fixed";holder.style.left="-100000px";holder.style.top="0";holder.style.width="680px";holder.innerHTML=buildPreviewHTML(header,sections,PRINT_FONTS[printFont],1);document.body.appendChild(holder);try{return await html2pdf().set({margin:0,image:{type:"jpeg",quality:.98},html2canvas:{scale:2,useCORS:true,backgroundColor:"#fff"},jsPDF:{unit:"mm",format:"a4",orientation:"portrait"},pagebreak:{mode:["css","legacy"]}}).from(holder.firstElementChild).outputPdf("blob")}finally{holder.remove()}}
  async function downloadPDF(){const blob=await pdfBlob();const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=filename("pdf");a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
  async function downloadDOCX(){try{const blob=await createDocxBlob(header,sections);const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=filename("docx");a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}catch(e){notify(`DOCX export failed. ${e.message||"Please try again."}`)}}
  function resetZoom(){setZoom(1)}

  if(view==="preview") return <div className="full-preview-screen">
    <header className="full-preview-header">
      <div className="full-preview-heading"><div className="full-preview-icon">GPS</div><div><h1>Full Paper Preview</h1><p>Page {previewPage+1} of {previewPages.length} · Review the complete A4 paper before exporting</p></div></div>
      <div className="full-preview-actions">
        <Btn onClick={()=>setView("build")} variant="secondary">← Back to Builder</Btn>
        <Btn onClick={()=>setPreviewPage(p=>Math.max(0,p-1))} variant="secondary" disabled={previewPage===0}>← Page</Btn>
        <Btn onClick={()=>setPreviewPage(p=>Math.min(previewPages.length-1,p+1))} variant="secondary" disabled={previewPage===previewPages.length-1}>Page →</Btn>
        <Btn onClick={downloadPDF}>Export PDF</Btn><Btn onClick={downloadDOCX} variant="secondary">Export DOCX</Btn><Btn onClick={()=>window.print()} variant="secondary">Print</Btn>
      </div>
    </header>
    <div className="full-preview-body">
      <aside className="preview-page-strip">
        <div className="preview-strip-title">PAGES</div>
        {previewPages.map((html,i)=><button key={i} className={i===previewPage?"preview-thumb active":"preview-thumb"} onClick={()=>setPreviewPage(i)}><span>PAGE {i+1}</span><div dangerouslySetInnerHTML={{__html:html}}/></button>)}
      </aside>
      <main className="full-preview-canvas">
        <div className="full-preview-page-label">PAGE {previewPage+1} / {previewPages.length}</div>
        <div className="full-preview-paper"><div dangerouslySetInnerHTML={{__html:previewPages[previewPage]||previewHTML}}/></div>
      </main>
    </div>
  </div>

  return <div className={`app-shell command-shell theme-${theme}`}>
    {notice&&<div className="toast">{notice}</div>}
    <header className="topbar command-topbar">
      <div className="brand command-brand"><div className="command-brand-mark">GPS</div><div><strong>GPS - QP Builder</strong><span>Goodwill Public School</span></div></div>
      <nav className="command-nav" data-tour="tour-create"><button className={view==="build"&&!showDraftPanel&&!showTemplates?"active":""} onClick={()=>{setView("build");setShowDraftPanel(false);setShowTemplates(false)}}>▣ <span>Create Paper</span></button><button className={showDraftPanel?"active":""} onClick={()=>{setShowDraftPanel(true);setShowTemplates(false)}}>▤ <span>My Papers</span></button><button className={showTemplates?"active":""} onClick={()=>{setShowTemplates(true);setShowDraftPanel(false)}}>▥ <span>Templates</span></button><button onClick={startWizard}>✦ <span>Paper Wizard</span></button><button onClick={onAccount}>⚙ <span>Settings</span></button></nav>
      <div className="command-user" onClick={onAccount}><span className="command-avatar">{String(user?.name||user?.username||"A").trim().charAt(0).toUpperCase()}</span><span className="command-user-copy"><b>{user?.name||user?.username||"Teacher"}</b><small>{user?.role==="admin"?"Administrator":"Teacher"}</small></span><span>⌄</span></div>
    </header>
    <main className="command-workspace">
      <aside className="command-sidebar"><div className="command-steps"><button className="command-step active" onClick={()=>document.getElementById("paper-details")?.scrollIntoView({behavior:"smooth"})}><span>1</span><div><b>Paper Details</b><small>Class, Subject, Marks</small></div></button><button className="command-step" onClick={()=>document.getElementById("add-questions")?.scrollIntoView({behavior:"smooth"})}><span>2</span><div><b>Add Questions</b><small>Create your questions</small></div></button><button className="command-step" onClick={()=>document.getElementById("live-preview")?.scrollIntoView({behavior:"smooth"})}><span>3</span><div><b>Preview & Export</b><small>Check and download</small></div></button></div><button className="command-help" data-tour="tour-help" onClick={startTour}><b>Need Help?</b><small>Quick Builder Tips</small><strong>›</strong></button></aside>
      <section className="command-main"><section id="paper-details" data-tour="tour-paper-details" className="command-card"><div className="command-card-head"><div className="command-icon">▧</div><div><h2>Paper Details</h2><p>Set up the basic information for your question paper</p></div></div>
        <div className="command-details-grid"><Input label="Class *" value={header.classVal} onChange={v=>setHeader(h=>({...h,classVal:v}))}/><Input label="Subject *" value={header.subject} onChange={v=>setHeader(h=>({...h,subject:v}))}/><Input label="Total Marks *" type="number" value={header.maxMarks} onChange={updateMaxMarks}/><div className="command-toggle-card" data-tour="tour-student-info"><b>Show on Paper (Class 1 & 2)</b>{isJuniorClass(header.classVal)?<label><span>Student Name & Roll No.</span><input type="checkbox" checked={header.showStudentInfo!==false} onChange={e=>setHeader(h=>({...h,showStudentInfo:e.target.checked}))}/></label>:<small>Student fields appear for Class 1 & 2.</small>}</div><Input label="Number of Questions *" value={sections.reduce((n,x)=>n+(Number(x.questionCount)||0),0)} onChange={()=>{}}/><Input label="Paper Title" value={header.exam} onChange={v=>setHeader(h=>({...h,exam:v}))}/><Input label="Date (Optional)" type="date" value={header.date} onChange={v=>setHeader(h=>({...h,date:v}))}/><div className="command-mini-note"><b>Goodwill Public School</b><small>{header.location} · Session {header.session}</small></div></div>
        <div className="command-details-extra"><Input label="School Name" value={header.school} onChange={v=>setHeader(h=>({...h,school:v}))}/><Input label="Duration" value={header.duration} onChange={v=>setHeader(h=>({...h,duration:v}))}/><div><label className="field-label">Paper Language</label><select value={header.language||"English"} onChange={e=>setHeader(h=>({...h,language:e.target.value}))} className="field-input"><option>English</option><option>Hindi</option><option>Sanskrit</option></select></div><div><label className="field-label">Question Numbering</label><select value={header.numbering||"alpha"} onChange={e=>setHeader(h=>({...h,numbering:e.target.value}))} className="field-input"><option value="alpha">A, B, C</option><option value="numeric">1, 2, 3</option><option value="roman">(i), (ii), (iii)</option></select></div></div>
        {isJuniorClass(header.classVal)&&<div className="junior-mode-banner"><div className="junior-mode-badge">JUNIOR</div><div><b>Junior Class Mode</b><span>Handwriting-friendly paper settings are active for Class 1–2. You can add answer lines and student details per question.</span></div><button onClick={()=>notify("Junior Class Mode is automatic for Class 1 and 2.")}>ⓘ</button></div>}
        <div data-tour="tour-instructions"><Textarea label="Instructions (optional)" value={header.instructions} onChange={v=>setHeader(h=>({...h,instructions:v}))} placeholder="e.g. 1. Attempt all questions."/></div></section>
        <div data-tour="tour-marks"><MarksSummary sections={sections} maxMarks={header.maxMarks}/><div className="autosave-status">Autosaved locally · {new Date(lastSavedAt).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"})}</div></div>
        <section id="add-questions" data-tour="tour-add-questions" className="command-card"><div className="command-card-head command-questions-head"><div className="command-icon">☷</div><div><h2>Add Questions</h2><p>Create questions and configure their settings</p></div><Btn onClick={addSection}>＋ Add Question</Btn></div><div>{sections.map((sec,i)=><div key={sec.id} className="section-card-wrap"><SectionEditor section={sec} index={i} total={sections.length} onChange={x=>updateSection(i,x)} onDelete={()=>deleteSection(i)} onMoveUp={()=>moveSection(i,i-1)} onMoveDown={()=>moveSection(i,i+1)} onDuplicate={()=>duplicateSection(i)} onSaveBank={(q,s)=>addToQuestionBank(q,s)} maxAllowed={max-(current-sec.questionCount*sec.marksPerQ)} notify={notify} language={header.language||"English"} numbering={header.numbering||"alpha"} junior={isJuniorClass(header.classVal)}/></div>)}</div><button onClick={addSection} className="add-section">＋ Add Section (Optional)</button></section>
      </section>
      <aside className="command-preview" id="live-preview" data-tour="tour-preview"><div className="command-preview-head"><div className="command-preview-title"><div className="command-icon">◉</div><div><h2>Live Preview</h2><p>See how your question paper will look</p></div></div><div className="command-preview-tabs"><button className="active">Preview</button><button data-tour="tour-full-preview" onClick={()=>setView("preview")}>Full Preview</button></div></div><div className="command-preview-scroll"><div ref={previewRef} className="preview-zoom-layer" style={{transform:`scale(${zoomLevel/100})`,transformOrigin:"top center"}} dangerouslySetInnerHTML={{__html:previewHTML}}/></div><div className="command-preview-footer" data-tour="tour-exports"><Btn onClick={()=>setView("preview")}>↗ Full Preview</Btn><Btn onClick={downloadPDF}>▣ Export as PDF</Btn><Btn onClick={downloadDOCX} variant="secondary">▣ Export as DOCX</Btn></div></aside>

    </main>


    {commandOpen&&<div className="feature-modal-backdrop command-palette-backdrop" onClick={e=>{if(e.target===e.currentTarget)setCommandOpen(false)}}><div className="command-palette"><div className="palette-search">⌘ <input autoFocus placeholder="Search commands..." onKeyDown={e=>{if(e.key==="Escape")setCommandOpen(false)}}/></div><div className="palette-grid">
      {[["Create New Paper",()=>{setView("build");setCommandOpen(false)}],["My Papers",()=>{setShowDraftPanel(true);setCommandOpen(false)}],["Templates",()=>{setShowTemplates(true);setCommandOpen(false)}],["Full Preview",()=>{setView("preview");setCommandOpen(false)}],["Save Paper",()=>{saveDraft();setCommandOpen(false)}],["Export PDF",()=>{downloadPDF();setCommandOpen(false)}],["Export DOCX",()=>{downloadDOCX();setCommandOpen(false)}],["Question Bank",()=>{setBankOpen(true);setCommandOpen(false)}],["Undo",()=>{undo();setCommandOpen(false)}],["Redo",()=>{redo();setCommandOpen(false)}],["Design",()=>{setDesignOpen(true);setCommandOpen(false)}],["Focus Mode",()=>{setFocusMode(true);setCommandOpen(false)}]].map(([name,fn])=><button key={name} onClick={fn}><span>⌘</span>{name}<kbd>↵</kbd></button>)}</div><small>Esc to close · Ctrl/Cmd + K to open</small></div></div>}
    {designOpen&&<div className="feature-modal-backdrop" onClick={e=>{if(e.target===e.currentTarget)setDesignOpen(false)}}><div className="feature-modal design-modal"><div className="feature-modal-head"><div><h2>Paper & Workspace Design</h2><p>Personalize the builder without changing the printed paper.</p></div><button className="modal-close" onClick={()=>setDesignOpen(false)}>×</button></div><div className="design-group"><b>Workspace Theme</b><div className="theme-grid">{[["gps","GPS Burgundy"],["midnight","Midnight"],["slate","Slate"],["emerald","Emerald"]].map(([id,n])=><button className={theme===id?"theme-chip selected":"theme-chip"} onClick={()=>setTheme(id)} key={id}><span className={"theme-dot "+id}/>{n}</button>)}</div></div><div className="design-group"><b>Preview Zoom</b><div className="zoom-row"><button onClick={()=>setZoomLevel(z=>Math.max(70,z-10))}>−</button><strong>{zoomLevel}%</strong><button onClick={()=>setZoomLevel(z=>Math.min(150,z+10))}>+</button><button onClick={()=>setZoomLevel(100)}>Fit</button></div></div><div className="design-group"><b>Paper Typography</b><div className="theme-grid"><button className={printFont==="small"?"theme-chip selected":"theme-chip"} onClick={()=>setPrintFont("small")}>Compact</button><button className={printFont==="medium"?"theme-chip selected":"theme-chip"} onClick={()=>setPrintFont("medium")}>Standard</button><button className={printFont==="large"?"theme-chip selected":"theme-chip"} onClick={()=>setPrintFont("large")}>Spacious</button></div></div></div></div>}
    {bankOpen&&<div className="feature-modal-backdrop" onClick={e=>{if(e.target===e.currentTarget)setBankOpen(false)}}><div className="feature-modal feature-wide"><div className="feature-modal-head"><div><h2>Question Bank</h2><p>Store reusable questions and insert them into compatible sections.</p></div><button className="modal-close" onClick={()=>setBankOpen(false)}>×</button></div><input className="field-input" value={bankSearch} onChange={e=>setBankSearch(e.target.value)} placeholder="Search saved questions..."/><div className="bank-list">{questionBank.filter(x=>(x.question?.text||"").toLowerCase().includes(bankSearch.toLowerCase())||(x.subject||"").toLowerCase().includes(bankSearch.toLowerCase())).map(item=><div className="bank-item" key={item.id}><div><b dangerouslySetInnerHTML={{__html:item.question?.text||"Untitled question"}}/><small>{item.subject||"Subject"} · {item.classVal||"Class"} · {item.marks||0} mark(s)</small></div><Btn onClick={()=>insertBankQuestion(item)} small>Insert</Btn></div>)}{!questionBank.length&&<div className="empty-state"><div>▦</div><b>Your Question Bank is empty</b><span>Use the Bank button on a question to save reusable questions.</span></div>}</div></div></div>}
    {focusMode&&<div className="focus-overlay"><header><b>GPS QP Builder · Focus Mode</b><div><Btn onClick={undo} variant="secondary" small disabled={!history.length}>↶</Btn><Btn onClick={redo} variant="secondary" small disabled={!future.length}>↷</Btn><Btn onClick={()=>setFocusMode(false)} small>Exit Focus</Btn></div></header><div className="focus-content"><div className="focus-paper-meta"><span>{header.classVal||"Class"} · {header.subject||"Subject"}</span><b>{sections.reduce((n,s)=>n+(Number(s.questionCount)||0),0)} Questions</b></div>{sections.flatMap(s=>s.questions.map((q,i)=>({q,s,i}))).map(({q,s,i})=><div className="focus-question" key={q.id}><span>Q{ i+1 } · {s.marksPerQ} mark(s)</span><RichTextEditor value={q.text} onChange={text=>{const si=sections.findIndex(x=>x.id===s.id);if(si>=0){const next=[...sections];next[si]={...next[si],questions:next[si].questions.map(x=>x.id===q.id?{...x,text}:x)};setSections(next)}}} placeholder="Write your question here"/></div>))}</div></div>}

    {recovery&&<div className="feature-modal-backdrop"><div className="feature-modal recovery-modal"><div className="feature-modal-icon">↻</div><h2>We recovered your latest work</h2><p>A local autosave is newer than the copy stored on the server. Would you like to restore it?</p><div className="feature-meta">Autosaved {new Date(recovery.ts).toLocaleString("en-IN")}</div><div className="feature-actions"><Btn onClick={discardRecovery} variant="secondary">Discard</Btn><Btn onClick={recoverLocal}>Recover Draft</Btn></div></div></div>}
    {showDraftPanel&&<div className="feature-modal-backdrop" onClick={e=>{if(e.target===e.currentTarget)setShowDraftPanel(false)}}><div className="feature-modal feature-wide"><div className="feature-modal-head"><div><h2>My Papers</h2><p>Save, reopen, duplicate and manage your question papers.</p></div><button className="modal-close" onClick={()=>setShowDraftPanel(false)}>×</button></div><div className="feature-save-row"><input className="field-input" value={draftName} onChange={e=>setDraftName(e.target.value)} placeholder="Paper name, e.g. Class VI Half Yearly"/><Btn onClick={saveDraft}>Save Current Paper</Btn><Btn onClick={saveTemplate} variant="secondary">Save as Template</Btn></div>{savedDrafts.length?<div className="paper-list">{savedDrafts.map(d=><div className="paper-list-item" key={d.ts}><div><b>{d.name}</b><small>{d.header?.classVal||"Class"} · {d.header?.subject||"Subject"} · {d.header?.maxMarks||0} marks · {new Date(d.ts).toLocaleDateString("en-IN")}</small></div><div className="paper-list-actions"><Btn onClick={()=>loadDraft(d)} small>Open</Btn><Btn onClick={()=>duplicateDraft(d)} variant="secondary" small>Duplicate</Btn><Btn onClick={()=>deleteDraft(d.ts)} variant="danger" small>Delete</Btn></div></div>)}</div>:<div className="empty-state"><div>▤</div><b>No saved papers yet</b><span>Save your current paper and it will appear here.</span></div>}</div></div>}
    {showTemplates&&<div className="feature-modal-backdrop" onClick={e=>{if(e.target===e.currentTarget)setShowTemplates(false)}}><div className="feature-modal feature-wide"><div className="feature-modal-head"><div><h2>Paper Templates</h2><p>Start from a proven structure, then customize everything.</p></div><button className="modal-close" onClick={()=>setShowTemplates(false)}>×</button></div><div className="template-grid">{PAPER_TEMPLATES.concat(templates).map(t=><div className="template-card" key={t.id}><div className="template-card-icon">▤</div><h3>{t.name}</h3><p>{t.description||"Saved paper structure"}</p><small>{t.pattern?t.pattern.reduce((n,x)=>n+x[1],0):t.sections?.reduce((n,x)=>n+(Number(x.questionCount)||0),0)} questions</small><div><Btn onClick={()=>applyTemplate(t)} small>Use Template</Btn>{!PAPER_TEMPLATES.some(x=>x.id===t.id)&&<Btn onClick={()=>deleteTemplate(t.id)} variant="danger" small>Delete</Btn>}</div></div>)}</div></div></div>}
    {showWizard&&<div className="feature-modal-backdrop"><div className="feature-modal wizard-modal"><div className="wizard-progress"><span className={wizardStep>=0?"on":""}>1</span><i/><span className={wizardStep>=1?"on":""}>2</span><i/><span className={wizardStep>=2?"on":""}>3</span><i/><span className={wizardStep>=3?"on":""}>4</span></div><div className="feature-modal-head"><div><h2>Paper Wizard</h2><p>Build a complete paper structure in under a minute.</p></div><button className="modal-close" onClick={()=>setShowWizard(false)}>×</button></div>
      {wizardStep===0&&<div className="wizard-grid"><Input label="Class" value={wizard.classVal} onChange={v=>setWizard(w=>({...w,classVal:v}))}/><Input label="Subject" value={wizard.subject} onChange={v=>setWizard(w=>({...w,subject:v}))}/><Input label="Exam / Paper Title" value={wizard.exam} onChange={v=>setWizard(w=>({...w,exam:v}))}/><Input label="Total Marks" type="number" value={wizard.maxMarks} onChange={v=>setWizard(w=>({...w,maxMarks:v}))}/><Input label="Duration" value={wizard.duration} onChange={v=>setWizard(w=>({...w,duration:v}))}/></div>}
      {wizardStep===1&&<div><div className="wizard-choice-grid">{[{id:"quick",name:"Quick Test",desc:"20–25 marks, compact"}, {id:"balanced",name:"Balanced",desc:"Mixed objective + written"}, {id:"detailed",name:"Detailed Exam",desc:"More sections and long answers"}].map(x=><button key={x.id} className={"wizard-choice "+(wizard.pattern===x.id?"selected":"")} onClick={()=>setWizard(w=>({...w,pattern:x.id}))}><b>{x.name}</b><span>{x.desc}</span></button>)}</div><div className="wizard-tip">The wizard creates editable question sections. You can change every question, mark and section after creation.</div></div>}
      {wizardStep===2&&<div className="junior-mode-panel"><div className="junior-mode-icon">✎</div><div><h3>{isJuniorClass(wizard.classVal)?"Junior Class Mode is ON":"Standard Class Mode"}</h3><p>{isJuniorClass(wizard.classVal)?"Class 1–2 papers will include handwriting-friendly answer lines and optional Name / Roll Number fields.":"Choose Class 1 or 2 in the previous step to activate Junior Class Mode automatically."}</p></div></div>}
      {wizardStep===3&&<div className="wizard-review"><div><b>{wizard.classVal}</b><span>Class</span></div><div><b>{wizard.subject}</b><span>Subject</span></div><div><b>{wizard.maxMarks}</b><span>Marks</span></div><div><b>{wizard.duration}</b><span>Duration</span></div><p>Ready to create your editable paper structure.</p></div>}
      <div className="wizard-footer"><Btn onClick={()=>wizardStep===0?setShowWizard(false):setWizardStep(s=>s-1)} variant="secondary">{wizardStep===0?"Cancel":"← Back"}</Btn>{wizardStep<3?<Btn onClick={()=>setWizardStep(s=>s+1)}>Next →</Btn>:<Btn onClick={finishWizard}>Create Paper ✓</Btn>}</div>
    </div></div>}

    {tourStep!==null&&<div className="tour-layer" onClick={e=>{if(e.target===e.currentTarget)finishTour()}}>
      <div className="tour-tooltip">
        <button className="tour-exit" onClick={finishTour} aria-label="Exit guide">×</button>
        <div className="tour-step-count">STEP {tourStep+1} OF {tourSteps.length}</div>
        <h3>{tourSteps[tourStep].title}</h3>
        <p>{tourSteps[tourStep].text}</p>
        <div className="tour-footer"><button className="tour-exit-text" onClick={finishTour}>Exit Guide</button>{tourStep<tourSteps.length-1?<Btn onClick={()=>setTourStep(tourStep+1)}>Next →</Btn>:<Btn onClick={finishTour}>Finish ✓</Btn>}</div>
      </div>
    </div>}
  </div>;

}

function SetupScreen({onDone}){
  const [key,setKey]=useState(""); const [busy,setBusy]=useState(false); const [error,setError]=useState("");
  async function submit(e){e.preventDefault();setError("");if(!key)return setError("Enter the setup key configured in Cloudflare.");setBusy(true);try{const r=await fetch("/api/setup/bootstrap",{method:"POST",headers:{"Content-Type":"application/json","x-setup-key":key},body:JSON.stringify({admin:{name:"masterdev",username:"masterdev",email:"",password:"masterdev"}})});const d=await r.json().catch(()=>({}));if(!r.ok)throw new Error(d.error||"Setup failed.");onDone()}catch(e){setError(e.message)}finally{setBusy(false)}}
  return <div className="auth-screen"><form className="auth-card" onSubmit={submit}><img src={gpsLogo} className="auth-logo" alt="Goodwill Public School"/><h1>School Setup</h1><p className="small-muted">Create the school administrator account.</p><Input label="Setup Key" type="password" value={key} onChange={setKey}/>{error&&<div className="auth-error">{error}</div>}<Btn full disabled={busy}>{busy?"Setting up…":"Create Admin Account"}</Btn></form></div>;
}
function LoginScreen({onLogin}){
  const [username,setUsername]=useState(""); const [password,setPassword]=useState(""); const [busy,setBusy]=useState(false); const [error,setError]=useState("");
  async function submit(e){e.preventDefault();setError("");setBusy(true);try{const r=await fetch("/api/auth/login",{method:"POST",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({username,password})});const d=await r.json().catch(()=>({}));if(!r.ok)throw new Error(d.error||"Login failed.");onLogin(d.user,d.workspace)}catch(e){setError(e.message)}finally{setBusy(false)}}
  return <div className="auth-screen"><form className="auth-card" onSubmit={submit}><img src={gpsLogo} className="auth-logo" alt="Goodwill Public School"/><h1>GPS QP Builder</h1><p className="small-muted">School Question Paper Builder</p><Input label="Username" value={username} onChange={setUsername}/><Input label="Password" type="password" value={password} onChange={setPassword}/>{error&&<div className="auth-error">{error}</div>}<Btn full disabled={busy}>{busy?"Signing in…":"Sign in"}</Btn></form></div>;
}
function AccountPanel({user,onClose,onUserUpdated,notify}){
  const [name,setName]=useState(user.name||""); const [username,setUsername]=useState(user.username||""); const [email,setEmail]=useState(user.email||""); const [currentPassword,setCurrentPassword]=useState(""); const [newPassword,setNewPassword]=useState(""); const [confirmPassword,setConfirmPassword]=useState(""); const [busy,setBusy]=useState(false);
  async function save(e){e.preventDefault();setBusy(true);try{const r=await fetch("/api/auth/account",{method:"PUT",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({name,username,email})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Could not update account.");onUserUpdated(d.user);notify("Account details updated.")}catch(e){notify(e.message)}finally{setBusy(false)}}
  async function changePassword(e){e.preventDefault();if(newPassword!==confirmPassword){notify("New passwords do not match.");return}setBusy(true);try{const r=await fetch("/api/auth/password",{method:"PUT",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({currentPassword,newPassword})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Could not change password.");setCurrentPassword("");setNewPassword("");setConfirmPassword("");notify("Password changed successfully.")}catch(e){notify(e.message)}finally{setBusy(false)}}
  return <div className="modal-backdrop"><div className="modal-card"><div className="modal-head"><h2>My Account</h2><button className="icon-btn" onClick={onClose}>✕</button></div><form onSubmit={save}><Input label="Name" value={name} onChange={setName}/><Input label="Username" value={username} onChange={setUsername}/><Input label="Email" value={email} onChange={setEmail}/><Btn disabled={busy}>Save Account Details</Btn></form><hr/><h3>Change Password</h3><form onSubmit={changePassword}><Input label="Current Password" type="password" value={currentPassword} onChange={setCurrentPassword}/><Input label="New Password" type="password" value={newPassword} onChange={setNewPassword}/><Input label="Confirm New Password" type="password" value={confirmPassword} onChange={setConfirmPassword}/><Btn variant="secondary" disabled={busy}>Change Password</Btn></form></div></div>;
}
function AdminPanel({onClose,notify}){
  const [users,setUsers]=useState([]); const [loading,setLoading]=useState(true); const [form,setForm]=useState({name:"",username:"",email:"",password:""}); const [bulk,setBulk]=useState(""); const [busy,setBusy]=useState(false);
  async function load(){setLoading(true);try{const r=await fetch("/api/admin/users",{credentials:"include"});const d=await r.json();if(!r.ok)throw new Error(d.error||"Could not load users.");setUsers(d.users||[])}catch(e){notify(e.message)}finally{setLoading(false)}} useEffect(()=>{load()},[]);
  async function create(e){e.preventDefault();setBusy(true);try{const r=await fetch("/api/admin/users",{method:"POST",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({...form,role:"teacher"})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Could not create user.");setForm({name:"",username:"",email:"",password:""});notify("Teacher account created.");load()}catch(e){notify(e.message)}finally{setBusy(false)}}
  async function toggle(u){try{const r=await fetch(`/api/admin/users/${u.id}`,{method:"PATCH",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({active:!u.active})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Update failed.");load()}catch(e){notify(e.message)}}
  async function reset(u){const pw=prompt(`Temporary password for ${u.username}:`,"GPS@ChangeMe1");if(!pw)return;try{const r=await fetch(`/api/admin/users/${u.id}/reset-password`,{method:"POST",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({password:pw})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Reset failed.");notify("Password reset.")}catch(e){notify(e.message)}}
  return <div className="modal-backdrop"><div className="modal-card admin-modal"><div className="modal-head"><h2>Admin Panel</h2><button className="icon-btn" onClick={onClose}>✕</button></div><h3>Accounts</h3>{loading?<div className="small-muted">Loading…</div>:users.map(u=><div className="admin-user" key={u.id}><div><b>{u.name||u.username}</b><div className="small-muted">@{u.username} · {u.role} · {u.active?"Active":"Disabled"}</div></div><div className="toolbar-row"><Btn small variant="secondary" onClick={()=>reset(u)}>Reset password</Btn><Btn small variant={u.active?"danger":"success"} onClick={()=>toggle(u)}>{u.active?"Disable":"Enable"}</Btn></div></div>)}<hr/><h3>Create teacher account</h3><form onSubmit={create}><div className="two-col"><Input label="Name" value={form.name} onChange={v=>setForm(f=>({...f,name:v}))}/><Input label="Username" value={form.username} onChange={v=>setForm(f=>({...f,username:v}))}/><Input label="Email" value={form.email} onChange={v=>setForm(f=>({...f,email:v}))}/><Input label="Temporary Password" type="password" value={form.password} onChange={v=>setForm(f=>({...f,password:v}))}/></div><Btn disabled={busy}>Create Teacher Account</Btn></form></div></div>;
}
export default function App(){
 const [auth,setAuth]=useState(null); const [loading,setLoading]=useState(true); const [setup,setSetup]=useState(false); const [panel,setPanel]=useState(null);
 useEffect(()=>{(async()=>{try{const r=await fetch("/api/auth/me",{credentials:"include"});if(r.ok){const d=await r.json();setAuth(d);return;}const s=await fetch("/api/setup/status",{cache:"no-store"});const sd=await s.json();setSetup(!!sd.needsSetup)}catch(e){}finally{setLoading(false)}})()},[]);
 async function afterSetup(){const r=await fetch("/api/setup/status",{cache:"no-store"});const d=await r.json();if(!d.needsSetup)setSetup(false)}
 async function logout(){await fetch("/api/auth/logout",{method:"POST",credentials:"include"});setAuth(null)}
 if(loading)return <div className="auth-screen"><div className="auth-card"><h2>Loading…</h2></div></div>;
 if(setup)return <SetupScreen onDone={afterSetup}/>;
 if(!auth)return <LoginScreen onLogin={(user,workspace)=>setAuth({user,workspace})}/>;
 return <><BuilderApp user={auth.user} initialWorkspace={auth.workspace} onWorkspaceChange={async data=>{await fetch("/api/workspace",{method:"PUT",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify(data)})}} onLogout={logout} onAccount={()=>setPanel("account")} onAdmin={()=>setPanel("admin")}/>{panel==="account"&&<AccountPanel user={auth.user} onClose={()=>setPanel(null)} onUserUpdated={u=>setAuth(a=>({...a,user:u}))} notify={m=>alert(m)}/>} {panel==="admin"&&auth.user.role==="admin"&&<AdminPanel onClose={()=>setPanel(null)} notify={m=>alert(m)}/>}</>;
}
