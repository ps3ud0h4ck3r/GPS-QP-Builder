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

const DEFAULT_HEADER = {
  school: "Goodwill Public School", location: "Uttam Nagar", exam: "Half Early Examination", session: "2026–27",
  classVal: "VI", subject: "Computer", duration: "1.5 Hours", maxMarks: "30", date: "", instructions: "", logoUrl: null, language: "English", numbering: "alpha",
};

function uid() { return Math.random().toString(36).slice(2, 10); }
function makeQuestion(typeId) {
  const type = QUESTION_TYPES.find(t => t.id === typeId);
  return { id: uid(), text: "", diagram: null, options: type.hasOptions ? Array.from({length:type.hasOptions}, () => "") : null,
    matchLeft: typeId === "match" ? ["","","","",""] : null, matchRight: typeId === "match" ? ["","","","",""] : null };
}
function makeSection() {
  return { id: uid(), typeId: "fill", heading: "", questionCount: 1, marksPerQ: 1, questions: [makeQuestion("fill")] };
}
function totalMarks(sections) { return sections.reduce((s, sec) => s + sec.questionCount * sec.marksPerQ, 0); }
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

function QuestionEditor({ q,index,typeId,onChange,showDiagram,onMoveUp,onMoveDown,onDelete,onReorder,numbering }) {
  const type=QUESTION_TYPES.find(t=>t.id===typeId); const [showDiag,setShowDiag]=useState(!!q.diagram);
  const upd=p=>onChange({...q,...p});
  return <div className={`question-editor ${q.pageBreakAfter?"question-break-after":""}`} draggable onDragStart={e=>e.dataTransfer.setData("text/plain",q.id)} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const source=e.dataTransfer.getData("text/plain");if(source&&source!==q.id)onReorder(source,q.id)}}>
    <div className="question-editor-head"><span className="drag-handle" title="Drag to reorder">⠿</span><Badge color="#ede9fe" textColor="#5b21b6">{ALPHA[index]}</Badge><span className="small-muted">{type.label}</span><span style={{marginLeft:"auto"}}><Btn onClick={onMoveUp} variant="ghost" small title="Move up">↑</Btn><Btn onClick={onMoveDown} variant="ghost" small title="Move down">↓</Btn><Btn onClick={onDelete} variant="danger" small title="Delete question">✕</Btn></span></div>
    {typeId==="match" ? <div className="two-col"><div><div className="field-label">Column A</div>{q.matchLeft.map((v,i)=><input key={i} value={v} placeholder={`${i+1}.`} onChange={e=>{const a=[...q.matchLeft];a[i]=e.target.value;upd({matchLeft:a})}} className="field-input compact"/>)}</div><div><div className="field-label">Column B</div>{q.matchRight.map((v,i)=><input key={i} value={v} placeholder={`${String.fromCharCode(97+i)}.`} onChange={e=>{const a=[...q.matchRight];a[i]=e.target.value;upd({matchRight:a})}} className="field-input compact"/>)}</div></div> : <>
      <RichTextEditor value={q.text} onChange={text=>upd({text})} placeholder={typeId==="fill"?"Write statement with ___ for blank":typeId==="tf"?"Write a true/false statement":"Write your question here"}/>
      {type.hasOptions && <div style={{marginTop:8}}>{Array.from({length:type.hasOptions},(_,i)=><div className="option-row" key={i}><span>{OPT_ALPHA[i]})</span><input value={q.options[i]} onChange={e=>{const a=[...q.options];a[i]=e.target.value;upd({options:a})}} placeholder={`Option ${OPT_ALPHA[i].toUpperCase()}`} className="field-input compact"/></div>)}</div>}
    </>}
    {showDiagram && type.hasDiagram && <div style={{marginTop:8}}>{!showDiag&&!q.diagram?<button className="link-btn" onClick={()=>setShowDiag(true)}>+ Add Diagram</button>:<DiagramPicker diagram={q.diagram} onChange={d=>{upd({diagram:d});if(!d)setShowDiag(false)}}/>}</div>}
    <label className="check-row"><input type="checkbox" checked={!!q.pageBreakAfter} onChange={e=>upd({pageBreakAfter:e.target.checked})}/> Start a new A4 page after this question</label>
  </div>;
}

function SectionEditor({ section,index,total,onChange,onDelete,onMoveUp,onMoveDown,onDuplicate,maxAllowed,notify,language,numbering }) {
  const [open,setOpen]=useState(true);
  const type=QUESTION_TYPES.find(t=>t.id===section.typeId); const sectionMarks=section.questionCount*section.marksPerQ;
  const upd=p=>onChange({...section,...p});
  function changeType(id){const t=QUESTION_TYPES.find(x=>x.id===id);const maxCount=Math.max(1,Math.floor(maxAllowed/t.defaultMarks)||1);const count=Math.min(section.questionCount,maxCount);upd({typeId:id,marksPerQ:t.defaultMarks,questionCount:count,questions:Array.from({length:count},()=>makeQuestion(id))});}
  function changeCount(n){const desired=Math.max(1,Math.min(20,parseInt(n)||1));const maxCount=Math.max(1,Math.floor(maxAllowed/section.marksPerQ)||1);if(desired*section.marksPerQ>maxAllowed){notify(`This section can have at most ${maxCount} question${maxCount===1?"":"s"}. Maximum paper marks are ${maxAllowed} here.`);return;}const qs=[...section.questions];while(qs.length<desired)qs.push(makeQuestion(section.typeId));upd({questionCount:desired,questions:qs.slice(0,desired)});}
  function changeMarks(n){const m=Math.max(1,Math.min(20,parseInt(n)||1));if(section.questionCount*m>maxAllowed){notify(`These marks would exceed the remaining ${maxAllowed} marks.`);return;}upd({marksPerQ:m});}
  function updateQuestion(id,p){upd({questions:section.questions.map(q=>q.id===id?p:q)});}
  function reorder(a,b){const arr=[...section.questions],ai=arr.findIndex(q=>q.id===a),bi=arr.findIndex(q=>q.id===b);if(ai<0||bi<0)return;const [item]=arr.splice(ai,1);arr.splice(bi,0,item);upd({questions:arr});}
  return <Card className="section-card"><div className="section-head"><div className="section-title"><button className="collapse-btn" onClick={()=>setOpen(o=>!o)}>{open?"▾":"▸"}</button><Badge>{index+1}</Badge><input value={section.heading} onChange={e=>upd({heading:e.target.value})} placeholder={`Section ${index+1} heading`} className="section-heading-input"/><span className="small-muted">{sectionMarks} marks</span></div><div className="toolbar-row"><Btn variant="ghost" small onClick={onMoveUp}>↑</Btn><Btn variant="ghost" small onClick={onMoveDown}>↓</Btn><Btn variant="secondary" small onClick={onDuplicate}>Duplicate</Btn><Btn variant="danger" small onClick={onDelete}>Delete</Btn></div></div>{open&&<div className="section-body"><div className="two-col"><div><label className="field-label">Question Type</label><select value={section.typeId} onChange={e=>changeType(e.target.value)} className="field-input">{QUESTION_TYPES.map(t=><option key={t.id} value={t.id}>{t.label}</option>)}</select></div><Input label="Number of Questions" type="number" value={section.questionCount} onChange={changeCount}/><Input label="Marks per Question" type="number" value={section.marksPerQ} onChange={changeMarks}/><Input label="Section Marks" value={sectionMarks} disabled onChange={()=>{}}/></div>{section.questions.map((q,i)=><QuestionEditor key={q.id} q={q} index={i} typeId={section.typeId} numbering={numbering} showDiagram onChange={p=>updateQuestion(q.id,p)} onMoveUp={()=>{const arr=[...section.questions],x=arr.splice(i,1)[0];arr.splice(Math.max(0,i-1),0,x);upd({questions:arr})}} onMoveDown={()=>{const arr=[...section.questions],x=arr.splice(i,1)[0];arr.splice(Math.min(arr.length,i+1),0,x);upd({questions:arr})}} onDelete={()=>{if(section.questions.length<=1){notify("Keep at least one question in each section.");return}upd({questionCount:section.questionCount-1,questions:section.questions.filter(x=>x.id!==q.id)})}} onReorder={reorder}/>)}<div className="small-muted">Language: {language}</div></div>}</Card>;
}

function PagedPreview({header,sections,zoom=1,printSize="medium"}) {
  const pages=[]; let page=[]; let used=0; const cap=900;
  const push=(node,h=70)=>{if(used+h>cap&&page.length){pages.push(page);page=[];used=0;}page.push(node);used+=h;};
  push(<div className="paper-header" key="head"><img src={header.logoUrl||gpsLogo} className="paper-logo" alt=""/><div className="paper-school">{header.school}</div><div className="paper-location">{header.location}</div><div className="paper-exam">{header.exam} {header.session}</div><div className="paper-class">CLASS {header.classVal}</div><div className="paper-meta">Time: {header.duration}<span>Subject: {header.subject}</span><span>M.M.: {header.maxMarks}</span></div>{header.instructions&&<div className="paper-instructions"><b>Instructions:</b>{String(header.instructions).replace(/\r\n/g,"\n").replace(/\s+(?=\d+\.\s)/g,"\n").split(/\n+/).filter(Boolean).map((x,i)=><div key={i}>{x.trim()}</div>)}</div>}</div>,260);
  sections.forEach((sec,si)=>{const type=QUESTION_TYPES.find(t=>t.id===sec.typeId);const marks=sec.marksPerQ===1?`(${sec.questionCount*sec.marksPerQ})`:`(${sec.marksPerQ}x${sec.questionCount}=${sec.questionCount*sec.marksPerQ})`;push(<div className="paper-section" key={`s-${sec.id}`}><div className="paper-section-head"><b>Q{si+1}. {sec.heading||type.label}</b><b>{marks}</b></div>{sec.questions.map((q,qi)=><div className="paper-question" key={q.id}><div><b>{questionLabel(qi,header.numbering||"alpha")}</b> {q.text? <span dangerouslySetInnerHTML={{__html:q.text}}/>:<span>___________________________________</span>}</div>{q.options&&<div className="paper-options">{q.options.map((o,i)=><div key={i}><b>{OPT_ALPHA[i]})</b> {o||"________"}</div>)}</div>}{sec.typeId==="match"&&<table className="question-table"><thead><tr><th>Column A</th><th>Column B</th></tr></thead><tbody>{(q.matchLeft||[]).map((l,i)=><tr key={i}><td>{i+1}. {l||"_______"}</td><td>{String.fromCharCode(97+i)}. {(q.matchRight||[])[i]||"_______"}</td></tr>)}</tbody></table>}{q.diagram&&<div className="paper-diagram">{q.diagram.mode==="space"?<div className="diagram-space">Diagram space</div>:<img src={q.diagram.dataUrl} alt="diagram"/>}</div>}</div>)}</div>,120+sec.questions.length*70);if(sec.pageBreak==="after") {pages.push(page);page=[];used=0;}});
  if(page.length)pages.push(page);
  return <div className="paged-preview" style={{transform:`scale(${zoom})`,transformOrigin:"top center"}}>{pages.map((p,i)=><div className="a4-page" key={i} style={{fontSize:PRINT_FONTS[printSize]}}>{p}</div>)}</div>;
}

function questionLabel(i,mode){if(mode==="numeric")return `${i+1}.`;if(mode==="roman")return `(${["i","ii","iii","iv","v","vi","vii","viii","ix","x"][i]||i+1})`;return `${ALPHA[i]||String.fromCharCode(65+i)}.`;}

function BuilderApp({user,initialWorkspace,onWorkspaceChange,onLogout,onAccount,onAdmin}) {
  const ws=initialWorkspace||{}; const [header,setHeader]=useState({...DEFAULT_HEADER,...(ws.current?.header||{})}); const [sections,setSections]=useState(ws.current?.sections?.length?ws.current.sections:[makeSection()]); const [savedDrafts,setSavedDrafts]=useState(ws.drafts||[]); const [templates,setTemplates]=useState(ws.templates||[]); const [view,setView]=useState("build"); const [notice,setNotice]=useState(""); const [dark,setDark]=useState(!!ws.settings?.dark); const [printSize,setPrintSize]=useState(ws.settings?.printSize||"medium"); const [zoom,setZoom]=useState(ws.settings?.zoom||1);
  const notify=m=>{setNotice(m);setTimeout(()=>setNotice(""),2600)};
  useEffect(()=>{document.documentElement.dataset.theme=dark?"dark":"light";},[dark]);
  useEffect(()=>{const t=setTimeout(()=>onWorkspaceChange({current:{header,sections},drafts:savedDrafts,templates,settings:{dark,printSize,zoom}}),500);return()=>clearTimeout(t)},[header,sections,savedDrafts,templates,dark,printSize,zoom]);
  function newPaper(){setHeader({...DEFAULT_HEADER});setSections([makeSection()]);setView("build");notify("New paper started.");}
  function updateSection(i,s){setSections(a=>a.map((x,j)=>j===i?s:x));}
  function duplicate(i){setSections(a=>{const copy={...a[i],id:uid(),questions:a[i].questions.map(q=>({...q,id:uid()}))};return [...a.slice(0,i+1),copy,...a.slice(i+1)]})}
  function saveDraft(){const name=prompt("Draft name:",`${header.subject||"Paper"} ${header.classVal||""}`);if(!name)return;setSavedDrafts(a=>[...a,{id:uid(),name,header,sections,savedAt:Date.now()}]);notify("Draft saved.")}
  function loadDraft(d){setHeader(d.header);setSections(d.sections);setView("build");notify(`Loaded ${d.name}.`)}
  function saveTemplate(){const name=prompt("Template name:",`${header.subject||"Template"}`);if(!name)return;setTemplates(a=>[...a,{id:uid(),name,header,sections,savedAt:Date.now()}]);notify("Template saved.")}
  function loadTemplate(t){setHeader(t.header);setSections(t.sections);setView("build");notify(`Loaded ${t.name}.`)}
  const total=totalMarks(sections), target=parseInt(header.maxMarks)||0, remaining=Math.max(0,target-total);
  async function exportPDF(){const el=document.querySelector(".paged-preview");if(!el)return;const clone=el.cloneNode(true);clone.style.transform="none";const holder=document.createElement("div");holder.style.position="fixed";holder.style.left="-10000px";holder.style.top="0";holder.appendChild(clone);document.body.appendChild(holder);await html2pdf().set({margin:0,filename:`GPS-${header.classVal||"Class"}-${header.subject||"Subject"}.pdf`,image:{type:"jpeg",quality:0.98},html2canvas:{scale:2,useCORS:true},jsPDF:{unit:"mm",format:"a4",orientation:"portrait"}}).from(clone).save();holder.remove()}
  async function exportDOCX(){const children=[];children.push(new Paragraph({children:[new TextRun({text:header.school||"Question Paper",bold:true,size:30})],alignment:AlignmentType.CENTER,spacing:{after:50}}));if(header.location)children.push(new Paragraph({children:[new TextRun({text:header.location,size:20})],alignment:AlignmentType.CENTER,spacing:{after:50}}));children.push(new Paragraph({children:[new TextRun({text:`${header.exam||""} ${header.session||""}`,bold:true,size:22})],alignment:AlignmentType.CENTER,spacing:{after:50}}));children.push(new Paragraph({children:[new TextRun({text:`CLASS ${header.classVal||""}`,bold:true,size:20})],alignment:AlignmentType.CENTER,spacing:{after:120}}));children.push(new Paragraph({children:[new TextRun({text:`Time: ${header.duration||""}     Subject: ${header.subject||""}     M.M.: ${header.maxMarks||""}`,bold:true,size:20})],spacing:{after:180}}));if(header.instructions){children.push(new Paragraph({children:[new TextRun({text:"Instructions:",bold:true,size:20})],spacing:{after:40}}));String(header.instructions).replace(/\r\n/g,"\n").replace(/\s+(?=\d+\.\s)/g,"\n").split(/\n+/).filter(Boolean).forEach(x=>children.push(new Paragraph({children:[new TextRun({text:x.trim(),size:20})],spacing:{after:40}})))}sections.forEach((sec,si)=>{const type=QUESTION_TYPES.find(t=>t.id===sec.typeId),total=sec.questionCount*sec.marksPerQ,marks=sec.marksPerQ===1?`(${total})`:`(${sec.marksPerQ}x${sec.questionCount}=${total})`;children.push(new Paragraph({children:[new TextRun({text:`Q${si+1}. ${sec.heading||type.label}`,bold:true,size:22}),new TextRun({text:`\t${marks}`,bold:true,size:22})],pageBreakBefore:si>0&&sec.pageBreak==="before",spacing:{before:160,after:100}}));sec.questions.forEach((q,qi)=>{children.push(new Paragraph({children:[new TextRun({text:`${questionLabel(qi,header.numbering||"alpha")} `,bold:true}),...htmlToDocxRuns(q.text||"___________________________________")],pageBreakBefore:si>0&&!!sec.questions[qi-1]?.pageBreakAfter,spacing:{after:80}}));if(q.options)q.options.forEach((o,i)=>children.push(new Paragraph({children:[new TextRun({text:`${OPT_ALPHA[i]}) ${plainFromHtml(o||"_______")}`})],indent:{left:360},spacing:{after:40}})));if(sec.typeId==="match"){children.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[new TableRow({children:[new TableCell({children:[new Paragraph({text:"Column A"})]}),new TableCell({children:[new Paragraph({text:"Column B"})]})]}),...(q.matchLeft||[]).map((l,i)=>new TableRow({children:[new TableCell({children:[new Paragraph({text:`${i+1}. ${l||"_______"}`})]}),new TableCell({children:[new Paragraph({text:`${String.fromCharCode(97+i)}. ${(q.matchRight||[])[i]||"_______"}`})]})]}))]}));}})});const blob=await Packer.toBlob(new Document({sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:720,right:720,bottom:720,left:720}}},children}]}));const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`GPS-${header.classVal||"Class"}-${header.subject||"Subject"}.docx`;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
  return <div className="app-shell"><header className="topbar"><div className="brand"><strong>GPS QP Builder</strong><span className="user-chip">{user.name||user.username}</span></div><div className="top-actions"><button className={view==="build"?"active":""} onClick={()=>setView("build")}>Build</button><button className={view==="preview"?"active":""} onClick={()=>setView("preview")}>Preview</button><button onClick={newPaper}>New Paper</button><button onClick={()=>setShowDraftPanel(d=>!d)}>Drafts {savedDrafts.length?`(${savedDrafts.length})`:""}</button><button onClick={()=>setShowTemplatePanel(d=>!d)}>Templates {templates.length?`(${templates.length})`:""}</button>{user.role==="admin"&&<button onClick={onAdmin}>Admin</button>}<button onClick={onAccount}>Account</button><button onClick={()=>setDark(d=>!d)} title="Toggle dark mode">{dark?"☀":"◐"}</button><button onClick={onLogout}>Logout</button></div></header>{notice&&<div className="toast">{notice}</div>}<main className="main-grid">{view==="build"?<div className="builder-panel"><Card><h2>Paper Details</h2><div className="two-col"><Input label="School" value={header.school} onChange={v=>setHeader(h=>({...h,school:v}))}/><Input label="Location" value={header.location} onChange={v=>setHeader(h=>({...h,location:v}))}/><Input label="Exam" value={header.exam} onChange={v=>setHeader(h=>({...h,exam:v}))}/><Input label="Session" value={header.session} onChange={v=>setHeader(h=>({...h,session:v}))}/><Input label="Class" value={header.classVal} onChange={v=>setHeader(h=>({...h,classVal:v}))}/><Input label="Subject" value={header.subject} onChange={v=>setHeader(h=>({...h,subject:v}))}/><Input label="Duration" value={header.duration} onChange={v=>setHeader(h=>({...h,duration:v}))}/><Input label="Maximum Marks" type="number" value={header.maxMarks} onChange={v=>setHeader(h=>({...h,maxMarks:v}))}/><Input label="Date" type="date" value={header.date} onChange={v=>setHeader(h=>({...h,date:v}))}/><div><label className="field-label">Language</label><select value={header.language} onChange={e=>setHeader(h=>({...h,language:e.target.value}))} className="field-input"><option>English</option><option>Hindi</option><option>Sanskrit</option></select></div><div><label className="field-label">Question numbering</label><select value={header.numbering} onChange={e=>setHeader(h=>({...h,numbering:e.target.value}))} className="field-input"><option value="alpha">A, B, C</option><option value="numeric">1, 2, 3</option><option value="roman">(i), (ii), (iii)</option></select></div></div><Textarea label="Instructions" value={header.instructions} onChange={v=>setHeader(h=>({...h,instructions:v}))} placeholder="1. Read all questions carefully.\n2. Answer all questions." rows={4}/></Card><div className="toolbar-row"><Btn onClick={()=>setSections(a=>[...a,makeSection()])}>+ Add Section</Btn><Btn variant="secondary" onClick={saveDraft}>Save Draft</Btn><Btn variant="secondary" onClick={saveTemplate}>Save Template</Btn></div>{sections.map((s,i)=><SectionEditor key={s.id} section={s} index={i} total={sections.length} maxAllowed={remaining+s.questionCount*s.marksPerQ} onChange={v=>updateSection(i,v)} onDelete={()=>{if(sections.length===1){notify("Keep at least one section.");return}setSections(a=>a.filter((_,j)=>j!==i))}} onMoveUp={()=>setSections(a=>{const x=a[i];if(i===0)return a;return [...a.slice(0,i-1),x,a[i-1],...a.slice(i+1)]})} onMoveDown={()=>setSections(a=>{if(i===a.length-1)return a;const x=a[i];return [...a.slice(0,i),a[i+1],x,...a.slice(i+2)]})} onDuplicate={()=>duplicate(i)} notify={notify} language={header.language} numbering={header.numbering}/>)}</div>:<div className="preview-panel"><div className="preview-controls"><label>Zoom <input type="range" min="0.7" max="1.2" step="0.05" value={zoom} onChange={e=>setZoom(parseFloat(e.target.value))}/></label><select value={printSize} onChange={e=>setPrintSize(e.target.value)}><option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option></select><Btn variant="secondary" onClick={exportPDF}>Export PDF</Btn><Btn variant="secondary" onClick={exportDOCX}>Export DOCX</Btn></div><PagedPreview header={header} sections={sections} zoom={zoom} printSize={printSize}/></div>}{showDraftPanel&&<Card className="floating-panel"><div className="modal-head"><h3>Drafts</h3><button className="icon-btn" onClick={()=>setShowDraftPanel(false)}>✕</button></div>{savedDrafts.length?savedDrafts.map(d=><div className="list-row" key={d.id}><span>{d.name}</span><span><Btn small variant="secondary" onClick={()=>loadDraft(d)}>Load</Btn><Btn small variant="danger" onClick={()=>setSavedDrafts(a=>a.filter(x=>x.id!==d.id))}>Delete</Btn></span></div>):<div className="small-muted">No drafts yet.</div>}</Card>}{showTemplatePanel&&<Card className="floating-panel"><div className="modal-head"><h3>Templates</h3><button className="icon-btn" onClick={()=>setShowTemplatePanel(false)}>✕</button></div>{templates.length?templates.map(t=><div className="list-row" key={t.id}><span>{t.name}</span><span><Btn small variant="secondary" onClick={()=>loadTemplate(t)}>Load</Btn><Btn small variant="danger" onClick={()=>setTemplates(a=>a.filter(x=>x.id!==t.id))}>Delete</Btn></span></div>):<div className="small-muted">No templates yet.</div>}</Card>}</main></div>;
}

function SetupScreen({onDone}) {
  const [admin,setAdmin]=useState({name:"masterdev",username:"masterdev",email:"",password:"masterdev"}); const [busy,setBusy]=useState(false); const [key,setKey]=useState(""); const [error,setError]=useState("");
  async function submit(e){e.preventDefault();setError("");if(!key)return setError("Enter the setup key configured in Cloudflare.");setBusy(true);try{const r=await fetch("/api/setup/bootstrap",{method:"POST",headers:{"Content-Type":"application/json","x-setup-key":key},body:JSON.stringify({admin})});const d=await r.json().catch(()=>({}));if(!r.ok)throw new Error(d.error||"Setup failed.");onDone()}catch(e){setError(e.message)}finally{setBusy(false)}}
  return <div className="auth-screen"><form className="auth-card setup-card" onSubmit={submit}><img src={gpsLogo} className="auth-logo" alt="Goodwill Public School"/><h1>School Setup</h1><p className="small-muted">Create the single school administrator account. Teacher accounts can be added later from the Admin Panel.</p><Input label="Setup Key" type="password" value={key} onChange={setKey}/><div className="setup-admin"><h3>Administrator</h3><div className="two-col"><Input label="Name" value={admin.name} disabled onChange={v=>setAdmin(a=>({...a,name:v}))}/><Input label="Username" value={admin.username} disabled onChange={v=>setAdmin(a=>({...a,username:v}))}/><Input label="Email" value={admin.email} disabled onChange={v=>setAdmin(a=>({...a,email:v}))}/><Input label="Password" type="password" value={admin.password} disabled onChange={v=>setAdmin(a=>({...a,password:v}))}/></div></div>{error&&<div className="auth-error">{error}</div>}<Btn full disabled={busy}>{busy?"Setting up…":"Create Admin Account"}</Btn></form></div>;
}

function LoginScreen({onLogin}) {
  const [username,setUsername]=useState(""); const [password,setPassword]=useState(""); const [busy,setBusy]=useState(false); const [error,setError]=useState("");
  async function submit(e){e.preventDefault();setError("");setBusy(true);try{const r=await fetch("/api/auth/login",{method:"POST",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({username,password})});const d=await r.json().catch(()=>({}));if(!r.ok)throw new Error(d.error||"Login failed.");onLogin(d.user,d.workspace)}catch(err){setError(err.message||"Login failed.")}finally{setBusy(false)}}
  return <div className="auth-screen"><form className="auth-card" onSubmit={submit}><img src={gpsLogo} className="auth-logo" alt="Goodwill Public School"/><h1>GPS QP Builder</h1><p className="small-muted">School Question Paper Builder</p><Input label="Username" value={username} onChange={setUsername} autoComplete="username"/><Input label="Password" type="password" value={password} onChange={setPassword} autoComplete="current-password"/>{error&&<div className="auth-error">{error}</div>}<Btn full disabled={busy}>{busy?"Signing in…":"Sign in"}</Btn></form></div>;
}

function AccountPanel({user,onClose,onUserUpdated,notify}) {
  const [name,setName]=useState(user.name||""); const [username,setUsername]=useState(user.username||""); const [email,setEmail]=useState(user.email||""); const [currentPassword,setCurrentPassword]=useState(""); const [newPassword,setNewPassword]=useState(""); const [confirmPassword,setConfirmPassword]=useState(""); const [busy,setBusy]=useState(false);
  async function save(e){e.preventDefault();setBusy(true);try{const r=await fetch("/api/auth/account",{method:"PUT",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({name,username,email})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Could not update account.");onUserUpdated(d.user);notify("Account details updated.");}catch(e){notify(e.message)}finally{setBusy(false)}}
  async function changePassword(e){e.preventDefault();if(newPassword!==confirmPassword){notify("New passwords do not match.");return}setBusy(true);try{const r=await fetch("/api/auth/password",{method:"PUT",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({currentPassword,newPassword})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Could not change password.");setCurrentPassword("");setNewPassword("");setConfirmPassword("");notify("Password changed successfully.");}catch(e){notify(e.message)}finally{setBusy(false)}}
  return <div className="modal-backdrop"><div className="modal-card"><div className="modal-head"><h2>My Account</h2><button className="icon-btn" onClick={onClose}>✕</button></div><form onSubmit={save}><Input label="Name" value={name} onChange={setName}/><Input label="Username" value={username} onChange={setUsername}/><Input label="Email" value={email} onChange={setEmail}/><Btn disabled={busy}>Save Account Details</Btn></form><hr/><h3>Change Password</h3><form onSubmit={changePassword}><Input label="Current Password" type="password" value={currentPassword} onChange={setCurrentPassword}/><Input label="New Password" type="password" value={newPassword} onChange={setNewPassword}/><Input label="Confirm New Password" type="password" value={confirmPassword} onChange={setConfirmPassword}/><Btn variant="secondary" disabled={busy}>Change Password</Btn></form></div></div>;
}

function AdminPanel({onClose,notify}) {
  const [users,setUsers]=useState([]); const [loading,setLoading]=useState(true); const [form,setForm]=useState({name:"",username:"",email:"",role:"teacher",password:""}); const [bulk,setBulk]=useState(""); const [busy,setBusy]=useState(false);
  async function load(){setLoading(true);try{const r=await fetch("/api/admin/users",{credentials:"include"});const d=await r.json();if(!r.ok)throw new Error(d.error||"Could not load users.");setUsers(d.users||[])}catch(e){notify(e.message)}finally{setLoading(false)}}
  useEffect(()=>{load()},[]);
  async function create(e){e.preventDefault();setBusy(true);try{const r=await fetch("/api/admin/users",{method:"POST",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({...form,role:"teacher"})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Could not create user.");setForm({name:"",username:"",email:"",role:"teacher",password:""});notify("Teacher account created.");load()}catch(e){notify(e.message)}finally{setBusy(false)}}
  async function bulkCreate(){const rows=bulk.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);if(!rows.length)return;setBusy(true);let ok=0;try{for(const row of rows){const [name,username,password,email]=row.split(",").map(x=>x.trim());if(!name||!username||!password)continue;const r=await fetch("/api/admin/users",{method:"POST",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({name,username,password,email,role:"teacher"})});if(r.ok)ok++;}setBulk("");notify(`Created ${ok} teacher account${ok===1?"":"s"}.`);load()}finally{setBusy(false)}}
  async function toggle(u){try{const r=await fetch(`/api/admin/users/${u.id}`,{method:"PATCH",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({active:!u.active})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Update failed.");load()}catch(e){notify(e.message)}}
  async function reset(u){const pw=prompt(`Temporary password for ${u.username}:`,"GPS@ChangeMe1");if(!pw)return;try{const r=await fetch(`/api/admin/users/${u.id}/reset-password`,{method:"POST",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify({password:pw})});const d=await r.json();if(!r.ok)throw new Error(d.error||"Reset failed.");notify(`Password reset. Give the temporary password to ${u.name||u.username}.`)}catch(e){notify(e.message)}}
  async function impersonate(u){try{const r=await fetch(`/api/admin/users/${u.id}/impersonate`,{method:"POST",credentials:"include"});const d=await r.json();if(!r.ok)throw new Error(d.error||"Could not switch account.");window.location.reload()}catch(e){notify(e.message)}}
  return <div className="modal-backdrop"><div className="modal-card admin-modal"><div className="modal-head"><h2>Admin Panel</h2><button className="icon-btn" onClick={onClose}>✕</button></div><h3>Accounts</h3>{loading?<div className="small-muted">Loading…</div>:<div className="admin-users">{users.map(u=><div className="admin-user" key={u.id}><div><b>{u.name||u.username}</b><div className="small-muted">@{u.username} · {u.role} · {u.active?"Active":"Disabled"}</div></div><div className="toolbar-row"><Btn small variant="secondary" onClick={()=>impersonate(u)}>Login as</Btn><Btn small variant="secondary" onClick={()=>reset(u)}>Reset password</Btn><Btn small variant={u.active?"danger":"success"} onClick={()=>toggle(u)}>{u.active?"Disable":"Enable"}</Btn></div></div>)}</div>}<hr/><h3>Create teacher account</h3><form onSubmit={create}><div className="two-col"><Input label="Name" value={form.name} onChange={v=>setForm(f=>({...f,name:v}))}/><Input label="Username" value={form.username} onChange={v=>setForm(f=>({...f,username:v}))}/><Input label="Email" value={form.email} onChange={v=>setForm(f=>({...f,email:v}))}/><Input label="Temporary Password" type="password" value={form.password} onChange={v=>setForm(f=>({...f,password:v}))}/></div><div className="small-muted" style={{marginBottom:8}}>New accounts created here are teacher accounts. There is only one administrator account.</div><Btn disabled={busy}>Create Teacher Account</Btn></form><div className="bulk-box"><h3>Bulk create teachers</h3><div className="small-muted">One account per line: Name, username, temporary password, email</div><textarea value={bulk} onChange={e=>setBulk(e.target.value)} rows={7} className="field-input" placeholder={'Anita Sharma, anita, TempPass123, anita@example.com\nPooja Verma, pooja, TempPass123, pooja@example.com'}/><Btn variant="secondary" disabled={busy||!bulk.trim()} onClick={bulkCreate}>Create All Rows</Btn></div></div></div>;
}

export default function App(){
  const [auth,setAuth]=useState(null); const [loading,setLoading]=useState(true); const [setup,setSetup]=useState(false); const [panel,setPanel]=useState(null);
  useEffect(()=>{(async()=>{try{const r=await fetch("/api/auth/me",{credentials:"include"});if(r.ok){const d=await r.json();setAuth(d);setLoading(false);return;}const s=await fetch("/api/setup/status",{cache:"no-store"});const sd=await s.json();setSetup(!!sd.needsSetup)}catch(e){setSetup(false)}finally{setLoading(false)}})()},[]);
  async function afterSetup(){const r=await fetch("/api/setup/status",{cache:"no-store"});const d=await r.json();if(!d.needsSetup)setSetup(false)}
  async function logout(){await fetch("/api/auth/logout",{method:"POST",credentials:"include"});setAuth(null)}
  function handleUserUpdated(user){setAuth(a=>({...a,user}))}
  if(loading)return <div className="auth-screen"><div className="auth-card"><h2>Loading…</h2></div></div>;
  if(setup)return <SetupScreen onDone={afterSetup}/>;
  if(!auth)return <LoginScreen onLogin={(user,workspace)=>setAuth({user,workspace})}/>;
  const acting=auth.user?.impersonator; return <>{acting&&<div className="impersonation-bar">⚠ Viewing as <b>{auth.user.name||auth.user.username}</b> · Admin session active <button onClick={async()=>{await fetch("/api/admin/stop-impersonation",{method:"POST",credentials:"include"});window.location.reload()}}>Return to Admin</button></div>}<BuilderApp user={auth.user} initialWorkspace={auth.workspace} onWorkspaceChange={async data=>{await fetch("/api/workspace",{method:"PUT",headers:{"Content-Type":"application/json"},credentials:"include",body:JSON.stringify(data)})}} onLogout={logout} onAccount={()=>setPanel("account")} onAdmin={()=>setPanel("admin")}/>{panel==="account"&&<AccountPanel user={auth.user} onClose={()=>setPanel(null)} onUserUpdated={handleUserUpdated} notify={m=>{}}/>}{panel==="admin"&&auth.user.role==="admin"&&<AdminPanel onClose={()=>setPanel(null)} notify={m=>{}}/>}</>;
}
