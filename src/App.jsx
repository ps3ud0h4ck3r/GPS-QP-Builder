import { useEffect, useMemo, useRef, useState } from "react";
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

const DEFAULT_HEADER = {
  school: "Goodwill Public School", location: "Uttam Nagar", exam: "Half Early Examination", session: "2026–27",
  classVal: "VI", subject: "Computer", duration: "1.5 Hours", maxMarks: "30", date: "", instructions: "", logoUrl: null,
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
      <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>command("italic")}><i>I</i></button>
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

function QuestionEditor({ q,index,typeId,onChange,showDiagram,onMoveUp,onMoveDown,onDelete,onReorder }) {
  const type=QUESTION_TYPES.find(t=>t.id===typeId); const [showDiag,setShowDiag]=useState(!!q.diagram);
  const upd=p=>onChange({...q,...p});
  return <div className="question-editor" draggable onDragStart={e=>e.dataTransfer.setData("text/plain",q.id)} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const source=e.dataTransfer.getData("text/plain");if(source&&source!==q.id)onReorder(source,q.id)}}>
    <div className="question-editor-head"><span className="drag-handle" title="Drag to reorder">⠿</span><Badge color="#ede9fe" textColor="#5b21b6">{ALPHA[index]}</Badge><span className="small-muted">{type.label}</span><span style={{marginLeft:"auto"}}><Btn onClick={onMoveUp} variant="ghost" small title="Move up">↑</Btn><Btn onClick={onMoveDown} variant="ghost" small title="Move down">↓</Btn><Btn onClick={onDelete} variant="danger" small title="Delete question">✕</Btn></span></div>
    {typeId==="match" ? <div className="two-col"><div><div className="field-label">Column A</div>{q.matchLeft.map((v,i)=><input key={i} value={v} placeholder={`${i+1}.`} onChange={e=>{const a=[...q.matchLeft];a[i]=e.target.value;upd({matchLeft:a})}} className="field-input compact"/>)}</div><div><div className="field-label">Column B</div>{q.matchRight.map((v,i)=><input key={i} value={v} placeholder={`${String.fromCharCode(97+i)}.`} onChange={e=>{const a=[...q.matchRight];a[i]=e.target.value;upd({matchRight:a})}} className="field-input compact"/>)}</div></div> : <>
      <RichTextEditor value={q.text} onChange={text=>upd({text})} placeholder={typeId==="fill"?"Write statement with ___ for blank":typeId==="tf"?"Write a true/false statement":"Write your question here"}/>
      {type.hasOptions && <div style={{marginTop:8}}>{Array.from({length:type.hasOptions},(_,i)=><div className="option-row" key={i}><span>{OPT_ALPHA[i]})</span><input value={q.options[i]} onChange={e=>{const a=[...q.options];a[i]=e.target.value;upd({options:a})}} placeholder={`Option ${OPT_ALPHA[i].toUpperCase()}`} className="field-input compact"/></div>)}</div>}
    </>}
    {showDiagram && type.hasDiagram && <div style={{marginTop:8}}>{!showDiag&&!q.diagram?<button className="link-btn" onClick={()=>setShowDiag(true)}>+ Add Diagram</button>:<DiagramPicker diagram={q.diagram} onChange={d=>{upd({diagram:d});if(!d)setShowDiag(false)}}/>}</div>}
  </div>;
}

function SectionEditor({ section,index,total,onChange,onDelete,onMoveUp,onMoveDown,onDuplicate,maxAllowed,notify }) {
  const [open,setOpen]=useState(true); const [aiLoading,setAiLoading]=useState(false); const [aiTopic,setAiTopic]=useState("");
  const [aiUses,setAiUses]=useState(()=>{try{const d=JSON.parse(localStorage.getItem("qp_ai_usage")||"{}");return d.date===new Date().toISOString().slice(0,10)?Number(d.count)||0:0}catch{return 0}});
  const type=QUESTION_TYPES.find(t=>t.id===section.typeId); const sectionMarks=section.questionCount*section.marksPerQ;
  const upd=p=>onChange({...section,...p});
  function changeType(id){const t=QUESTION_TYPES.find(x=>x.id===id);const maxCount=Math.max(1,Math.floor(maxAllowed/t.defaultMarks)||1);const count=Math.min(section.questionCount,maxCount);upd({typeId:id,marksPerQ:t.defaultMarks,questionCount:count,questions:Array.from({length:count},()=>makeQuestion(id))});}
  function changeCount(n){const desired=Math.max(1,Math.min(20,parseInt(n)||1));const maxCount=Math.max(1,Math.floor(maxAllowed/section.marksPerQ)||1);if(desired*section.marksPerQ>maxAllowed){notify(`This section can have at most ${maxCount} question${maxCount===1?"":"s"}. Maximum paper marks are ${maxAllowed} here.`);return;}const qs=[...section.questions];while(qs.length<desired)qs.push(makeQuestion(section.typeId));upd({questionCount:desired,questions:qs.slice(0,desired)});}
  function changeMarks(n){const m=Math.max(1,Math.min(20,parseInt(n)||1));if(section.questionCount*m>maxAllowed){notify(`These marks would exceed the remaining ${maxAllowed} marks. Reduce questions or marks per question.`);return;}upd({marksPerQ:m});}
  function updateQuestion(i,q){const qs=[...section.questions];qs[i]=q;upd({questions:qs});}
  function reorderQuestion(sourceId,targetId){const qs=[...section.questions],from=qs.findIndex(x=>x.id===sourceId),to=qs.findIndex(x=>x.id===targetId);if(from<0||to<0||from===to)return;const [moving]=qs.splice(from,1);qs.splice(to,0,moving);upd({questions:qs});}
  async function generateAI(){if(!aiTopic.trim())return;if(aiUses>=AI_DAILY_LIMIT){notify(`You have reached the ${AI_DAILY_LIMIT}-generation daily limit on this device.`);return;}setAiLoading(true);try{const count=section.questionCount;const prompt=`Generate ${count} "${type.label}" questions for a school exam on topic: "${aiTopic}". Use clear, age-appropriate school language. Do not repeat questions. Return ONLY valid JSON array, no markdown. ${section.typeId.startsWith("mcq")?`[{"text":"Question?","options":["opt1","opt2","opt3"${section.typeId==="mcq4"?',"opt4"':''}]}]`:section.typeId==="match"?'[{"matchLeft":["term1","term2","term3","term4","term5"],"matchRight":["def1","def2","def3","def4","def5"]}]':'[{"text":"Question or statement here"}]'} Generate exactly ${count} items.`;const res=await fetch("/api/generate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({prompt})});const data=await res.json().catch(()=>({}));if(!res.ok)throw new Error(data.error||`Request failed (${res.status})`);const clean=(data.content||"").replace(/```json|```/g,"").trim();const a=clean.indexOf("["),b=clean.lastIndexOf("]");const parsed=JSON.parse(a>=0&&b>a?clean.slice(a,b+1):clean);if(!Array.isArray(parsed))throw new Error("AI returned an invalid list.");const qs=parsed.slice(0,count).map(item=>({id:uid(),text:item.text||"",diagram:null,options:item.options|| (type.hasOptions?Array.from({length:type.hasOptions},()=>""):null),matchLeft:item.matchLeft|| (section.typeId==="match"?["","","","",""]:null),matchRight:item.matchRight|| (section.typeId==="match"?["","","","",""]:null)}));upd({questions:qs});const next=aiUses+1;setAiUses(next);localStorage.setItem("qp_ai_usage",JSON.stringify({date:new Date().toISOString().slice(0,10),count:next}));}catch(e){notify(`AI generation failed. ${e.message||"Please fill manually."}`)}setAiLoading(false);}
  return <Card className="section-card" style={{marginBottom:16}}>
    <div className="section-head"><div className="drag-handle section-drag" draggable onDragStart={e=>e.dataTransfer.setData("section-id",section.id)} title="Drag to reorder section">⠿</div><div className="section-arrows"><button onClick={onMoveUp} disabled={index===0}>▲</button><button onClick={onMoveDown} disabled={index===total-1}>▼</button></div><Badge color="#fef3c7" textColor="#92400e">Q{index+1}</Badge><div className="section-title">{section.heading||type.label}</div><Badge color="#f0fdf4" textColor="#166534">{sectionMarks} marks</Badge><Btn onClick={onDuplicate} variant="secondary" small title="Duplicate this section">Duplicate</Btn><button className="icon-btn" onClick={()=>setOpen(o=>!o)}>{open?"▾":"▸"}</button><button className="icon-btn danger-text" onClick={onDelete}>✕</button></div>
    {open&&<>
      <div className="qp-section-grid section-config"><div><label className="field-label">Type</label><select value={section.typeId} onChange={e=>changeType(e.target.value)} className="field-input"><option value="fill">Fill in the Blanks</option>{QUESTION_TYPES.slice(1).map(t=><option key={t.id} value={t.id}>{t.label}</option>)}</select></div><div><label className="field-label">No. of Questions</label><input type="number" min="1" max="20" value={section.questionCount} onChange={e=>changeCount(e.target.value)} className="field-input"/></div><div><label className="field-label">Marks/Question</label><input type="number" min="1" max="20" value={section.marksPerQ} onChange={e=>changeMarks(e.target.value)} className="field-input"/></div></div>
      <div style={{marginBottom:14}}><label className="field-label">Custom Heading (optional)</label><input value={section.heading} onChange={e=>upd({heading:e.target.value})} placeholder='e.g. "Section A" or leave blank to use type name' className="field-input"/></div>
      <div className="ai-box"><span className="ai-label">AI Generate</span><span className="small-muted">{AI_DAILY_LIMIT-aiUses} left today</span><input value={aiTopic} onChange={e=>setAiTopic(e.target.value)} placeholder="Topic e.g. Photosynthesis" className="field-input compact ai-topic"/><Btn onClick={generateAI} disabled={aiLoading} small>{aiLoading?"...":"Generate"}</Btn></div>
      <div>{section.questions.map((q,i)=><QuestionEditor key={q.id} q={q} index={i} typeId={section.typeId} onChange={nq=>updateQuestion(i,nq)} onReorder={reorderQuestion} showDiagram onMoveUp={()=>{if(i>0){const qs=[...section.questions];[qs[i-1],qs[i]]=[qs[i],qs[i-1]];upd({questions:qs})}}} onMoveDown={()=>{if(i<section.questions.length-1){const qs=[...section.questions];[qs[i+1],qs[i]]=[qs[i],qs[i+1]];upd({questions:qs})}}} onDelete={()=>{if(section.questions.length===1){notify("A section must contain at least one question.");return;}upd({questions:section.questions.filter(x=>x.id!==q.id),questionCount:section.questionCount-1})}}/>)}</div>
    </>}
  </Card>;
}

function formatInstructions(text){if(!text)return"";const normalized=String(text).replace(/\r\n/g,"\n").trim();if(!normalized)return"";const lines=normalized.replace(/\s+(?=\d+\.\s)/g,"\n").split(/\n+/).map(x=>x.trim()).filter(Boolean);return lines.map(x=>`<div style="margin:2px 0">${escapeHtml(x)}</div>`).join("");}
function renderQuestionText(text){if(!text)return"___________________________________";return hasHtml(text)?text:escapeHtml(text).replace(/\n/g,"<br>");}
function buildPreviewHTML(header,sections,fontSize=15,zoom=1){
  const {school,location,exam,session,classVal,subject,duration,maxMarks,date,instructions,logoUrl}=header;
  const sectionHTML=sections.map((sec,si)=>{const type=QUESTION_TYPES.find(t=>t.id===sec.typeId),heading=sec.heading||type.label,total=sec.questionCount*sec.marksPerQ,marksStr=sec.marksPerQ===1?`(${total})`:`(${sec.marksPerQ}x${sec.questionCount}=${total})`;const qHTML=sec.questions.map((q,qi)=>{let inner="";if(sec.typeId==="match"){inner=`<table class="preview-table"><tr><th>Column A</th><th>Column B</th></tr>${(q.matchLeft||[]).map((l,i)=>`<tr><td>${i+1}. ${escapeHtml(l||"_______")}</td><td>${String.fromCharCode(97+i)}. ${escapeHtml((q.matchRight||[])[i]||"_______")}</td></tr>`).join("")}</table>`}else{inner=`<div>${renderQuestionText(q.text)}</div>`;if(sec.typeId==="tf")inner+=` <span style="margin-left:24px">__________</span>`;if(q.options)inner+=`<div style="margin-left:20px;margin-top:4px">${q.options.map((o,oi)=>`<div>${OPT_ALPHA[oi]}) ${escapeHtml(o||"_______")}</div>`).join("")}</div>`;if(q.diagram)inner+=q.diagram.mode==="space"?`<div class="preview-diagram-space">Diagram</div>`:q.diagram.dataUrl?`<div class="preview-diagram"><img src="${q.diagram.dataUrl}"/></div>`:"";}return `<div class="preview-question"><span class="question-label">${ALPHA[qi]}.</span><div class="question-body">${inner}</div></div>`}).join("");return `<div class="preview-section"><div class="preview-section-head"><span>Q${si+1}. ${escapeHtml(heading)}</span><span>${marksStr}</span></div>${qHTML}</div>`}).join("");
  return `<div class="paper" style="--paper-font:${fontSize}px;transform:scale(${zoom});transform-origin:top center"><div class="paper-header">${logoUrl?`<img class="paper-logo" src="${logoUrl}"/>`:""}<div class="school-name">${escapeHtml(school)}</div><div class="school-location">${escapeHtml(location)}</div><div class="exam-line">${escapeHtml(exam)} ${escapeHtml(session)}</div><div class="class-line">CLASS ${escapeHtml(classVal)}</div>${date?`<div class="date-line">Date: ${escapeHtml(date)}</div>`:""}</div><div class="meta-row"><span>Time: ${escapeHtml(duration)}</span><span>Subject: ${escapeHtml(subject)}</span><span>M.M.: ${escapeHtml(maxMarks)}</span></div>${instructions?`<div class="instructions"><strong>Instructions:</strong>${formatInstructions(instructions)}</div>`:""}${sectionHTML}<div class="paper-end">***************************************<br/><span>1</span></div></div>`;
}

function MarksSummary({sections,maxMarks}){const total=totalMarks(sections),max=parseInt(maxMarks)||0,diff=max-total,over=total>max,pct=max?Math.min(100,total/max*100):0;return <div className={`marks-summary ${over?"over":""}`}><div><span>{over?`${total-max} marks over limit`:diff===0?"Paper complete":`${diff} marks remaining`}</span><b>{total} / {max}</b></div><div className="progress"><div style={{width:`${pct}%`}}/></div></div>}

function htmlToDocxRuns(html=""){
  const doc=new DOMParser().parseFromString(`<div>${html}</div>`,"text/html");
  const out=[];
  function walk(node,styles={}){node.childNodes.forEach(n=>{if(n.nodeType===Node.TEXT_NODE){if(n.textContent)out.push(new TextRun({text:n.textContent,bold:!!styles.bold,italics:!!styles.italics,superScript:!!styles.sup,subScript:!!styles.sub}));}else if(n.nodeType===Node.ELEMENT_NODE){const tag=n.tagName.toLowerCase();if(tag==="br"){out.push(new TextRun({text:"\n"}));return;}walk(n,{bold:styles.bold||tag==="b"||tag==="strong",italics:styles.italics||tag==="i"||tag==="em",sup:styles.sup||tag==="sup",sub:styles.sub||tag==="sub"});}})}
  walk(doc.body); return out.length?out:[new TextRun("")];
}
function makeDocParagraph(html,opts={}){return new Paragraph({children:htmlToDocxRuns(html),spacing:{after:opts.after??100},alignment:opts.alignment});}
async function createDocxBlob(header,sections){const children=[];if(header.logoUrl){/* DOCX cannot reliably embed data URLs in all browsers without extra conversion, so the school name remains as the text header. */}children.push(new Paragraph({children:[new TextRun({text:header.school||"Question Paper",bold:true,size:30})],alignment:AlignmentType.CENTER,spacing:{after:50}}));if(header.location)children.push(new Paragraph({children:[new TextRun({text:header.location,size:20})],alignment:AlignmentType.CENTER,spacing:{after:50}}));children.push(new Paragraph({children:[new TextRun({text:`${header.exam||""} ${header.session||""}`,bold:true,size:22})],alignment:AlignmentType.CENTER,spacing:{after:50}}));children.push(new Paragraph({children:[new TextRun({text:`CLASS ${header.classVal||""}`,bold:true,size:20})],alignment:AlignmentType.CENTER,spacing:{after:120}}));children.push(new Paragraph({children:[new TextRun({text:`Time: ${header.duration||""}     Subject: ${header.subject||""}     M.M.: ${header.maxMarks||""}`,bold:true,size:20})],spacing:{after:180}}));if(header.instructions){children.push(new Paragraph({children:[new TextRun({text:"Instructions:",bold:true,size:20})],spacing:{after:40}}));String(header.instructions).replace(/\r\n/g,"\n").replace(/\s+(?=\d+\.\s)/g,"\n").split(/\n+/).filter(Boolean).forEach(x=>children.push(new Paragraph({children:[new TextRun({text:x.trim(),size:20})],spacing:{after:40}})))}sections.forEach((sec,si)=>{const type=QUESTION_TYPES.find(t=>t.id===sec.typeId),total=sec.questionCount*sec.marksPerQ,marks=sec.marksPerQ===1?`(${total})`:`(${sec.marksPerQ}x${sec.questionCount}=${total})`;children.push(new Paragraph({children:[new TextRun({text:`Q${si+1}. ${sec.heading||type.label}`,bold:true,size:22}),new TextRun({text:`\t${marks}`,bold:true,size:22})],spacing:{before:160,after:100}}));sec.questions.forEach((q,qi)=>{children.push(new Paragraph({children:[new TextRun({text:`${ALPHA[qi]}. `,bold:true}),...htmlToDocxRuns(q.text||"___________________________________")],spacing:{after:80}}));if(q.options)q.options.forEach((o,i)=>children.push(new Paragraph({children:[new TextRun({text:`${OPT_ALPHA[i]}) ${plainFromHtml(o||"_______")}`})],indent:{left:360},spacing:{after:40}})));if(sec.typeId==="match"){children.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[new TableRow({children:[new TableCell({children:[new Paragraph({text:"Column A"})]}),new TableCell({children:[new Paragraph({text:"Column B"})]})]}),...(q.matchLeft||[]).map((l,i)=>new TableRow({children:[new TableCell({children:[new Paragraph({text:`${i+1}. ${l||"_______"}`})]}),new TableCell({children:[new Paragraph({text:`${String.fromCharCode(97+i)}. ${(q.matchRight||[])[i]||"_______"}`})]})]}))]}));}})});const doc=new Document({sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:720,right:720,bottom:720,left:720}}},children}]});return Packer.toBlob(doc);}

export default function App(){
  const [view,setView]=useState("build"); const [header,setHeader]=useState(DEFAULT_HEADER); const [sections,setSections]=useState([makeSection()]); const [savedDrafts,setSavedDrafts]=useState(()=>{try{return JSON.parse(localStorage.getItem("qp_drafts")||"[]")}catch{return[]}}); const [draftName,setDraftName]=useState(""); const [showDraftPanel,setShowDraftPanel]=useState(false); const [dark,setDark]=useState(()=>localStorage.getItem("qp_dark")==="1"); const [printFont,setPrintFont]=useState("medium"); const [zoom,setZoom]=useState(1); const [notice,setNotice]=useState(""); const logoRef=useRef(); const previewRef=useRef();
  useEffect(()=>{document.documentElement.classList.toggle("dark",dark);localStorage.setItem("qp_dark",dark?"1":"0")},[dark]);
  function notify(msg){setNotice(msg);window.clearTimeout(notify.t);notify.t=window.setTimeout(()=>setNotice(""),3500)}
  const max=parseInt(header.maxMarks)||0; const current=totalMarks(sections);
  function updateMaxMarks(v){const n=Math.max(1,parseInt(v)||1);if(n<current){notify(`Max marks cannot be set below the current ${current} marks. Remove marks first.`);return;}setHeader(h=>({...h,maxMarks:String(n)}));}
  function addSection(){if(current>=max){notify(`You have used all ${max} marks. Increase Max Marks or reduce another section first.`);return;}setSections(s=>[...s,makeSection()]);}
  function updateSection(i,sec){setSections(s=>s.map((x,idx)=>idx===i?sec:x));}
  function deleteSection(i){if(sections.length===1){notify("Keep at least one section.");return;}setSections(s=>s.filter((_,idx)=>idx!==i));}
  function moveSection(i,j){setSections(s=>{const a=[...s];if(j<0||j>=a.length)return a;[a[i],a[j]]=[a[j],a[i]];return a})}
  function dragSection(id,targetId){if(id===targetId)return;setSections(s=>{const a=[...s],from=a.findIndex(x=>x.id===id),to=a.findIndex(x=>x.id===targetId);if(from<0||to<0)return a;const [item]=a.splice(from,1);a.splice(to,0,item);return a})}
  function duplicateSection(i){const src=sections[i];const needed=src.questionCount*src.marksPerQ;if(current+needed>max){notify(`Duplicating this section needs ${needed} more marks, but only ${Math.max(0,max-current)} are available.`);return;}const copy=JSON.parse(JSON.stringify(src));copy.id=uid();copy.heading=src.heading?`${src.heading} Copy`:"";copy.questions=copy.questions.map(q=>({...q,id:uid()}));setSections(s=>[...s.slice(0,i+1),copy,...s.slice(i+1)]);}
  function saveDraft(){const name=draftName.trim()||`Draft ${new Date().toLocaleDateString("en-IN")}`,draft={name,ts:Date.now(),header,sections};const drafts=[draft,...savedDrafts.filter(d=>d.name!==name)].slice(0,10);setSavedDrafts(drafts);localStorage.setItem("qp_drafts",JSON.stringify(drafts));setDraftName("");notify(`Saved: "${name}"`)}
  function loadDraft(d){setHeader(d.header);setSections(d.sections);setShowDraftPanel(false);setView("build")}
  function deleteDraft(ts){const d=savedDrafts.filter(x=>x.ts!==ts);setSavedDrafts(d);localStorage.setItem("qp_drafts",JSON.stringify(d))}
  function handleLogo(e){const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>setHeader(h=>({...h,logoUrl:ev.target.result}));r.readAsDataURL(f)}
  function newPaper(){if(!confirm("Start a new paper? Unsaved changes on this screen will be lost."))return;setHeader(DEFAULT_HEADER);setSections([makeSection()]);setView("build");setShowDraftPanel(false);window.scrollTo({top:0,behavior:"smooth"})}
  const previewHTML=useMemo(()=>buildPreviewHTML(header,sections,PRINT_FONTS[printFont],zoom),[header,sections,printFont,zoom]);
  function filename(ext){return `${(header.school||"Question-Paper").replace(/[^a-z0-9]+/gi,"-")}-${(header.classVal||"").replace(/[^a-z0-9]+/gi,"-")}-${(header.subject||"").replace(/[^a-z0-9]+/gi,"-")}.${ext}`.replace(/-+/g,"-")}
  async function pdfBlob(){const holder=document.createElement("div");holder.style.position="fixed";holder.style.left="-100000px";holder.style.top="0";holder.style.width="680px";holder.innerHTML=buildPreviewHTML(header,sections,PRINT_FONTS[printFont],1);document.body.appendChild(holder);try{return await html2pdf().set({margin:0,image:{type:"jpeg",quality:.98},html2canvas:{scale:2,useCORS:true,backgroundColor:"#fff"},jsPDF:{unit:"mm",format:"a4",orientation:"portrait"},pagebreak:{mode:["css","legacy"]}}).from(holder.firstElementChild).outputPdf("blob")}finally{holder.remove()}}
  async function downloadPDF(){const blob=await pdfBlob();const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=filename("pdf");a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
  async function downloadDOCX(){try{const blob=await createDocxBlob(header,sections);const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=filename("docx");a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}catch(e){notify(`DOCX export failed. ${e.message||"Please try again."}`)}}
  async function sharePDF(){try{const blob=await pdfBlob();const file=new File([blob],filename("pdf"),{type:"application/pdf"});if(navigator.share&&navigator.canShare?.({files:[file]})){await navigator.share({title:`${header.school} Question Paper`,text:`${header.subject} Class ${header.classVal} Question Paper`,files:[file]});return;}notify("Your browser cannot attach the PDF directly. Use WhatsApp or Email below.");}catch(e){if(e.name!=="AbortError")notify("Sharing was cancelled or unavailable.")}}
  function shareWhatsApp(){const text=`${header.school}\n${header.exam} ${header.session}\nClass ${header.classVal} | Subject: ${header.subject}\nTime: ${header.duration} | M.M.: ${header.maxMarks}\n\nI have prepared this question paper in GPS QP Builder.`;window.open(`https://wa.me/?text=${encodeURIComponent(text)}`,"_blank")}
  function shareEmail(){const subject=`${header.school} - Class ${header.classVal} ${header.subject} Question Paper`;const body=`${header.school}\n${header.exam} ${header.session}\nClass ${header.classVal} | Subject: ${header.subject}\nTime: ${header.duration} | M.M.: ${header.maxMarks}\n\nPrepared in GPS QP Builder.`;window.location.href=`mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`}
  function resetZoom(){setZoom(1)}

  return <div className="app-shell">
    {notice&&<div className="toast">{notice}</div>}
    <header className="topbar"><div className="brand"><img src="/icon-192.png" alt="GPS"/><strong>GPS QP Builder</strong></div><div className="brand-school"><img src="/gps-logo.png" alt="Goodwill Public School"/></div><div className="top-actions"><button className={view==="build"?"active":""} onClick={()=>setView("build")}>Build</button><button className={view==="preview"?"active":""} onClick={()=>setView("preview")}>Preview</button><button onClick={newPaper}>New Paper</button><button onClick={()=>setShowDraftPanel(d=>!d)}>Drafts {savedDrafts.length?`(${savedDrafts.length})`:""}</button><button onClick={()=>setDark(d=>!d)} title="Toggle dark mode">{dark?"☀":"◐"}</button></div></header>
    {showDraftPanel&&<div className="draft-panel"><div className="narrow"><div className="panel-title">Saved Drafts</div><div className="toolbar-row"><input value={draftName} onChange={e=>setDraftName(e.target.value)} placeholder="Draft name (optional)" className="field-input"/><Btn onClick={saveDraft} small>Save Current</Btn></div>{savedDrafts.length===0?<div className="small-muted">No drafts saved yet.</div>:savedDrafts.map(d=><div className="draft-row" key={d.ts}><div><b>{d.name}</b><div className="small-muted">{new Date(d.ts).toLocaleString("en-IN")}</div></div><span style={{marginLeft:"auto"}}><Btn onClick={()=>loadDraft(d)} small variant="secondary">Load</Btn><Btn onClick={()=>deleteDraft(d.ts)} small variant="danger">Delete</Btn></span></div>)}</div></div>}
    <main className="page">
      {view==="build"&&<>
        <Card><div className="card-title">Paper Details</div><div className="qp-header-grid"><Input label="School Name" value={header.school} onChange={v=>setHeader(h=>({...h,school:v}))}/><Input label="Location" value={header.location} onChange={v=>setHeader(h=>({...h,location:v}))}/><Input label="Examination" value={header.exam} onChange={v=>setHeader(h=>({...h,exam:v}))}/><Input label="Session" value={header.session} onChange={v=>setHeader(h=>({...h,session:v}))}/><Input label="Class" value={header.classVal} onChange={v=>setHeader(h=>({...h,classVal:v}))}/><Input label="Subject" value={header.subject} onChange={v=>setHeader(h=>({...h,subject:v}))}/><Input label="Duration" value={header.duration} onChange={v=>setHeader(h=>({...h,duration:v}))}/><Input label="Max Marks" type="number" value={header.maxMarks} onChange={updateMaxMarks}/><Input label="Exam Date (optional)" value={header.date} onChange={v=>setHeader(h=>({...h,date:v}))}/></div><Textarea label="Instructions (optional)" value={header.instructions} onChange={v=>setHeader(h=>({...h,instructions:v}))} placeholder="e.g. 1. Attempt all questions. 2. No calculators allowed."/><div><label className="field-label">School Logo</label><div className="logo-upload"><div className="logo-placeholder"><img src="/gps-logo.png" alt="GPS logo"/></div><Btn onClick={()=>logoRef.current.click()} variant="secondary" small>{header.logoUrl?"Change Paper Logo":"Add Paper Logo (optional)"}</Btn>{header.logoUrl&&<Btn onClick={()=>setHeader(h=>({...h,logoUrl:null}))} variant="danger" small>Remove</Btn>}</div><input ref={logoRef} type="file" accept="image/*" hidden onChange={handleLogo}/><div className="small-muted" style={{marginTop:5}}>The GPS logo in the top bar is built in. A paper logo is optional and independent.</div></div></Card>
        <MarksSummary sections={sections} maxMarks={header.maxMarks}/>
        <div onDragOver={e=>{if(e.dataTransfer.types.includes("section-id"))e.preventDefault()}} onDrop={e=>{const id=e.dataTransfer.getData("section-id");if(id){e.preventDefault();const target=e.target.closest(".section-card-wrap")?.dataset.sectionId;if(target)dragSection(id,target)}}}>
          {sections.map((sec,i)=><div key={sec.id} className="section-card-wrap" data-section-id={sec.id}><SectionEditor section={sec} index={i} total={sections.length} onChange={s=>updateSection(i,s)} onDelete={()=>deleteSection(i)} onMoveUp={()=>moveSection(i,i-1)} onMoveDown={()=>moveSection(i,i+1)} onDuplicate={()=>duplicateSection(i)} maxAllowed={max-(current-sec.questionCount*sec.marksPerQ)} notify={notify}/></div>)}
        </div>
        <button onClick={addSection} className="add-section">+ Add Section</button>
        <Btn onClick={()=>setView("preview")} full>Preview Paper</Btn>
      </>}
      {view==="preview"&&<>
        <div className="preview-controls"><div className="toolbar-row"><Btn onClick={()=>setView("build")} variant="secondary">Edit</Btn><Btn onClick={downloadPDF}>Download PDF</Btn><Btn onClick={downloadDOCX} variant="secondary">Download DOCX</Btn><Btn onClick={sharePDF} variant="secondary">Share PDF</Btn><Btn onClick={()=>window.print()} variant="secondary">Print</Btn></div><div className="toolbar-row"><label className="control-label">Print font <select value={printFont} onChange={e=>setPrintFont(e.target.value)}><option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option></select></label><button className="zoom-btn" onClick={()=>setZoom(z=>Math.max(.7,Math.round((z-.1)*10)/10))}>−</button><span className="zoom-value">{Math.round(zoom*100)}%</span><button className="zoom-btn" onClick={()=>setZoom(z=>Math.min(1.4,Math.round((z+.1)*10)/10))}>+</button><button className="zoom-btn" onClick={resetZoom}>Reset</button></div><div className="toolbar-row share-row"><span className="small-muted">Share:</span><Btn onClick={shareWhatsApp} variant="success" small>WhatsApp</Btn><Btn onClick={shareEmail} variant="secondary" small>Email</Btn></div></div>
        <Card style={{padding:0,overflow:"auto"}}><div ref={previewRef} dangerouslySetInnerHTML={{__html:previewHTML}}/></Card>
      </>}
    </main>
  </div>;
}
