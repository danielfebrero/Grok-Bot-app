"use strict";
// Dashboard RPC — vanilla JS, polling /api/stats toutes les 2 s. Aucune ressource externe.
const $ = (id) => document.getElementById(id);
const REFRESH_MS = 2000;
let lastData = null;

function esc(s){ return String(s==null?"":s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

function fmtTime(ts){
  if(!ts) return "–";
  const d = new Date(ts);
  if(isNaN(d)) return esc(ts);
  return d.toLocaleTimeString("fr-FR",{hour12:false});
}
function shortMethod(m){ if(!m) return "?"; return m.length>60?("…"+m.slice(-58)):m; }

function statusClass(st){
  const s = String(st==null?"":st).toLowerCase();
  if(s==="completed") return "ok";
  if(s==="open") return "warn";
  const n = parseInt(st,10);
  if(isNaN(n)) return "err";
  if(n>=500) return "err"; if(n>=400) return "warn"; if(n>=200&&n<300) return "ok"; return "warn";
}
function legClass(leg){
  if(leg==="inference-effect") return "leg3";
  return (leg&&leg.indexOf("agent")===0)?"leg2":"leg1";
}

async function tick(){
  const days = $("days").value;
  try{
    const r = await fetch(`/api/stats?days=${days}&last=50`,{cache:"no-store"});
    if(!r.ok) throw new Error("HTTP "+r.status);
    const d = await r.json();
    lastData = d;
    render(d);
    $("liveDot").classList.add("ok");
  }catch(e){
    $("liveDot").classList.remove("ok");
  }
}

function render(d){
  $("cTotal").textContent = d.total.toLocaleString("fr-FR");
  $("cRate").textContent  = d.rate_now;
  $("cP95").textContent   = d.latency.p95 + " ms";
  $("cLat").textContent   = `p50 ${d.latency.p50} ms · max ${d.latency.max} ms`;
  $("cErr").textContent   = d.error_rate + " %";
  $("cErrN").textContent  = `${d.error_count} erreur(s)`;
  $("gen").textContent    = "maj " + fmtTime(d.generated_at);
  $("logdir").textContent = d.log_dir;

  drawTimeline(d.timeline);
  drawLegs(d.by_leg);
  drawBars("methods", d.top_methods, false);
  drawStatusBars("statuses", d.status_dist);
  drawTurns(d.inference_turns||[], d.inference_note);
  drawRows(d.last);
}

function drawBars(elId, pairs, _unused){
  const el = $(elId); el.innerHTML="";
  const max = pairs.reduce((m,p)=>Math.max(m,p[1]),0)||1;
  pairs.forEach(([lbl,n])=>{
    const row=document.createElement("div"); row.className="bar";
    row.innerHTML=`<div class="lbl mono" title="${esc(lbl)}">${esc(shortMethod(lbl))}</div>`+
      `<div class="track"><div class="fill" style="width:${(100*n/max).toFixed(1)}%"></div></div>`+
      `<div class="n">${n}</div>`;
    el.appendChild(row);
  });
  if(!pairs.length) el.innerHTML='<div class="dim">aucune donnée</div>';
}

function drawStatusBars(elId, pairs){
  const el = $(elId); el.innerHTML="";
  const max = pairs.reduce((m,p)=>Math.max(m,p[1]),0)||1;
  pairs.forEach(([st,n])=>{
    const cls = statusClass(st);
    const sc = cls==="ok"?"s2":cls==="warn"?"s3":"s4";
    const row=document.createElement("div"); row.className="bar "+sc;
    row.innerHTML=`<div class="lbl mono">${esc(st)}</div>`+
      `<div class="track"><div class="fill" style="width:${(100*n/max).toFixed(1)}%"></div></div>`+
      `<div class="n">${n}</div>`;
    el.appendChild(row);
  });
  if(!pairs.length) el.innerHTML='<div class="dim">aucune donnée</div>';
}

function drawTurns(turns, note){
  if(note) $("infNote").textContent = note;
  $("infCount").textContent = turns.length ? "("+turns.length+")" : "";
  const tb=$("turns"); tb.innerHTML="";
  turns.forEach((t)=>{
    const ids=(t.transcript_entry_ids||[]).join(", ");
    const snip=(t.snippets&&t.snippets[0]&&t.snippets[0].text)||"";
    const st=t.status==="completed"?"terminé":"ouvert";
    const tr=document.createElement("tr");
    tr.innerHTML=`<td class="mono">${fmtTime(t.ended_at||t.started_at)}</td>`+
      `<td><span class="pill ${t.status==="completed"?"ok":"warn"}">${esc(st)}</span></td>`+
      `<td class="mono" title="${esc(t.agent_id||"")}">${esc(t.agent_id||"–")}</td>`+
      `<td class="mono">${esc(t.request_id||"")}</td>`+
      `<td class="mono">${t.duree_ms!=null?t.duree_ms+" ms":"–"}</td>`+
      `<td class="mono clip" title="${esc(ids)}">${esc(ids||"–")}</td>`+
      `<td class="clip" title="${esc(snip)}">${esc(snip||"–")}</td>`;
    tr.onclick=()=>openTurn(t);
    tb.appendChild(tr);
  });
  if(!turns.length){
    tb.innerHTML='<tr><td colspan="7" class="dim">aucun tour corrélé — markAgentRequestStarted/Ended + transcript (SSE ou ListGrokBotTranscriptEntries)</td></tr>';
  }
}

function openTurn(t){
  $("mTitle").textContent = `Tour d'inférence (effets seulement) ${t.request_id||""} → ${t.status||""}`;
  $("mReqH").textContent = t.note||"";
  $("mReqB").textContent = JSON.stringify({
    requestId: t.request_id, agentId: t.agent_id, agent_ids: t.agent_ids,
    started_at: t.started_at, ended_at: t.ended_at, duree_ms: t.duree_ms,
    sse_events: t.sse_events, list_responses: t.list_responses
  }, null, 2);
  $("mResH").textContent = "transcript_entry_ids";
  $("mResB").textContent = JSON.stringify({
    transcript_entry_ids: t.transcript_entry_ids, snippets: t.snippets
  }, null, 2);
  $("modal").classList.remove("hidden");
}

function drawRows(rows){
  const tb=$("rows"); tb.innerHTML="";
  rows.forEach((r,i)=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`<td class="mono">${fmtTime(r.ts)}</td>`+
      `<td><span class="pill ${legClass(r.leg)}">${esc(r.leg||"?")}</span></td>`+
      `<td class="mono" title="${esc(r.method)}">${esc(shortMethod(r.method))}</td>`+
      `<td class="mono">${esc(r.http_method||"")}</td>`+
      `<td><span class="pill ${statusClass(r.status)}">${esc(r.status)}</span></td>`+
      `<td class="mono">${r.duree_ms!=null?r.duree_ms+" ms":""}</td>`;
    tr.onclick=()=>openModal(r);
    tb.appendChild(tr);
  });
  if(!rows.length) tb.innerHTML='<tr><td colspan="6" class="dim">aucun échange capturé</td></tr>';
}

function openModal(r){
  if(r.leg==="inference-effect" || r.kind==="inference-turn"){
    $("mTitle").textContent = `Tour d'inférence (effets seulement) ${r.turn_id||""} [${r.leg}] → ${r.status}`;
  }else{
    const op = r.ws_opcode ? " "+r.ws_opcode : "";
    const proto = r.protocol ? " · "+r.protocol : "";
    const extra = (r.ws_path && r.ws_path !== r.method) ? "  "+r.ws_path : "";
    $("mTitle").textContent = `${r.http_method||""}${op} ${r.method||""}${extra}  [${r.leg||"?"}]${proto} → ${r.status}`;
  }
  $("mReqH").textContent = JSON.stringify(r.req_headers||{},null,2);
  $("mReqB").textContent = r.req_body==null?"(vide)":(typeof r.req_body==="string"?r.req_body:JSON.stringify(r.req_body,null,2));
  $("mResH").textContent = JSON.stringify(r.resp_headers||{},null,2);
  $("mResB").textContent = r.resp_body==null?"(vide)":(typeof r.resp_body==="string"?r.resp_body:JSON.stringify(r.resp_body,null,2));
  $("modal").classList.remove("hidden");
}
$("modalClose").onclick=()=>$("modal").classList.add("hidden");
$("modal").onclick=(e)=>{ if(e.target.id==="modal") $("modal").classList.add("hidden"); };

// ---- graphe timeline en <canvas> natif ----
function drawTimeline(timeline){
  const cv=$("timeline"); const ctx=cv.getContext("2d");
  const W=cv.width, H=cv.height, pad=28;
  ctx.clearRect(0,0,W,H);
  const vals = timeline.map(t=>t[1]);
  const labels = timeline.map(t=>t[0]);
  const maxV = Math.max(1,...vals);
  // grille
  ctx.strokeStyle="#30363d"; ctx.fillStyle="#8b949e"; ctx.font="10px monospace"; ctx.lineWidth=1;
  for(let g=0;g<=4;g++){
    const y=pad+(H-2*pad)*g/4;
    ctx.beginPath(); ctx.moveTo(pad,y); ctx.lineTo(W-6,y); ctx.stroke();
    ctx.fillText(Math.round(maxV*(1-g/4)), 2, y+3);
  }
  if(!vals.length){ ctx.fillText("aucune donnée",W/2-40,H/2); return; }
  const n=vals.length;
  const x=(i)=> pad + (W-pad-6)*(n===1?0.5:i/(n-1));
  const y=(v)=> pad + (H-2*pad)*(1 - v/maxV);
  // aire
  ctx.beginPath(); ctx.moveTo(x(0),H-pad);
  vals.forEach((v,i)=>ctx.lineTo(x(i),y(v)));
  ctx.lineTo(x(n-1),H-pad); ctx.closePath();
  const grad=ctx.createLinearGradient(0,pad,0,H-pad);
  grad.addColorStop(0,"rgba(88,166,255,.35)"); grad.addColorStop(1,"rgba(88,166,255,0)");
  ctx.fillStyle=grad; ctx.fill();
  // ligne
  ctx.beginPath(); ctx.strokeStyle="#58a6ff"; ctx.lineWidth=2;
  vals.forEach((v,i)=>{ const px=x(i),py=y(v); i?ctx.lineTo(px,py):ctx.moveTo(px,py); });
  ctx.stroke();
  // points
  ctx.fillStyle="#58a6ff";
  vals.forEach((v,i)=>{ ctx.beginPath(); ctx.arc(x(i),y(v),2.5,0,7); ctx.fill(); });
  // labels extrêmes
  ctx.fillStyle="#8b949e";
  if(labels.length){
    ctx.fillText((labels[0]||"").slice(11), pad, H-8);
    const lastL=(labels[labels.length-1]||"").slice(11);
    ctx.fillText(lastL, W-6-ctx.measureText(lastL).width, H-8);
  }
}

// ---- donut des legs en <canvas> ----
function drawLegs(byLeg){
  const cv=$("legs"); const ctx=cv.getContext("2d");
  const W=cv.width,H=cv.height; ctx.clearRect(0,0,W,H);
  const entries=Object.entries(byLeg);
  const total=entries.reduce((s,[,v])=>s+v,0);
  const cx=W/2, cy=H/2, r=70, ir=42;
  const colors=["#58a6ff","#bc8cff","#3fb950","#d29922","#f85149"];
  const legend=$("legLegend"); legend.innerHTML="";
  if(!total){ ctx.fillStyle="#8b949e"; ctx.font="12px monospace"; ctx.fillText("aucune donnée",cx-40,cy); return; }
  let a=-Math.PI/2;
  entries.forEach(([k,v],i)=>{
    const frac=v/total, a2=a+frac*2*Math.PI, col=colors[i%colors.length];
    ctx.beginPath(); ctx.moveTo(cx,cy); ctx.arc(cx,cy,r,a,a2); ctx.closePath();
    ctx.fillStyle=col; ctx.fill(); a=a2;
    const sp=document.createElement("span");
    sp.innerHTML=`<i style="background:${col}"></i>${esc(k)} — ${v} (${(100*frac).toFixed(0)}%)`;
    legend.appendChild(sp);
  });
  // trou central
  ctx.beginPath(); ctx.arc(cx,cy,ir,0,7); ctx.fillStyle="#161b22"; ctx.fill();
  ctx.fillStyle="#e6edf3"; ctx.font="bold 20px sans-serif"; ctx.textAlign="center";
  ctx.fillText(total, cx, cy+6); ctx.textAlign="start";
}

$("days").onchange=tick;
tick();
setInterval(tick, REFRESH_MS);
