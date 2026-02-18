// GOLDEN INDEX 完全版（最小ファイル版）
const $ = (id)=>document.getElementById(id);

const state = {
  db: null,
  date: null,
  venue: null,
  race: "1",
  day: 1,
  budget: 5000,
  pro: false,
  expo: {} // reg -> {gyo, nobi, maw, time}
};

function clamp(x, a, b){ return Math.max(a, Math.min(b, x)); }
function mean(arr){ return arr.reduce((s,x)=>s+x,0)/Math.max(1,arr.length); }
function std(arr){
  const m = mean(arr);
  const v = mean(arr.map(x=>(x-m)*(x-m)));
  return Math.sqrt(v) || 1e-9;
}
function zScores(values, invert=false){
  const m = mean(values), s = std(values);
  return values.map(v=>{
    let z = (v - m)/s;
    if(invert) z = -z;
    return clamp(z, -2, 2); // 暴れ防止クリップ
  });
}

function weightsByDay(day){
  // ベース（ユーザー要件に沿って「開催日数でモーター↓ 展示↑」）
  let w = {motor:0.30, st:0.25, course:0.20, local:0.15, expo:0.10};
  const shift = clamp((day-1)*0.12, 0, 0.45); // 最大45%までモーターを展示へ移す
  const delta = w.motor * shift;
  w.motor -= delta;
  w.expo += delta;
  // 正規化
  const sum = Object.values(w).reduce((a,b)=>a+b,0);
  Object.keys(w).forEach(k=>w[k]/=sum);
  return w;
}

function expoScoreFor(boat){
  const e = state.expo[boat.reg] || {gyo:5,nobi:5,maw:5,time:null};
  const gyo = +e.gyo, nobi=+e.nobi, maw=+e.maw;
  // time: 入力があれば「速いほど+」。未入力なら0として扱い、不確実性は後段で補正。
  let tScore = 0;
  if(e.time !== null && e.time !== "" && !Number.isNaN(+e.time)){
    const t = +e.time;
    // 6艇内相対は後でz化するのでここは実数
    tScore = -t; // 小さいほど良い
  }
  // 行/伸/回は 0〜10 を 0〜1 に
  const tri = (gyo + nobi + maw) / 30; // 0..1
  return {tri, tRaw: tScore, hasTime: (e.time !== null && e.time !== "" && !Number.isNaN(+e.time))};
}

function computeScores(boats){
  const motor = boats.map(b=>+b.motor_2);
  const st = boats.map(b=>+b.st_avg);
  const course = boats.map(b=>+b.course2);
  const local = boats.map(b=>+b.local3);

  const expo = boats.map(b=>expoScoreFor(b));
  const expoTri = expo.map(e=>e.tri);
  const expoTRaw = expo.map(e=>e.tRaw);

  const zMotor = zScores(motor);
  const zSt = zScores(st, true); // STは小さいほど良い
  const zCourse = zScores(course);
  const zLocal = zScores(local);

  // 展示：行/伸/回 と タイムを合成（タイム未入力なら tri のみ）
  const zExpoTri = zScores(expoTri);
  // タイムは入力がある艇だけで z、ない艇は0（=平均との差0扱い）にして不確実性補正へ
  const timeVals = expoTRaw.filter(v=>v!==0);
  const useTime = timeVals.length >= 3;
  let zExpoTime = expoTRaw.map(()=>0);
  if(useTime){
    const m = mean(expoTRaw), s = std(expoTRaw);
    zExpoTime = expoTRaw.map(v=>clamp((v-m)/s, -2, 2));
  }
  const zExpo = zExpoTri.map((z,i)=> clamp(0.65*z + 0.35*zExpoTime[i], -2, 2));

  const w = weightsByDay(state.day);
  const raw = boats.map((b,i)=>{
    let z = w.motor*zMotor[i] + w.st*zSt[i] + w.course*zCourse[i] + w.local*zLocal[i] + w.expo*zExpo[i];

    // データ欠損は平均補完＋不確実性補正（欠損がある艇は少し下げる）
    const uncertain = +b.uncertain || 0;
    const expoHasTime = expo[i].hasTime ? 0 : 1;
    const penalty = (uncertain*0.08) + (expoHasTime*0.04);
    z = z * (1 - penalty);

    // 0..100へ
    const score = clamp(50 + 15*z, 0, 100);
    return score;
  });

  return {scores: raw, z: {zMotor,zSt,zCourse,zLocal,zExpo}, w};
}

function grade(score){
  if(score>=75) return "◎";
  if(score>=65) return "○";
  if(score>=55) return "△";
  return "×";
}

function raceType(ranked){
  const top = ranked[0];
  const lane = top.boat.lane;
  if(lane===1 && top.score>=75) return "イン鉄板";
  if(lane===1) return "イン有利";
  if(lane<=3) return "センター主導";
  return "外まくり/差し";
}

function arerIndex(sortedScores){
  // 先頭と中位の差が小さいほど荒れやすい
  const top = sortedScores[0], mid = sortedScores[2] ?? sortedScores[sortedScores.length-1];
  const spread = top - mid; // 大きいほど堅い
  return clamp(100 - spread*3.2, 0, 100);
}
function arerLabel(a){
  if(a>=70) return "高";
  if(a>=45) return "中";
  return "低";
}
function shoubuDo(sortedScores, arer){
  const d12 = (sortedScores[0] - (sortedScores[1]||0));
  const d23 = ((sortedScores[1]||0) - (sortedScores[2]||0));
  let v = d12*3 + d23*1.5 + (50 - arer)*0.6;
  return clamp(v, 0, 100);
}

function pickTickets(ranked){
  // シンプル本命6点：上位3-4艇で組む
  const a = ranked[0]?.boat.lane, b = ranked[1]?.boat.lane, c = ranked[2]?.boat.lane, d = ranked[3]?.boat.lane;
  if(!a||!b||!c) return [];
  const t = [];
  t.push(`${a}-${b}-${c}`);
  t.push(`${a}-${c}-${b}`);
  if(d){ t.push(`${a}-${b}-${d}`); t.push(`${a}-${c}-${d}`); }
  t.push(`${b}-${a}-${c}`);
  if(d) t.push(`${c}-${a}-${b}`); else t.push(`${b}-${c}-${a}`);
  return t.slice(0,6);
}

function allocate(budget, arer){
  // 荒れ高いほど分散
  const base = [0.28,0.20,0.16,0.14,0.12,0.10];
  const spread = clamp((arer-40)/60, 0, 1); // 0..1
  const mix = base.map((x,i)=> x*(1-0.35*spread) + (1/6)*0.35*spread);
  const sum = mix.reduce((a,b)=>a+b,0);
  const yen = mix.map(x=>Math.round((x/sum)*budget/100)*100);
  // 端数調整
  let diff = budget - yen.reduce((a,b)=>a+b,0);
  for(let i=0; diff!==0 && i<yen.length; i++){
    const step = diff>0 ? 100 : -100;
    yen[i]+=step; diff-=step;
  }
  return yen;
}

function renderExpo(boats){
  const rows = boats.map(b=>{
    const e = state.expo[b.reg] || {gyo:5,nobi:5,maw:5,time:""};
    return `
      <div class="box" style="margin:10px 0;">
        <div class="row" style="grid-template-columns:44px 1fr 1fr 1fr 1fr 120px; gap:10px;">
          <div class="pill lane">${b.lane}</div>
          <div>
            <div class="name">${b.name}</div>
            <div class="small">モーター2連 ${b.motor_2}% / ST平均 ${b.st_avg} / コース2連 ${(b.course2*100).toFixed(1)}% / 当地3連 ${(b.local3*100).toFixed(1)}%</div>
          </div>
          <div>
            <div class="small">行き足</div>
            <input type="range" min="0" max="10" step="1" value="${e.gyo}" data-reg="${b.reg}" data-k="gyo">
          </div>
          <div>
            <div class="small">伸び</div>
            <input type="range" min="0" max="10" step="1" value="${e.nobi}" data-reg="${b.reg}" data-k="nobi">
          </div>
          <div>
            <div class="small">回り足</div>
            <input type="range" min="0" max="10" step="1" value="${e.maw}" data-reg="${b.reg}" data-k="maw">
          </div>
          <div>
            <div class="small">タイム</div>
            <input type="number" step="0.01" placeholder="例 6.72" value="${e.time ?? ""}" data-reg="${b.reg}" data-k="time">
          </div>
        </div>
      </div>
    `;
  }).join("");
  $("expoTable").innerHTML = rows;

  // bind
  $("expoTable").querySelectorAll("input").forEach(inp=>{
    inp.addEventListener("input", ()=>{
      const reg = inp.dataset.reg, k = inp.dataset.k;
      state.expo[reg] = state.expo[reg] || {gyo:5,nobi:5,maw:5,time:""};
      state.expo[reg][k] = inp.value;
    });
  });
}

function renderScores(boats, ranked, arer, shoubu, rtype, warn, w){
  const pro = state.pro;

  // ランキング（数値は見せず、視認性優先）
  const max = Math.max(...ranked.map(x=>x.score));
  const min = Math.min(...ranked.map(x=>x.score));
  const span = Math.max(1,(max-min));

  const rows = ranked.map((x,idx)=>{
    const g = grade(x.score);
    const pct = ((x.score - min)/span)*100;
    const blur = (!pro && idx>=2) ? 'style="filter:blur(4px); opacity:.65"' : "";
    return `
      <div class="scoreRow" ${blur}>
        <div class="scoreRank">#${idx+1}</div>
        <div class="scoreMain">
          <div class="scoreName"><span class="laneBadge">${x.boat.lane}</span> ${x.boat.name}</div>
          <div class="scoreBar"><div class="scoreFill" style="width:${pct.toFixed(1)}%"></div></div>
        </div>
        <div class="scoreGrade"><span class="badge grade-${g}">${g}</span></div>
      </div>
    `;
  }).join("");

  $("scoreTable").innerHTML = rows;

  // 結論ダッシュボード
  const dash = `
    <div class="kpi wide"><div class="t">会場</div><div class="v">${state.venue} / ${state.race}R</div></div>
    <div class="kpi"><div class="t">レースタイプ</div><div class="v">${rtype}</div></div>
    <div class="kpi"><div class="t">荒れ度</div><div class="v">${arerLabel(arer)}</div></div>
    <div class="kpi"><div class="t">勝負度</div><div class="v">${shoubu.toFixed(0)}</div></div>
    <div class="kpi"><div class="t">注意</div><div class="v">${warn}</div></div>
  `;
  $("dashboard").innerHTML = dash;
}


function renderTickets(tickets, allocYen){
  if(tickets.length===0){ $("tickets").innerHTML = '<div class="small">買い目生成に失敗（データ不足）</div>'; return; }
  $("tickets").innerHTML = tickets.map((t,i)=>`<div class="pill" style="margin:6px 6px 0 0;">${t}</div>`).join("");
  $("allocation").innerHTML = tickets.map((t,i)=>`<div class="small">${t}：${allocYen[i]||0}円</div>`).join("");
}

function makeShareText(date, venue, race, ranked, arer, shoubu, tickets, allocYen){
  const top = ranked[0];
  const lines = [];
  lines.push(`GOLDEN INDEX｜${date} ${venue} ${race}R`);
  lines.push(`本命: ${top.boat.lane} ${top.boat.name}（指数${top.score.toFixed(1)} ${grade(top.score)}）`);
  lines.push(`荒れ度: ${arerLabel(arer)} / 勝負度: ${shoubu.toFixed(0)}`);
  lines.push(`買い目（6点）`);
  tickets.forEach((t,i)=> lines.push(`${t} ${allocYen[i]||0}円`));
  lines.push(`#競艇 #ボートレース #GOLDENINDEX`);
  return lines.join("\n");
}

function recalc(){
  const boats = state.db.data[state.date][state.venue][state.race];
  renderExpo(boats);

  const {scores, w} = computeScores(boats);
  const ranked = boats.map((b,i)=>({boat:b, score:scores[i]})).sort((a,b)=>b.score-a.score);
  const sortedScores = ranked.map(x=>x.score);
  const arer = arerIndex(sortedScores);
  const rtype = raceType(ranked);
  const shoubu = shoubuDo(sortedScores, arer);

  let warn = "通常";
  if(arer>=75 && shoubu<35) warn = "危険：荒れ警戒";
  if(ranked.some(x=>x.boat.uncertain)) warn = warn==="通常" ? "注意：欠損補完あり" : warn+" / 欠損補完あり";

  renderScores(boats, ranked, arer, shoubu, rtype, warn, w);

  const tickets = pickTickets(ranked);
  const allocYen = allocate(state.budget, arer);
  renderTickets(tickets, allocYen);

  const share = makeShareText(state.date, state.venue, state.race, ranked, arer, shoubu, tickets, allocYen);
  $("shareText").value = share;
}

function bind(){
  $("recalc").addEventListener("click", recalc);
  $("daySel").addEventListener("change", ()=>{ state.day = +$("daySel").value; recalc(); });
  $("budget").addEventListener("change", ()=>{ state.budget = +$("budget").value || 0; recalc(); });
  $("pro").addEventListener("change", ()=>{ state.pro = $("pro").checked; recalc(); });

  $("copy").addEventListener("click", async ()=>{
    const text = $("shareText").value;
    try{
      await navigator.clipboard.writeText(text);
      $("copy").textContent = "コピーした";
      setTimeout(()=>$("copy").textContent="共有文コピー", 900);
    }catch(e){
      // fallback
      $("shareText").classList.remove("hidden");
      $("shareText").select();
      document.execCommand("copy");
      $("shareText").classList.add("hidden");
    }
  });
}

function populate(){
  const dates = Object.keys(state.db.data).sort();
  const dateSel = $("dateSel");
  dateSel.innerHTML = dates.map(d=>`<option value="${d}">${d}</option>`).join("");
  state.date = dates[dates.length-1];
  dateSel.value = state.date;

  const setVenues = ()=>{
    const dayObj = (state.db?.data?.[state.date]) || {};
    const active = new Set(Object.keys(dayObj));
    const venueSel = $("venueSel");

    // 24場を常に表示（開催なしは選択不可にして表記）
    const items = VENUE_MASTER.map(v=>({v, ok: active.has(v)}))
      .sort((a,b)=> (b.ok-a.ok) || (VENUE_MASTER.indexOf(a.v)-VENUE_MASTER.indexOf(b.v)));

    venueSel.innerHTML = '';
    for(const it of items){
      const opt = document.createElement('option');
      opt.value = it.v;
      opt.textContent = it.ok ? it.v : `${it.v}（開催なし）`;
      if(!it.ok) opt.disabled = true;
      venueSel.appendChild(opt);
    }

    // venueが開催なしなら開催中の先頭へ
    if(!active.has(state.venue || '')){
      state.venue = [...active][0] || VENUE_MASTER[0];
    }
    venueSel.value = state.venue;

    // ヘッダーに開催会場一覧
    const el = document.getElementById('venueInfo');
    if(el){
      const list = [...active];
      el.textContent = list.length ? `開催中：${list.join(' / ')}` : 'この日は開催データなし';
    }
  };
  setVenues();

  const setRaces = ()=>{
    const races = Object.keys(state.db.data[state.date][state.venue]).sort((a,b)=>+a-+b);
    const raceSel = $("raceSel");
    raceSel.innerHTML = races.map(r=>`<option value="${r}">${r}</option>`).join("");
    state.race = races[0];
    raceSel.value = state.race;
  };
  setRaces();

  dateSel.addEventListener("change", ()=>{
    state.date = dateSel.value;
    setVenues(); setRaces(); recalc();
  });
  $("venueSel").addEventListener("change", ()=>{
    state.venue = $("venueSel").value;
    setRaces(); recalc();
  });
  $("raceSel").addEventListener("change", ()=>{
    state.race = $("raceSel").value;
    recalc();
  });

  $("meta").textContent = `データ：番組 ${state.db.meta.days}日分 / 生成 ${state.db.meta.generated_at}`;
}

async function main(){
  const res = await fetch("./data/db.json", {cache:"no-store"});
  state.db = await res.json();
  bind();
  populate();
  recalc();
}

main().catch(err=>{
  console.error(err);
  $("meta").textContent = "読み込み失敗：data/db.json を確認";
});
