/* お客様に渡す資料（A4の印刷・PDF保存／PowerPoint）。画面の計算結果を受け取って整形するだけで、数字の計算はしない。 */
(function(){
'use strict';
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const r1=n=>(Math.round(n*10)/10).toLocaleString('ja-JP');
const yen=n=>'¥'+Math.round(n).toLocaleString('ja-JP');
const hr=n=>r1(n)+'時間';
const pct=n=>Math.round(n*100)+'%';
const dateJa=d=>`${d.getFullYear()}年${d.getMonth()+1}月${d.getDate()}日`;
const ymd=d=>`${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
const TITLE='AI業務効率化 診断レポート';
const KIND={diag:'ヒアリングにもとづく試算',h:'HPの内容をもとにした推定'};

const CSS=`
.rp{--rp-ink:#1b2a2f;--rp-mut:#5a6b72;--rp-line:#cfd8db;--rp-acc:#0d6b63;--rp-soft:#eef6f4;
  background:#fff;color:var(--rp-ink);font-family:'Hiragino Kaku Gothic ProN','Hiragino Sans','Yu Gothic UI','Yu Gothic','Meiryo','Noto Sans JP',sans-serif;
  font-size:10.5pt;line-height:1.65;max-width:820px;margin:0 auto;padding:28px 32px;box-shadow:0 1px 8px rgba(0,0,0,.18)}
.rp *{box-sizing:border-box}
.rp-head{border-bottom:3px solid var(--rp-acc);padding-bottom:12px;margin-bottom:16px}
.rp-kicker{font-size:9pt;letter-spacing:.12em;color:var(--rp-acc);font-weight:700}
.rp h1{font-size:20pt;line-height:1.3;margin:2px 0 8px;color:var(--rp-ink)}
.rp-cust{font-size:13pt;font-weight:700}
.rp-meta{font-size:9pt;color:var(--rp-mut);margin-top:2px}
.rp h2{font-size:11.5pt;margin:20px 0 8px;padding-left:8px;border-left:4px solid var(--rp-acc);color:var(--rp-ink)}
.rp-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}
.rp-stat{background:var(--rp-soft);border-radius:6px;padding:10px 10px 8px}
.rp-stat b{display:block;font-size:15pt;line-height:1.25;color:var(--rp-acc);font-variant-numeric:tabular-nums}
.rp-stat span{font-size:8.5pt;color:var(--rp-mut)}
.rp-basis{font-size:9pt;color:var(--rp-mut);margin:8px 0 0}
.rp-biz{border:1px solid var(--rp-line);border-left:4px solid var(--rp-acc);border-radius:4px;padding:8px 10px;font-size:9.5pt;margin-top:12px}
.rp-biz b{display:block;font-size:8.5pt;color:var(--rp-mut)}
.rp-top{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.rp-card{border:1px solid var(--rp-line);border-radius:6px;padding:9px 10px;break-inside:avoid}
.rp-card em{font-style:normal;font-size:8.5pt;font-weight:700;color:var(--rp-acc)}
.rp-card b{display:block;font-size:10pt;line-height:1.4;margin:1px 0 4px}
.rp-card .n{font-size:9pt;color:var(--rp-mut)}
.rp-card p{font-size:8.5pt;margin:4px 0 0;color:var(--rp-ink)}
.rp-tw{overflow-x:auto}
.rp table{width:100%;border-collapse:collapse;font-size:9pt}
.rp th{background:var(--rp-acc);color:#fff;font-weight:700;padding:5px 6px;text-align:left;white-space:nowrap}
.rp td{padding:5px 6px;border-bottom:1px solid var(--rp-line);vertical-align:top}
.rp tbody tr:nth-child(even) td{background:#f7fafa}
.rp tr{break-inside:avoid}
.rp .r{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.rp td small{display:block;color:var(--rp-mut);font-size:8pt;line-height:1.45;margin-top:1px}
.rp tfoot td{font-weight:700;border-top:2px solid var(--rp-acc);border-bottom:0;background:#fff}
.rp-notes{font-size:8.8pt;color:var(--rp-ink);padding-left:1.2em;margin:0}
.rp-notes li{margin:2px 0}
.rp-foot{margin-top:18px;padding-top:8px;border-top:1px solid var(--rp-line);font-size:8.5pt;color:var(--rp-mut);display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap}
@media (max-width:640px){.rp{padding:18px 14px}.rp-stats{grid-template-columns:repeat(2,1fr)}.rp-top{grid-template-columns:1fr}.rp h1{font-size:16pt}}
@media print{
  @page{size:A4;margin:13mm}
  .rp{box-shadow:none;padding:0;max-width:none;font-size:9.5pt}
  .rp-tw{overflow:visible}
  .rp-stats{grid-template-columns:repeat(4,1fr)!important}.rp-top{grid-template-columns:repeat(3,1fr)!important}
  .rp *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .rp h2{break-after:avoid}
}`;

function injectCss(){
  if(document.getElementById('rp-css'))return;
  const s=document.createElement('style');s.id='rp-css';s.textContent=CSS;document.head.appendChild(s);
}

function notes(d){
  const base=[
    `この資料の数字は、生成AI（ChatGPT・Gemini・Copilotなど）を使い、人が最終確認することを前提にした目安です。効果を保証するものではありません。`,
    `月の削減時間 ＝ 月の回数 × 1回の所要時間 × 削減率。人件費換算は、1時間あたり${yen(d.wage)}で計算しています。`,
    d.kind==='h'
      ?'業務の洗い出しと、回数・時間・削減率は、AIがHPの内容などから推定した値です。実際の業務とは異なる場合があります。'
      :'回数と所要時間は、お客様へのヒアリングにもとづく値です。削減率は目安として置いた値です。',
    '実際の効果は、会社の運用やデータの状態によって変わります。個人情報や機密情報は、社内ルールに従って扱ってください。'
  ];
  return base;
}

function html(d){
  injectCss();
  const rows=d.rows;
  const top=rows.filter(r=>r.save>0).slice(0,3);
  const cust=d.customer?`<div class="rp-cust">${esc(d.customer)} 御中</div>`:'';
  const meta=`作成日：${dateJa(d.date)}${d.author?`　／　作成：${esc(d.author)}`:''}`;
  const trs=rows.map((r,i)=>`<tr><td class="r">${i+1}</td><td>${esc(r.name)}${r.ai?`<small>${esc(r.ai)}</small>`:''}</td>
    <td class="r">${r1(r.cnt)}</td><td class="r">${r1(r.min)}</td><td class="r">${r1(r.now)}</td><td class="r">${Math.round(r.rate)}%</td>
    <td class="r"><b>${r1(r.save)}</b></td><td class="r">${yen(r.save*12*d.wage)}</td></tr>`).join('');
  return `<div class="rp">
  <div class="rp-head"><div class="rp-kicker">参考資料　${esc(KIND[d.kind]||'')}</div><h1>${TITLE}</h1>${cust}<div class="rp-meta">${meta}</div></div>
  <h2>試算の結果（目安）</h2>
  <div class="rp-stats">
    <div class="rp-stat"><b>${r1(d.save)}時間</b><span>月間の削減時間</span></div>
    <div class="rp-stat"><b>${r1(d.year)}時間</b><span>年間の削減時間</span></div>
    <div class="rp-stat"><b>${yen(d.yen)}</b><span>人件費に換算（年間）</span></div>
    <div class="rp-stat"><b>${pct(d.pct)}</b><span>削減できる割合</span></div>
  </div>
  <p class="rp-basis">試算の対象：${rows.length}業務（現在かかっている時間は月${r1(d.now)}時間）／ 1時間あたりの人件費 ${yen(d.wage)}</p>
  ${d.summary?`<div class="rp-biz"><b>事業内容（AIによる読み取り）</b>${esc(d.summary)}</div>`:''}
  ${top.length?`<h2>まず着手したい業務 上位${top.length}つ</h2><div class="rp-top">${top.map((r,i)=>`<div class="rp-card"><em>候補 ${i+1}</em><b>${esc(r.name)}</b><div class="n">月 ${r1(r.save)}時間の削減 ／ 年間 ${yen(r.save*12*d.wage)}</div>${r.ai?`<p>${esc(r.ai)}</p>`:''}</div>`).join('')}</div>`:''}
  <h2>業務別の試算（全${rows.length}件・削減時間の大きい順）</h2>
  <div class="rp-tw"><table><thead><tr><th class="r">順位</th><th>業務 ／ AIの使い方の例</th><th class="r">月の回数</th><th class="r">1回(分)</th><th class="r">現在の月間(時間)</th><th class="r">削減率</th><th class="r">月の削減(時間)</th><th class="r">年間の人件費換算</th></tr></thead>
  <tbody>${trs||'<tr><td colspan="8">業務が入力されていません</td></tr>'}</tbody>
  <tfoot><tr><td colspan="4">合計</td><td class="r">${r1(d.now)}</td><td class="r">${pct(d.pct)}</td><td class="r">${r1(d.save)}</td><td class="r">${yen(d.yen)}</td></tr></tfoot></table></div>
  <h2>前提と注意</h2>
  <ul class="rp-notes">${notes(d).map(n=>`<li>${esc(n)}</li>`).join('')}</ul>
  <div class="rp-foot"><span>${esc(TITLE)}（参考資料）</span><span>${esc(meta)}</span></div>
</div>`;
}

/* ---------- PowerPoint ---------- */
const PPTX_URL='https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
function loadPptx(){
  if(window.PptxGenJS)return Promise.resolve();
  return new Promise((res,rej)=>{
    const s=document.createElement('script');s.src=PPTX_URL;
    s.onload=()=>window.PptxGenJS?res():rej(new Error('x'));
    s.onerror=()=>rej(new Error('x'));
    document.head.appendChild(s);
  });
}
const safeName=s=>String(s||'').replace(/[\\/:*?"<>|\s]+/g,'_').slice(0,40);
function fileBase(d){return `AI業務診断_${safeName(d.customer)||'お客様'}_${ymd(d.date)}`}

async function pptx(d){
  try{await loadPptx()}catch(e){throw new Error('PowerPointの部品を読み込めませんでした。ネットワークを確認してください。')}
  const P=new window.PptxGenJS();
  P.layout='LAYOUT_16x9';
  P.title=TITLE;
  const F='Yu Gothic', INK='1B2A2F', ACC='0D6B63', MUT='5A6B72', SOFT='EEF6F4', LINE='CFD8DB';
  const meta=`作成日：${dateJa(d.date)}${d.author?`　／　作成：${d.author}`:''}`;
  const frame=(s,title,n)=>{
    s.background={color:'FFFFFF'};
    s.addShape('rect',{x:0,y:0,w:10,h:0.08,fill:{color:ACC},line:{color:ACC}});
    s.addText(title,{x:0.5,y:0.25,w:9,h:0.5,fontFace:F,fontSize:22,bold:true,color:INK});
    s.addText(`${TITLE}（参考資料）　${n}`,{x:0.5,y:5.25,w:9,h:0.25,fontFace:F,fontSize:9,color:MUT});
  };
  // 表紙
  let s=P.addSlide();
  s.background={color:'FFFFFF'};
  s.addShape('rect',{x:0,y:0,w:0.35,h:5.625,fill:{color:ACC},line:{color:ACC}});
  s.addText(`参考資料　${KIND[d.kind]||''}`,{x:0.9,y:1.2,w:8.5,h:0.35,fontFace:F,fontSize:13,bold:true,color:ACC});
  s.addText(TITLE,{x:0.9,y:1.6,w:8.5,h:1.1,fontFace:F,fontSize:34,bold:true,color:INK});
  if(d.customer)s.addText(`${d.customer} 御中`,{x:0.9,y:2.9,w:8.5,h:0.5,fontFace:F,fontSize:20,bold:true,color:INK});
  s.addText(meta,{x:0.9,y:3.6,w:8.5,h:0.4,fontFace:F,fontSize:13,color:MUT});
  // 結果
  s=P.addSlide();frame(s,'試算の結果（目安）','');
  const stats=[[`${r1(d.save)}時間`,'月間の削減時間'],[`${r1(d.year)}時間`,'年間の削減時間'],[yen(d.yen),'人件費に換算（年間）'],[pct(d.pct),'削減できる割合']];
  stats.forEach((st,i)=>{
    const x=0.5+i*2.3;
    s.addShape('roundRect',{x,y:1.0,w:2.15,h:1.5,fill:{color:SOFT},line:{color:SOFT},rectRadius:0.08});
    s.addText(st[0],{x,y:1.15,w:2.15,h:0.7,fontFace:F,fontSize:i===2?20:26,bold:true,color:ACC,align:'center',valign:'middle'});
    s.addText(st[1],{x,y:1.9,w:2.15,h:0.4,fontFace:F,fontSize:12,color:MUT,align:'center'});
  });
  s.addText(`試算の対象：${d.rows.length}業務（現在かかっている時間は月${r1(d.now)}時間）／ 1時間あたりの人件費 ${yen(d.wage)}`,{x:0.5,y:2.75,w:9,h:0.4,fontFace:F,fontSize:12,color:MUT});
  if(d.summary)s.addText([{text:'事業内容（AIによる読み取り）\n',options:{bold:true,color:MUT,fontSize:11}},{text:d.summary,options:{color:INK,fontSize:13}}],{x:0.5,y:3.3,w:9,h:1.5,fontFace:F,valign:'top',fill:{color:'FFFFFF'},line:{color:LINE,width:1},margin:10});
  // 上位3
  const top=d.rows.filter(r=>r.save>0).slice(0,3);
  if(top.length){
    s=P.addSlide();frame(s,`まず着手したい業務 上位${top.length}つ`,'');
    top.forEach((r,i)=>{
      const x=0.5+i*3.05;
      s.addShape('roundRect',{x,y:1.0,w:2.9,h:3.9,fill:{color:'FFFFFF'},line:{color:LINE,width:1},rectRadius:0.08});
      s.addText(`候補 ${i+1}`,{x:x+0.15,y:1.1,w:2.6,h:0.3,fontFace:F,fontSize:12,bold:true,color:ACC});
      s.addText(r.name,{x:x+0.15,y:1.4,w:2.6,h:0.9,fontFace:F,fontSize:16,bold:true,color:INK,valign:'top'});
      s.addText(`月 ${r1(r.save)}時間の削減\n年間 ${yen(r.save*12*d.wage)}`,{x:x+0.15,y:2.35,w:2.6,h:0.7,fontFace:F,fontSize:13,color:ACC,bold:true,valign:'top'});
      if(r.ai)s.addText(r.ai,{x:x+0.15,y:3.1,w:2.6,h:1.7,fontFace:F,fontSize:12,color:INK,valign:'top'});
    });
  }
  // 業務別の表（1枚6行）
  const per=6, pages=Math.max(1,Math.ceil(d.rows.length/per));
  const hd=t=>({text:t,options:{bold:true,color:'FFFFFF',fill:{color:ACC},fontFace:F,fontSize:10,align:'center',valign:'middle'}});
  for(let p=0;p<pages;p++){
    s=P.addSlide();frame(s,`業務別の試算（削減時間の大きい順）${pages>1?` ${p+1}/${pages}`:''}`,'');
    const chunk=d.rows.slice(p*per,(p+1)*per);
    const body=chunk.map((r,i)=>[
      {text:String(p*per+i+1),options:{align:'center',fontFace:F,fontSize:10}},
      {text:r.name+(r.ai?`\n${r.ai}`:''),options:{fontFace:F,fontSize:10,valign:'top'}},
      {text:r1(r.cnt),options:{align:'right',fontFace:F,fontSize:10}},
      {text:r1(r.min),options:{align:'right',fontFace:F,fontSize:10}},
      {text:`${Math.round(r.rate)}%`,options:{align:'right',fontFace:F,fontSize:10}},
      {text:r1(r.save),options:{align:'right',bold:true,fontFace:F,fontSize:10,color:ACC}},
      {text:yen(r.save*12*d.wage),options:{align:'right',fontFace:F,fontSize:10}}
    ]);
    s.addTable([[hd('順位'),hd('業務 ／ AIの使い方の例'),hd('月の回数'),hd('1回(分)'),hd('削減率'),hd('月の削減(時間)'),hd('年間の人件費換算')],...body.length?body:[[{text:'',options:{}},{text:'業務が入力されていません',options:{fontFace:F,fontSize:10}},'','','','','']]],
      {x:0.5,y:0.95,w:9,colW:[0.55,3.85,0.8,0.7,0.7,1.0,1.4],border:{type:'solid',pt:0.5,color:LINE},fontFace:F,color:INK,margin:[3,5,3,5]});
    if(p===pages-1)s.addText(`合計　現在 月${r1(d.now)}時間 → 削減 月${r1(d.save)}時間（${pct(d.pct)}）／ 年間 ${yen(d.yen)}`,{x:0.5,y:4.85,w:9,h:0.3,fontFace:F,fontSize:11,bold:true,color:INK});
  }
  // 前提
  s=P.addSlide();frame(s,'前提と注意','');
  s.addText(notes(d).map(t=>({text:t,options:{bullet:true,breakLine:true}})),{x:0.5,y:1.0,w:9,h:3.8,fontFace:F,fontSize:13,color:INK,valign:'top',paraSpaceAfter:8});
  await P.writeFile({fileName:fileBase(d)+'.pptx'});
}

window.AIReport={html,pptx,fileBase};
})();
