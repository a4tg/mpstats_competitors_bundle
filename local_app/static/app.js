let token='', active=null, lastPhase='', loadedRun=null;
const $=s=>document.querySelector(s);
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const number=v=>typeof v==='number'?new Intl.NumberFormat('ru-RU',{maximumFractionDigits:0}).format(v):'—';
async function api(url,options={}){const r=await fetch(url,options);const d=await r.json();if(!r.ok)throw Error(d.error||'Ошибка запроса');return d;}
async function display(run){
 active=run.id;
 if(run.phase==='done'){
  if(loadedRun===run.id&&lastPhase==='done')return;
  const data=await api(`/runs/${run.id}/data`);loadedRun=run.id;
  $('#result').innerHTML=`<div class="stats"><div class="stat"><strong>${data.summary.selected}</strong><span>Выбрано товаров</span></div><div class="stat"><strong>${data.summary.matched}</strong><span>Есть сводные метрики</span></div><div class="stat"><strong>${data.summary.needs_review}</strong><span>Проверить связку</span></div></div><div class="actions"><a class="download" href="/runs/${run.id}/download">Скачать пакет для ИИ ↓</a>${run.has_queries?'<button class="secondary" id="collect">Собрать конкурентов в MPSTATS</button>':''}</div><div class="warnings"><b>Перед анализом</b><ul>${data.warnings.map(w=>`<li>${esc(w)}</li>`).join('')}</ul></div><div class="table-wrap"><table><thead><tr><th>SKU / АРТИКУЛ</th><th>ТОВАР</th><th>ВЫРУЧКА, ₽</th><th>ПРИБЫЛЬ, ₽</th><th>ПРОВЕРКА</th></tr></thead><tbody>${data.products.map(p=>`<tr><td>${esc(p.sku||'Не определён')}<br><small>${esc(p.article)}</small></td><td class="title" title="${esc(p.name)}">${esc(p.name||'—')}</td><td>${number(p.metrics?.revenue)}</td><td>${number(p.metrics?.profit)}</td><td title="${esc(p.issues.join('; '))}">${p.issues.length?esc(p.issues.join('; ')):'Связка найдена'}</td></tr>`).join('')}</tbody></table></div>`;
  if($('#collect'))$('#collect').onclick=async()=>{try{$('#collect').disabled=true;await api(`/collect/${run.id}/start`,{method:'POST',headers:{'X-App-Token':token}});loadedRun=null;await refresh();}catch(e){$('#form-error').textContent=e.message;if($('#collect'))$('#collect').disabled=false;}};
 }else if(run.phase==='error'){
  $('#result').innerHTML=`<div class="error"><h3>Нужно проверить данные</h3><p>${esc(run.message)}</p><a href="/runs/${run.id}/log" target="_blank">Журнал сборщика</a>${run.has_queries?'<p><button class="secondary" id="retry">Повторить сбор MPSTATS</button></p>':''}</div>`;
  if($('#retry'))$('#retry').onclick=async()=>{try{await api(`/collect/${run.id}/start`,{method:'POST',headers:{'X-App-Token':token}});await refresh();}catch(e){$('#form-error').textContent=e.message;}};
 }else $('#result').innerHTML=`<div class="status"><span class="spinner"></span>${esc(run.message)}${run.phase==='collecting'?`<p><a href="/runs/${run.id}/log" target="_blank">Открыть журнал прогресса</a></p><small>Браузер MPSTATS открывается отдельно. Оставьте приложение запущенным до завершения.</small>`:''}</div>`;
 lastPhase=run.phase;
}
let refreshing=false;
async function refresh(){if(refreshing)return;refreshing=true;try{const runs=await api('/api/runs');$('#history-count').textContent=runs.length?`${runs.length} запусков`:'';$('#history').innerHTML=runs.length?runs.map(r=>`<div class="history-row"><span>${esc(r.title||r.id.slice(0,8))}<br><small>${esc(r.message)}</small></span><button data-run="${r.id}">Открыть</button></div>`).join(''):'<small>Предыдущих запусков пока нет.</small>';document.querySelectorAll('[data-run]').forEach(b=>b.onclick=()=>{loadedRun=null;display(runs.find(r=>r.id===b.dataset.run)).catch(e=>$('#form-error').textContent=e.message);});const run=runs.find(r=>r.id===active)||runs[0];if(run)await display(run);}finally{refreshing=false;}}
$('#upload').onsubmit=async e=>{e.preventDefault();$('#form-error').textContent='';$('#submit').disabled=true;try{const d=await api('/api/import',{method:'POST',headers:{'X-App-Token':token},body:new FormData(e.target)});active=d.id;loadedRun=null;await refresh();}catch(e){$('#form-error').textContent=e.message;}finally{$('#submit').disabled=false;}};
api('/api/session').then(d=>{token=d.token;return refresh();}).catch(e=>$('#form-error').textContent=e.message);
setInterval(()=>refresh().catch(e=>$('#form-error').textContent='Нет связи с приложением. Проверьте, что окно запуска открыто.'),2500);
