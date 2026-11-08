
(function(){
"use strict";
const $=s=>document.querySelector(s);
const state={config:null,workbook:{},current:null,values:{},menuOpen:false};
const colNum=s=>{let n=0;for(const c of s.match(/[A-Z]+/)[0])n=n*26+c.charCodeAt(0)-64;return n;};
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const labelFromCell=(sheet,addr)=>{
  const m=addr.match(/^([A-Z]+)(\d+)$/);if(!m)return "Bezeichnung in Excel nicht hinterlegt";
  const row=Number(m[2]),col=colNum(m[1]),cells=sheet.cells;
  const find=(r,startCol)=>{
    for(let c=startCol;c>=1;c--){
      const v=cells[letter(c)+r];
      if(!v||v.formula||typeof v.value!=="string"||!v.value.trim())continue;
      const txt=v.value.trim();
      if(txt==="B"||txt==="H")continue;
      return txt.replace(/\s+/g," ");
    }
    return "";
  };
  const same=find(row,col-1);
  if(same)return same;
  for(let c=col-1;c>=1;c--){
    const v=cells[letter(c)+row];
    if(v&&!v.formula&&typeof v.value==="string"&&v.value.trim()==="B")return "Breite";
    if(v&&!v.formula&&typeof v.value==="string"&&v.value.trim()==="H")return "Höhe";
  }
  const previous=find(row-1,col-1);
  if(previous)return previous;
  const above=cells[addr.replace(/\d+$/,String(row-1))];
  if(above&&!above.formula&&typeof above.value==="string"&&above.value.trim())return above.value.trim();
  return "Bezeichnung in Excel nicht hinterlegt";
};
function letter(n){let s="";while(n){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}
function parseOptions(validation,sheetName,addr){
  if(!validation||validation.type!=="list"||!validation.formula1)return null;
  let f=validation.formula1;
  if(f.startsWith('"')&&f.endsWith('"'))return f.slice(1,-1).split(",").map(x=>x.trim()).filter(Boolean);
  if(f.toUpperCase()==="INDIRECT(B7)" && sheetName==="Zylinder Rechner" && addr==="B8"){
    const brand=getValue(sheetName,"B7");
    const ref=brand==="Griffwerk"?["Zylinder","I",2,30]:["Zylinder","A",2,88];
    const helper=state.workbook[ref[0]];
    if(!helper)return null;
    const values=[];
    for(let r=ref[2];r<=ref[3];r++){const v=helper.cells[ref[1]+r]?.value;if(v!==null&&v!==undefined&&String(v).trim())values.push(String(v));}
    return values;
  }
  return null;
}
function inputType(input){
  const v=input.validation;
  if(v&&v.type==="list")return "select";
  if(v&&v.type==="whole")return "number";
  return "text";
}
function getValue(sheet,addr){
  const k=sheet+"!"+addr;
  if(Object.prototype.hasOwnProperty.call(state.values,k))return state.values[k];
  const c=state.workbook[sheet]?.cells?.[addr];
  return c&&!c.formula?c.value:null;
}
function setValue(sheet,addr,value){
  const k=sheet+"!"+addr;
  state.values[k]=value===""?null:(/^-?\d+(?:[.,]\d+)?$/.test(String(value))?Number(String(value).replace(",",".")):value);
  renderCurrent();
}
function buildRuntimeData(){
  const data={};
  for(const [name,sheet] of Object.entries(state.workbook)){
    data[name]={cells:{}};
    for(const [addr,c] of Object.entries(sheet.cells)){
      data[name].cells[addr]={value:c.value,formula:c.formula};
    }
  }
  for(const [key,value] of Object.entries(state.values)){
    const p=key.indexOf("!"),s=key.slice(0,p),a=key.slice(p+1);
    if(data[s]?.cells[a])data[s].cells[a]={value:value,formula:false};
    else if(data[s])data[s].cells[a]={value:value,formula:false};
  }
  return data;
}
function makeBook(){return new window.MasslehreFormulaBook(buildRuntimeData());}
function display(v){
  if(v&&v.__excelError)return v.code;
  if(v===null||v===undefined||v==="")return "–";
  if(typeof v==="number")return Number.isInteger(v)?String(v):String(Math.round(v*1000)/1000).replace(".",",");
  return String(v);
}
function inputMarkup(sheet,input){
  const addr=input.cell, label=input.label||labelFromCell(sheet,addr), val=getValue(sheet.name,addr), typ=inputType(input), opts=parseOptions(input.validation,sheet.name,addr);
  let control="";
  if(typ==="select"&&opts){
    control=`<select name="${esc(addr)}" aria-label="${esc(label)}" data-cell="${esc(addr)}">${opts.map(o=>`<option ${String(o)===String(val)?"selected":""} value="${esc(o)}">${esc(o)}</option>`).join("")}</select>`;
  }else{
    control=`<input name="${esc(addr)}" aria-label="${esc(label)}" data-cell="${esc(addr)}" type="${typ==="number"?"number":"text"}" step="any" value="${esc(val??"")}" placeholder="Wert eingeben">`;
  }
  return `<label class="field"><span class="field-caption"><span class="field-label">${esc(label)}</span></span>${control}</label>`;
}
function measureAxis(sheet,addr){
  const match=addr.match(/^([A-Z]+)(\d+)$/);if(!match)return "";
  const row=Number(match[2]),col=colNum(match[1]);
  for(let c=col-1;c>=1;c--){
    const cell=sheet.cells[letter(c)+row];
    if(!cell||cell.formula||cell.value===null||cell.value===undefined)continue;
    const marker=String(cell.value).trim().toUpperCase();
    if(marker==="B")return "Breite";
    if(marker==="H")return "Höhe";
    if(marker)return "";
  }
  return "";
}
function pairedOutputGroups(sheet,formulas){
  const groups=[],claimed=new Set();
  for(const range of sheet.mergedRanges||[]){
    const match=range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);if(!match)continue;
    const startCol=colNum(match[1]),startRow=Number(match[2]),endCol=colNum(match[3]),endRow=Number(match[4]);
    if(endRow<=startRow)continue;
    const labelCell=sheet.cells[match[1]+startRow];
    const label=labelCell&&!labelCell.formula&&typeof labelCell.value==="string"?labelCell.value.trim():"";
    if(!label)continue;
    const items=formulas.filter(([addr])=>{
      const cell=addr.match(/^([A-Z]+)(\d+)$/);if(!cell)return false;
      const row=Number(cell[2]),col=colNum(cell[1]);
      return row>=startRow&&row<=endRow&&col>endCol&&Boolean(measureAxis(sheet,addr));
    });
    const axes=items.map(([addr])=>measureAxis(sheet,addr));
    if(!axes.includes("Breite")||!axes.includes("Höhe"))continue;
    groups.push({label,items,axisLabels:true});
    for(const [addr] of items)claimed.add(addr);
  }
  return {groups,claimed};
}
function outputRows(sheet,book){
  const formulas=Object.entries(sheet.cells).filter(([,cell])=>cell.formula).sort(([a],[b])=>{
    const ma=a.match(/^([A-Z]+)(\d+)$/),mb=b.match(/^([A-Z]+)(\d+)$/);
    return Number(ma?.[2]||0)-Number(mb?.[2]||0)||colNum(ma?.[1]||"A")-colNum(mb?.[1]||"A");
  });
  const paired=pairedOutputGroups(sheet,formulas);
  const rows=[...paired.groups.map(group=>({row:Number(group.items[0][0].match(/\d+$/)[0]),label:group.label,items:group.items,axisLabels:true}))];
  for(const [addr] of formulas){
    if(paired.claimed.has(addr))continue;
    rows.push({row:Number(addr.match(/\d+$/)[0]),label:labelFromCell(sheet,addr),items:[[addr,sheet.cells[addr]]],axisLabels:false});
  }
  rows.sort((a,b)=>a.row-b.row);
  return rows.map(row=>`<div class="result-row"><div class="result-label">${esc(row.label||"Ergebnis")}</div><div class="result-values">${row.items.map(([addr])=>`<div class="result-value">${row.axisLabels?`<span class="result-axis">${esc(measureAxis(sheet,addr))}</span>`:""}<strong>${esc(display(book.value(sheet.name,addr)))}</strong></div>`).join("")}</div></div>`).join("")||`<p class="muted">Für dieses Blatt sind keine Formeln hinterlegt.</p>`;
}
function staticNotes(sheet){
  const items=[];
  for(const [addr,c] of Object.entries(sheet.cells)){
    if(c.formula||c.value===null||c.value===undefined||typeof c.value!=="string")continue;
    if(c.value.trim()==="")continue;
    const col=colNum(addr),row=Number(addr.match(/\d+$/)[0]);
    if(col<=6 && row>2 && row<sheet.maxRow && !/^(Breite|Höhe|TB Breite|TB Höhe):?$/i.test(c.value.trim())){
      // Keep explanatory text and section labels, not every small row marker.
      if(c.value.length>32||/^[A-ZÄÖÜ0-9].*:$/.test(c.value.trim()))items.push(c.value.trim());
    }
  }
  const unique=[...new Set(items)].slice(0,16);
  return unique.length?`<details class="notes"><summary>Hinweise aus dem Excel-Blatt</summary><ul>${unique.map(x=>`<li>${esc(x).replace(/\n/g," ")}</li>`).join("")}</ul></details>`:"";
}
function renderNav(){
  const root=$("#nav-list");root.innerHTML="";
  const cats=state.config.categories||{"Rechner":state.config.navigationSheets};
  for(const [cat,names] of Object.entries(cats)){
    if(!names.length)continue;
    const section=document.createElement("section");section.className="nav-section";
    section.innerHTML=`<h3>${esc(cat)}</h3>`;
    for(const name of names){
      const b=document.createElement("button");b.className="nav-item";b.textContent=name;b.dataset.sheet=name;
      if(name===state.current)b.classList.add("active");
      b.addEventListener("click",()=>selectSheet(name,true));section.appendChild(b);
    }
    root.appendChild(section);
  }
}

function referenceTable(sheet,book){
  const maxCol=Math.min(sheet.maxColumn||20,45), rows=[];
  for(let r=1;r<=sheet.maxRow;r++){
    const cells=[];
    for(let c=1;c<=maxCol;c++){
      const addr=letter(c)+r,cell=sheet.cells[addr];
      if(!cell)continue;
      const v=cell.formula?book.value(sheet.name,addr):cell.value;
      if(v===null||v===undefined||v==="")continue;
      cells.push(`<td><div>${esc(display(v))}</div></td>`);
    }
    if(cells.length)rows.push(`<tr><th>${r}</th>${cells.join("")}</tr>`);
  }
  return `<div class="table-wrap"><table class="reference-table"><tbody>${rows.join("")}</tbody></table></div>`;
}

function renderCurrent(){
  const sheet=state.workbook[state.current];if(!sheet)return;
  const book=makeBook();
  $("#page-title").textContent=sheet.name;
  $("#page-subtitle").textContent="Tür- und Zargensystem · Datenmodell "+state.config.fileVersion;
  const fields=sheet.inputs||[];
  $("#inputs").innerHTML=fields.length?fields.map(x=>inputMarkup(sheet,x)).join(""):`<p class="muted">Dieses Tabellenblatt enthält keine erkannten Eingabefelder.</p>`;
  $("#results").innerHTML=(state.config.referenceSheets||[]).includes(sheet.name)?referenceTable(sheet,book):outputRows(sheet,book);
  $("#notes").innerHTML=staticNotes(sheet);
  $("#input-count").textContent=fields.length+" Eingaben";
  $("#formula-count").textContent=sheet.formulaCount||Object.values(sheet.cells).filter(c=>c.formula).length+" Formeln";
  $("#inputs").querySelectorAll("[data-cell]").forEach(el=>{
    el.addEventListener("change",()=>setValue(sheet.name,el.dataset.cell,el.value));
  });
  renderNav();
}
async function loadSheet(name){
  if(state.workbook[name])return state.workbook[name];
  const summary=state.config.sheets.find(s=>s.name===name);
  if(!summary)throw new Error("Blatt nicht gefunden: "+name);
  const data=await loadSheetData(summary);
  state.workbook[name]=data;return data;
}
async function loadSheetData(summary){
  const response=await fetch("data/"+summary.file);
  if(!response.ok)throw new Error("Datenblatt konnte nicht geladen werden: "+summary.name);
  if(summary.file.endsWith(".json.gz")){
    if(typeof DecompressionStream==="undefined")throw new Error("Dieser Browser unterstützt keine gzip-Dekomprimierung.");
    return await new Response(response.body.pipeThrough(new DecompressionStream("gzip"))).json();
  }
  return await response.json();
}
async function selectSheet(name,closeMenu){
  try{
    await loadSheet(name);state.current=name;
    $("#welcome").hidden=true;$("#calculator").hidden=false;
    if(closeMenu)closeNav();
    renderCurrent();
    window.scrollTo({top:0,behavior:"smooth"});
  }catch(e){$("#load-error").textContent=e.message;$("#load-error").hidden=false;}
}
function openNav(){state.menuOpen=true;$("#sidebar").classList.add("open");$("#scrim").hidden=false;$("#menu-toggle").setAttribute("aria-expanded","true");}
function closeNav(){state.menuOpen=false;$("#sidebar").classList.remove("open");$("#scrim").hidden=true;$("#menu-toggle").setAttribute("aria-expanded","false");}
async function init(){
  try{
    const r=await fetch("data/app-config.json");state.config=await r.json();
    const idx=await fetch("data/workbook-index.json");const index=await idx.json();state.config.sheets=index.sheets;
    $("#version").textContent="Excel-Datenstand "+state.config.fileVersion;
    // Load helper sheets as well as calculator sheets so cross-sheet Excel formulas can resolve.
    await Promise.all(index.sheets.map(async summary=>{
      state.workbook[summary.name]=await loadSheetData(summary);
    }));
    $("#welcome").hidden=false;renderNav();
    $("#menu-toggle").addEventListener("click",()=>state.menuOpen?closeNav():openNav());
    $("#scrim").addEventListener("click",closeNav);
    $("#search").addEventListener("input",e=>{
      const q=e.target.value.toLocaleLowerCase("de");
      document.querySelectorAll(".nav-item").forEach(b=>b.hidden=!b.textContent.toLocaleLowerCase("de").includes(q));
      document.querySelectorAll(".nav-section").forEach(s=>s.hidden=!s.querySelector(".nav-item:not([hidden])"));
    });
    $("#home").addEventListener("click",()=>{$("#calculator").hidden=true;$("#welcome").hidden=false;closeNav();});
    $("#reset").addEventListener("click",()=>{state.values={};renderCurrent();});
    $("#load-error").hidden=true;
    if(state.config.navigationSheets.includes("Standard"))await selectSheet("Standard",false);
  }catch(e){$("#load-error").textContent="Die Maßlehre+ konnte nicht gestartet werden: "+e.message;$("#load-error").hidden=false;}
}
document.addEventListener("DOMContentLoaded",init);
})();
