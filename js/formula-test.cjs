
const fs=require("fs"),path=require("path"),vm=require("vm");
const ctx={window:{}};vm.createContext(ctx);
vm.runInContext(fs.readFileSync(process.argv[2],"utf8"),ctx);
const dir=process.argv[3],index=JSON.parse(fs.readFileSync(path.join(dir,"workbook-index.json"),"utf8"));
const data={};
for(const s of index.sheets){data[s.name]=JSON.parse(fs.readFileSync(path.join(dir,s.file),"utf8"));}
const Book=ctx.window.MasslehreFormulaBook,book=new Book(data);
let total=0,compared=0,matches=0,errors=0;const failures=[];
for(const s of index.sheets){
 const sh=data[s.name];
 for(const [addr,c] of Object.entries(sh.cells)){
  if(!c.formula)continue;
  total++;
  if(c.cached===null||c.cached===undefined)continue;
  compared++;
  const v=book.value(s.name,addr);
  const val=v&&v.__excelError?v.code:v;
  const norm=x=>typeof x==="number"?Math.round(x*100000)/100000:String(x??"");
  if(norm(val)===norm(c.cached))matches++;
  else {errors++;if(failures.length<80)failures.push({sheet:s.name,cell:addr,formula:c.value,expected:c.cached,actual:val});}
 }
}
console.log(JSON.stringify({total,compared,matches,errors,matchRate:compared?Math.round(matches/compared*10000)/100:0,failures},null,2));
