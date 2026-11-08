
(function(){
  "use strict";
  const ERR = (code, message) => ({__excelError:true, code, message:message||code});
  const isErr = v => !!(v && v.__excelError);
  const colToNum = s => { let n=0; for(const ch of s.toUpperCase()) n=n*26+ch.charCodeAt(0)-64; return n; };
  const numToCol = n => { let s=""; while(n){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s; };
  function tokenize(src){
    src=String(src).replace(/^=/,"");
    const t=[]; let i=0;
    while(i<src.length){
      if(/\s/.test(src[i])){i++;continue;}
      if(src[i]==='"'){
        i++; let s="";
        while(i<src.length){if(src[i]==='"'&&src[i+1]==='"'){s+='"';i+=2;continue;} if(src[i]==='"'){i++;break;}s+=src[i++];}
        t.push({type:"str",value:s});continue;
      }
      if(src[i]==="'"){
        i++;let sh="";
        while(i<src.length){if(src[i]==="'"&&src[i+1]==="'"){sh+="'";i+=2;continue;}if(src[i]==="'"){i++;break;}sh+=src[i++];}
        if(src[i]==="!")i++;
        const m=src.slice(i).match(/^\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/i);
        if(m){t.push({type:"ref",sheet:sh,range:m[0]});i+=m[0].length;continue;}
        t.push({type:"ident",value:sh});continue;
      }
      const rest=src.slice(i);
      const sheetRef=rest.match(/^([A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/i);
      if(sheetRef){
        const bang=sheetRef[0].indexOf("!");
        t.push({type:"ref",sheet:sheetRef[1],range:sheetRef[0].slice(bang+1)});
        i+=sheetRef[0].length;continue;
      }
      const ref=rest.match(/^\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/i);
      if(ref){t.push({type:"ref",sheet:null,range:ref[0]});i+=ref[0].length;continue;}
      const num=rest.match(/^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/);
      if(num){t.push({type:"num",value:Number(num[0])});i+=num[0].length;continue;}
      const id=rest.match(/^[_A-Za-z\\][A-Za-z0-9_.\\]*/);
      if(id){t.push({type:"ident",value:id[0].replace(/^_xlfn\./i,"")});i+=id[0].length;continue;}
      const op=rest.match(/^(<=|>=|<>|[+\-*/^&=<>(),:%])/);
      if(op){t.push({type:"op",value:op[0]});i+=op[0].length;continue;}
      throw new Error("Unbekanntes Formelzeichen: "+src.slice(i,i+12));
    }
    t.push({type:"eof"});return t;
  }
  function parse(src){
    const toks=tokenize(src);let p=0;
    const peek=()=>toks[p], take=()=>toks[p++];
    const prec={"=":1,"<>":1,"<":1,">":1,"<=":1,">=":1,"&":2,"+":3,"-":3,"*":4,"/":4,"^":5};
    function primary(){
      const x=take();
      if(x.type==="num"||x.type==="str")return {type:"lit",value:x.value};
      if(x.type==="ref")return {type:"ref",sheet:x.sheet,range:x.range};
      if(x.type==="op"&&(x.value==="+"||x.value==="-"))return {type:"unary",op:x.value,arg:expr(6)};
      if(x.type==="op"&&x.value==="("){const n=expr(0);if(peek().value!==")")throw new Error("Klammer fehlt");take();return n;}
      if(x.type==="ident"){
        const name=x.value.toUpperCase();
        if(peek().value==="("){
          take();const args=[];
          while(peek().value!==")"&&peek().type!=="eof"){
            if(peek().value===","){args.push({type:"lit",value:null});take();continue;}
            args.push(expr(0));
            if(peek().value===","){take();continue;}
            break;
          }
          if(peek().value!==")")throw new Error("Funktionsklammer fehlt bei "+name);take();
          return {type:"call",name,args};
        }
        if(name==="TRUE")return {type:"lit",value:true};
        if(name==="FALSE")return {type:"lit",value:false};
        return {type:"name",name:x.value};
      }
      throw new Error("Unerwarteter Formelteil "+JSON.stringify(x));
    }
    function expr(min){
      let left=primary();
      while(true){
        const op=peek().value, pr=prec[op];
        if(pr===undefined||pr<min)break;
        take();const right=expr(pr+(op==="^"?0:1));
        left={type:"bin",op,left,right};
      }
      return left;
    }
    const ast=expr(0);
    if(peek().type!=="eof")throw new Error("Nicht verarbeiteter Formelrest "+peek().value);
    return ast;
  }
  const splitAddr = a => {const m=String(a).replace(/\$/g,"").match(/^([A-Z]+)(\d+)$/i);return m?{col:colToNum(m[1]),row:Number(m[2])}:null;};
  const flatten = x => Array.isArray(x)?x.flat(Infinity):[x];
  const truth = v => {if(isErr(v))return v;if(v===null||v===undefined||v==="")return false;if(typeof v==="string")return v.toLowerCase()!=="false";return !!v;};
  const number = v => {if(isErr(v))return v;if(v===null||v===undefined||v==="")return 0;if(typeof v==="boolean")return v?1:0;if(typeof v==="number")return v;const n=Number(String(v).replace(",","."));return Number.isFinite(n)?n:ERR("#VALUE!");};
  function evalAst(n,ctx){
    if(n.type==="lit")return n.value;
    if(n.type==="name")return ERR("#NAME?","Unbekannter Name "+n.name);
    if(n.type==="ref"){
      const sheet=n.sheet||ctx.sheet;
      const parts=n.range.split(":");
      if(parts.length===1)return ctx.getCell(sheet,parts[0].replace(/\$/g,""));
      const a=splitAddr(parts[0]),b=splitAddr(parts[1]);
      if(!a||!b)return ERR("#REF!");
      const values=[];
      for(let r=Math.min(a.row,b.row);r<=Math.max(a.row,b.row);r++){
        const row=[];
        for(let c=Math.min(a.col,b.col);c<=Math.max(a.col,b.col);c++)row.push(ctx.getCell(sheet,numToCol(c)+r));
        values.push(row);
      }
      return values;
    }
    if(n.type==="unary"){const v=number(evalAst(n.arg,ctx));return isErr(v)?v:(n.op==="-"?-v:v);}
    if(n.type==="bin"){
      const a=evalAst(n.left,ctx),b=evalAst(n.right,ctx);
      if(isErr(a))return a;if(isErr(b))return b;
      switch(n.op){
        case "+":{const x=number(a),y=number(b);return isErr(x)?x:(isErr(y)?y:x+y);}
        case "-":{const x=number(a),y=number(b);return isErr(x)?x:(isErr(y)?y:x-y);}
        case "*":{const x=number(a),y=number(b);return isErr(x)?x:(isErr(y)?y:x*y);}
        case "/":{const x=number(a),y=number(b);return isErr(x)?x:(isErr(y)?y:(y===0?ERR("#DIV/0!"):x/y));}
        case "^":{const x=number(a),y=number(b);return isErr(x)?x:(isErr(y)?y:Math.pow(x,y));}
        case "&":return (a==null?"":String(a))+(b==null?"":String(b));
        case "=":return String(a).toLowerCase()===String(b).toLowerCase() || (typeof a==="number"&&typeof b==="number"&&a===b);
        case "<>":return !(String(a).toLowerCase()===String(b).toLowerCase() || (typeof a==="number"&&typeof b==="number"&&a===b));
        case "<":return a<b;case ">":return a>b;case "<=":return a<=b;case ">=":return a>=b;
      }
    }
    if(n.type==="call"){
      const arg=(i)=>i<n.args.length?evalAst(n.args[i],ctx):null;
      const all=()=>n.args.map((_,i)=>arg(i));
      switch(n.name){
        case "IF": {const c=truth(arg(0));return isErr(c)?c:evalAst(n.args[c?1:2]||{type:"lit",value:false},ctx);}
        case "IFNA": {const v=arg(0);return isErr(v)&&v.code==="#N/A"?arg(1):v;}
        case "SUM": {let s=0;for(const v of flatten(all())){if(isErr(v))return v;if(typeof v==="number")s+=v;else if(typeof v==="boolean")s+=v?1:0;}return s;}
        case "ISBLANK": {const v=arg(0);return v===null||v===undefined;}
        case "ISNUMBER": {const v=arg(0);return typeof v==="number"&&Number.isFinite(v);}
        case "ISNA": {const v=arg(0);return isErr(v)&&v.code==="#N/A";}
        case "OR": {let result=false;for(const v of flatten(all())){if(isErr(v))return v;result=result||truth(v);}return result;}
        case "AND": {let result=true;for(const v of flatten(all())){if(isErr(v))return v;result=result&&truth(v);}return result;}
        case "VLOOKUP": {
          const key=arg(0), table=arg(1), index=number(arg(2)), exact=arg(3)===false;
          if(isErr(key)||isErr(table)||isErr(index))return ERR("#N/A");
          if(!Array.isArray(table)||index<1)return ERR("#REF!");
          let found=-1;
          for(let r=0;r<table.length;r++){
            const v=table[r][0];
            if(String(v).toLowerCase()===String(key).toLowerCase() || (typeof v==="number"&&typeof key==="number"&&v===key)){found=r;break;}
            if(!exact && v<=key)found=r;
          }
          if(found<0)return ERR("#N/A");
          return table[found][index-1]===undefined?ERR("#REF!"):table[found][index-1];
        }
        default:return ERR("#NAME?","Nicht unterstützte Funktion "+n.name);
      }
    }
    return ERR("#VALUE!");
  }
  function FormulaBook(data){
    this.data=data;this.cache={};this.active=new Set();
    this.getCell=(sheet,addr)=>{
      const key=sheet+"!"+addr;
      if(Object.prototype.hasOwnProperty.call(this.cache,key))return this.cache[key];
      if(this.active.has(key))return ERR("#CIRC!","Zirkelbezug "+key);
      const sh=this.data[sheet];if(!sh)return ERR("#REF!","Blatt fehlt: "+sheet);
      const cell=sh.cells[addr];if(!cell)return null;
      if(!cell.formula)return cell.value;
      this.active.add(key);
      let result;
      try{result=evalAst(parse(cell.value),{sheet,getCell:this.getCell});}
      catch(e){result=ERR("#FORMULA!",e.message);}
      this.active.delete(key);this.cache[key]=result;return result;
    };
    this.value=(sheet,addr)=>this.getCell(sheet,addr);
    this.clear=()=>{this.cache={};this.active.clear();};
  }
  window.MasslehreFormulaBook=FormulaBook;
  window.MasslehreExcelError=ERR;
})();
