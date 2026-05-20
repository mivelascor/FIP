/**
 * generar_folleto.js — Folleto comercial Vantrust.
 * Layout fiel al HTML de referencia (FIP_Alto_Aporte__1_.html).
 * A4 portrait: col izq 68mm (negra), col der contenido.
 */
"use strict";
const pptxgen = require("pptxgenjs");
const fs = require("fs");

const C = { negro:"000000", blanco:"FFFFFF", grisOsc:"333333",
            grisMed:"666666", grisClr:"999999", grisLin:"DDDDDD" };
const W=8.27, H=11.69;
const IZQ=2.68, DER=2.85, DER_W=W-DER-0.15, P=0.13;

function pct(v, d=2) {
  if (v===null||v===undefined) return "—";
  if (typeof v==="string") return v;
  if (typeof v==="number"&&!isNaN(v)) return (v*100).toFixed(d)+"%";
  return "—";
}
function lin(sl,y,x=DER,w=DER_W,col=C.grisLin,pt=0.5){
  sl.addShape("line",{x,y,w,h:0,line:{color:col,width:pt}});
}

/* ══════════════════════════════════════════════════════════════════════
   SLIDE 1
══════════════════════════════════════════════════════════════════════ */
function slide1(pres, d) {
  const sl = pres.addSlide();
  sl.background = {color:C.blanco};

  /* ── COLUMNA IZQUIERDA ─────────────────────────────────────────── */
  sl.addShape("rect",{x:0,y:0,w:IZQ,h:1.9,fill:{color:C.negro}});

  // Título: "Liquidez\nActiva" — parte en la última palabra
  const words = (d.nombre_corto||"").split(" ");
  const l1 = words.length>1 ? words.slice(0,-1).join(" ") : words[0];
  const l2 = words.length>1 ? words[words.length-1] : "";
  sl.addText(l2?`${l1}\n${l2}`:l1, {
    x:P,y:0.18,w:IZQ-P*2,h:1.3,
    fontSize:22,bold:true,color:C.blanco,fontFace:"Calibri",
    valign:"top",wrap:true,margin:0
  });
  sl.addText("F O N D O",{x:P,y:1.56,w:IZQ-P*2,h:0.18,
    fontSize:7,color:"AAAAAA",fontFace:"Calibri",charSpacing:4,margin:0});
  sl.addShape("line",{x:IZQ,y:0,w:0,h:H,line:{color:"EEEEEE",width:0.5}});

  let y=2.0;
  const LW=IZQ-P*2;
  function secL(t){
    sl.addText(t,{x:P,y,w:LW,h:0.18,fontSize:8,bold:true,color:C.negro,fontFace:"Calibri",margin:0});
    y+=0.2;sl.addShape("line",{x:P,y,w:LW,h:0,line:{color:C.negro,width:1}});y+=0.07;
  }
  function divL(){
    sl.addShape("line",{x:P,y,w:LW,h:0,line:{color:C.grisLin,width:0.5}});y+=0.1;
  }
  function fi(k,v,h=0.17){
    sl.addText(k,{x:P,y,w:1.0,h,fontSize:6.5,bold:true,color:C.negro,fontFace:"Calibri",valign:"top",margin:0,wrap:true});
    sl.addText(v||"",{x:P+1.02,y,w:LW-1.02,h,fontSize:6.5,color:"555555",fontFace:"Calibri",valign:"top",margin:0,wrap:true});
    y+=h+0.02;
  }

  secL("Información General");
  fi("Administradora",d.administradora);
  fi("RUT Fondo",d.rut);
  fi("Moneda",d.moneda);
  fi("Tipo de Fondo",d.tipo);
  fi("Fecha Inicio",d.fecha_inicio);
  fi("Benchmark",d.benchmark);
  fi("Fondo Rescatable","Sí");
  fi("Plazo Rescate",d.plazo_rescate);
  fi("Riesgos","Mercado- Crédito - Liquidez - Tasa de interés - Derivados",0.27);
  fi("Remuneración",d.remuneracion||"0,295% IVA Incluido",0.27);
  fi("Custodio","Vantrust Capital C. de Bolsa");
  divL();
  secL("Objetivo");
  sl.addText(d.objetivo||"Invertir los recursos del fondo en instrumentos de deuda de corto y mediano plazo, en una cartera diversificada, obteniendo una rentabilidad igual o superior al ICP.",
    {x:P,y,w:LW,h:0.7,fontSize:7,color:"555555",fontFace:"Calibri",valign:"top",margin:0,wrap:true});
  y+=0.77;divL();
  secL("Rentabilidad");
  sl.addText(d.rentabilidad_texto||`La rentabilidad esperada del ${d.nombre_fondo}, es la tasa de política monetaria promedio del Banco Central de Chile.`,
    {x:P,y,w:LW,h:0.65,fontSize:7,color:"555555",fontFace:"Calibri",valign:"top",margin:0,wrap:true});
  y+=0.72;divL();
  secL("Inversionistas");
  sl.addText(d.inversionistas||"Dirigida a empresas y personas que buscan invertir sus excedentes de caja con una rentabilidad de corto plazo y baja tolerancia al riesgo.",
    {x:P,y,w:LW,h:0.55,fontSize:7,color:"555555",fontFace:"Calibri",valign:"top",margin:0,wrap:true});

  /* ── COLUMNA DERECHA ─────────────────────────────────────────── */
  let yr=0.15;

  // Comentario PM
  sl.addText("Comentario Portafolio Manager",{x:DER,y:yr,w:DER_W,h:0.27,
    fontSize:13,bold:true,color:C.negro,fontFace:"Calibri",margin:0});
  yr+=0.3;lin(sl,yr,DER,DER_W,C.negro,1.5);yr+=0.1;
  sl.addText(d.comentario||"",{x:DER,y:yr,w:DER_W,h:1.8,
    fontSize:8,color:C.grisOsc,fontFace:"Calibri",valign:"top",margin:0,wrap:true,align:"justify"});
  yr+=1.9;

  // Gráfico
  const graf = d.grafico||{};
  const gLabels = graf.labels||[];
  if (gLabels.length>0) {
    sl.addText(d.nombre_corto||"",{x:DER,y:yr,w:DER_W,h:0.16,
      fontSize:7.5,bold:true,color:C.negro,fontFace:"Calibri",align:"center",margin:0});
    yr+=0.17;
    const series=[];
    const gICP  = graf.icp||[], gFIP = graf.fip||[], gComp=graf.comp||[];
    if(gICP.some(v=>v!==null&&v!==undefined))  series.push({name:"ICP Nom.",labels:gLabels,values:gICP});
    if(gFIP.some(v=>v!==null&&v!==undefined))  series.push({name:d.nombre_fip||"FIP",labels:gLabels,values:gFIP});
    if(gComp.some(v=>v!==null&&v!==undefined)) series.push({name:"Competencia Relevante (*)",labels:gLabels,values:gComp});
    if(series.length){
      const allV=series.flatMap(s=>s.values).filter(v=>v!==null&&!isNaN(v));
      sl.addChart("line",series,{
        x:DER,y:yr,w:DER_W,h:1.85,
        chartColors:["999999","000000","555555"],
        lineSize:1.2,lineSmooth:false,
        showLegend:true,legendPos:"r",legendFontSize:6,
        catAxisLabelFontSize:6,valAxisLabelFontSize:6,
        catAxisLabelColor:C.grisMed,valAxisLabelColor:C.grisMed,
        catAxisLabelFrequency:Math.max(1,Math.floor(gLabels.length/10)),
        valGridLine:{color:"EEEEEE",size:0.3},catGridLine:{style:"none"},
        chartArea:{fill:{color:C.blanco}},plotArea:{fill:{color:C.blanco}},showTitle:false,
        valAxisMinVal:allV.length?Math.floor(Math.min(...allV)*0.97):90,
      });
    }
    yr+=2.02;
  }

  // Evolución Rentabilidad — tabla resumen
  sl.addText("Evolución Rentabilidad",{x:DER,y:yr,w:DER_W,h:0.25,
    fontSize:13,bold:true,color:C.negro,fontFace:"Calibri",margin:0});
  yr+=0.28;lin(sl,yr,DER,DER_W,C.negro,1.5);yr+=0.08;

  const colW=[1.42,0.62,0.72,0.72,0.62,0.65];
  const bTop=[{pt:1.5,color:C.negro},{pt:0.3,color:C.grisLin},{pt:1.5,color:C.negro},{pt:0.3,color:C.grisLin}];
  const bRow=[{pt:0.3,color:C.grisLin},{pt:0.3,color:C.grisLin},{pt:0.3,color:C.grisLin},{pt:0.3,color:C.grisLin}];
  const mkH=(t,l=false)=>({text:t,options:{bold:true,fontSize:7.5,fontFace:"Calibri",
    align:l?"left":"center",color:C.negro,fill:{color:C.blanco},border:bTop}});
  const mkD=(t,l=false,bg="FFFFFF")=>({text:String(t||"—"),options:{fontSize:7.5,fontFace:"Calibri",
    align:l?"left":"center",color:l?C.negro:C.grisOsc,bold:l,fill:{color:bg},border:bRow}});

  const acumLbl = (d.acum_label||"Acum. 2026 (*)").replace("Acum. ","Acum.\n").replace("Acum\n","Acum.\n");
  const resumen = d.resumen||[];
  const tblResumen = [
    [mkH("Rentabilidad",true),mkH("Mensual"),mkH("Trimestral"),mkH("Semestral"),mkH("Anual"),mkH(acumLbl)]
  ];
  resumen.forEach((r,i)=>{
    const bg = i%2===0?"F9F9F9":"FFFFFF";
    tblResumen.push([
      mkD(r.nombre,true,bg),
      mkD(pct(r.m),"",bg), mkD(pct(r.t),"",bg),
      mkD(pct(r.s),"",bg), mkD(pct(r.a),"",bg),
      mkD(pct(r.ac),"",bg),
    ]);
  });
  // Asegurar siempre 3 filas de datos (ICP, Comp, FIP)
  while(tblResumen.length < 4) {
    tblResumen.push([mkD("—",true),mkD("—"),mkD("—"),mkD("—"),mkD("—"),mkD("—")]);
  }
  sl.addTable(tblResumen,{x:DER,y:yr,w:DER_W,colW,rowH:0.2});
  yr+=tblResumen.length*0.2+0.05;

  sl.addText("*Valores correspondientes a la rentabilidad anualizada, no acumulada",
    {x:DER,y:yr,w:DER_W,h:0.14,fontSize:6,color:C.grisClr,fontFace:"Calibri",italic:true,margin:0});
  yr+=0.18;

  // Tabla histórica — todos los años desde el inicio del fondo
  const historico = d.historico||[];
  if(historico.length>0){
    lin(sl,yr,DER,DER_W,C.grisLin,0.5);yr+=0.07;

    const mL=["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic","Total\nAño"];
    // año(0.22) + fondo(0.88) + 12×mes(0.32) + total(0.33) = 5.27" = DER_W
    const cw=[0.22,0.88,0.32,0.32,0.32,0.32,0.32,0.32,0.32,0.32,0.32,0.32,0.32,0.32,0.33];
    const bH=[{pt:1.5,color:C.negro},{pt:0.2,color:"EEEEEE"},{pt:1.5,color:C.negro},{pt:0.2,color:"EEEEEE"}];
    const bD=[{pt:0.2,color:"EEEEEE"},{pt:0.2,color:"EEEEEE"},{pt:0.2,color:"EEEEEE"},{pt:0.2,color:"EEEEEE"}];
    const mkTH=(t,l=false)=>({text:t,options:{
      fontSize:5.5,bold:true,fontFace:"Calibri",color:C.negro,
      align:l?"left":"center",border:bH,wrap:false}});

    const trows=[[mkTH("Año",true),mkTH("Fondo",true),...mL.map(m=>mkTH(m))]];

    function colorSerie(nombre) {
      const n=(nombre||"").toLowerCase();
      if(n.includes("icp")||n.includes("benchmark")) return "777777";
      if(n.includes("compet")) return "999999";
      return C.negro;
    }

    historico.forEach(añoData=>{
      let isFirst=true;
      (añoData.filas||[]).forEach(fila=>{
        const col = colorSerie(fila.nombre);
        // Fuente 4.5pt para el nombre (permite hasta 25 chars en 0.88")
        const fontNombre = (fila.nombre||"").length > 18 ? 4 : 5;
        const meses = fila.meses||Array(12).fill(null);
        const total = fila.total;
        const opt=(extra={})=>({fontSize:5,fontFace:"Calibri",color:col,align:"center",border:bD,wrap:false,...extra});
        const row=[
          {text:isFirst?String(añoData.año):"",
           options:{fontSize:5,bold:true,fontFace:"Calibri",color:C.negro,align:"left",border:bD,wrap:false}},
          {text:fila.nombre||"",
           options:{fontSize:fontNombre,fontFace:"Calibri",color:col,align:"left",border:bD,wrap:false}},
          ...meses.map(v=>({
            text:v!==null&&v!==undefined?pct(v):"",
            options:opt()
          })),
          {text:total!==null&&total!==undefined?pct(total):"—",
           options:opt({bold:true})}
        ];
        trows.push(row);
        isFirst=false;
      });
    });

    // rowH dinámico para que quepa en el espacio restante
    const espacioDisp = H - yr - 0.05;
    const rowH = Math.min(0.135, Math.max(0.085, espacioDisp / trows.length));
    sl.addTable(trows,{x:DER,y:yr,w:DER_W,colW:cw,rowH});
  }
}

/* ══════════════════════════════════════════════════════════════════════
   SLIDE 2
══════════════════════════════════════════════════════════════════════ */
function slide2(pres, d) {
  const sl=pres.addSlide();
  sl.background={color:C.blanco};
  sl.addShape("rect",{x:0,y:0,w:IZQ,h:1.45,fill:{color:C.negro}});
  sl.addText("F O N D O",{x:P,y:1.12,w:IZQ-P*2,h:0.2,
    fontSize:7,color:"AAAAAA",fontFace:"Calibri",charSpacing:4,margin:0});
  sl.addShape("line",{x:IZQ,y:0,w:0,h:H,line:{color:"EEEEEE",width:0.5}});

  let y=1.6;
  const LW=IZQ-P*2;
  function secL(t){
    sl.addText(t,{x:P,y,w:LW,h:0.18,fontSize:8,bold:true,color:C.negro,fontFace:"Calibri",margin:0});
    y+=0.2;sl.addShape("line",{x:P,y,w:LW,h:0,line:{color:C.negro,width:1}});y+=0.08;
  }
  function divL(){sl.addShape("line",{x:P,y,w:LW,h:0,line:{color:C.grisLin,width:0.5}});y+=0.1;}
  function rowL(lab,val){
    const vs=typeof val==="number"?((val*100).toFixed(2)+"%"):String(val||"");
    sl.addText(String(lab),{x:P,y,w:1.5,h:0.19,fontSize:7.5,fontFace:"Calibri",margin:0,wrap:true});
    sl.addText(vs,{x:P+1.5,y,w:LW-1.5,h:0.19,fontSize:7.5,fontFace:"Calibri",align:"right",bold:true,margin:0});
    y+=0.21;
  }

  secL("Composición por Moneda");
  const mon=d.comp_moneda||[];mon.length?mon.forEach(([n,v])=>rowL(n,v)):rowL("—","—");divL();
  secL("Composición por Instrumento");
  const ins=d.comp_instrumentos||[];ins.length?ins.forEach(([n,v])=>rowL(n,v)):rowL("—","—");divL();
  secL("Composición por Duración");
  const dur=d.comp_duracion||[];dur.length?dur.forEach(([n,v])=>rowL(n,v)):rowL("—","—");

  /* ── Col derecha: Glosario + Disclaimer (TEXTO FIJO) ──────────── */
  let yr=0.15;
  sl.addText("Glosario",{x:DER,y:yr,w:DER_W,h:0.27,
    fontSize:13,bold:true,color:C.negro,fontFace:"Calibri",margin:0});
  yr+=0.3;sl.addShape("line",{x:DER,y:yr,w:DER_W,h:0,line:{color:C.negro,width:1.5}});yr+=0.1;

  const GL=[
    ["Riesgo de Mercado","Este es el riesgo de una variación adversa en el precio o tasa de mercado en los instrumentos en que invierte el Fondo"],
    ["Riesgo de Crédito","Es la posible pérdida que se asume como consecuencia del incumplimiento de las obligaciones directas, indirectas o de derivados que conllevan al no pago parcial, total o falta de oportunidad del pago de los emisores de los instrumentos en que invierte el Fondo y que pudieran ocasionar una pérdida financiera."],
    ["Riesgo de Liquidez","Riesgo asociado a la capacidad de generación de recursos del Fondo para cumplir con sus obligaciones de rescate o vencimiento del mismo."],
    ["Riesgo Profundidad de Mercado","Corresponde a la posibilidad de comprar o vender un activo financiero al valor de mercado en un período de tiempo acorde a las características del instrumento. Períodos de alta volatilidad de mercado afectan negativamente estos tiempos."],
    ["Riesgo de Tasa de interés","Exposición a pérdidas ocasionadas por cambios adversos en las tasas de interés de mercado y que afecten el valor de los instrumentos, contratos y demás operaciones registradas en el balance."],
    ["Riesgo de Moneda","Es la exposición a pérdidas ocasionadas por cambios adversos en el valor en moneda nacional de las monedas extranjeras que están expresados a los instrumentos, contratos y demás operaciones del balance del Fondo."],
    ["Riesgo Sectorial","Este riesgo está asociado a condiciones de mercado adversas que pueden afectar a un sector industrial en particular y por ende la rentabilidad del Fondo."],
    ["Gastos del fondo","Corresponden a los gastos directos e indirectos necesarios para el correcto funcionamiento del fondo los que están detallados en el Reglamento Interno"],
    ["Forma de Ingreso y Pago del Fondo","La moneda en que el inversionista entra al fondo es aportando pesos chilenos, y al rescate de las cuotas, el fondo le entrega pesos chilenos."],
  ];
  GL.forEach(([t,desc])=>{
    sl.addText([{text:t,options:{bold:true,breakLine:true}},{text:desc}],
      {x:DER,y:yr,w:DER_W,h:0.43,fontSize:7,color:C.grisOsc,fontFace:"Calibri",
       valign:"top",margin:0,wrap:true});
    yr+=0.46;
  });
  yr+=0.08;lin(sl,yr,DER,DER_W,C.grisLin,0.5);yr+=0.14;
  sl.addText("Disclaimer",{x:DER,y:yr,w:DER_W,h:0.25,
    fontSize:13,bold:true,color:C.negro,fontFace:"Calibri",margin:0});
  yr+=0.28;lin(sl,yr,DER,DER_W,C.negro,1.5);yr+=0.1;
  sl.addText("Conforme a la Ley Única de Fondos, las administradoras de fondos de inversión privados están sujetas a las obligaciones de información establecidas por la Comisión para el Mercado Financiero. Tales fondos no están sometidos a fiscalización de la Comisión y no hacemos oferta pública de sus cuotas.",
    {x:DER,y:yr,w:DER_W,h:0.7,fontSize:7.5,color:C.grisOsc,fontFace:"Calibri",
     valign:"top",margin:0,wrap:true});
}

/* ── MAIN ─────────────────────────────────────────────────────────── */
async function main(){
  const ai=process.argv.indexOf("--data"),oi=process.argv.indexOf("--out");
  if(ai===-1||oi===-1){
    console.error("Uso: node generar_folleto.js --data datos.json --out output.pptx");
    process.exit(1);
  }
  const d=JSON.parse(fs.readFileSync(process.argv[ai+1],"utf8"));
  const pres=new pptxgen();
  pres.defineLayout({name:"A4P",width:8.27,height:11.69});
  pres.layout="A4P";
  slide1(pres,d);
  slide2(pres,d);
  await pres.writeFile({fileName:process.argv[oi+1]});
  console.log("OK:"+process.argv[oi+1]);
}
main().catch(e=>{console.error(e);process.exit(1);});
