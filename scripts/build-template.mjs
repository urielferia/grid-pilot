import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const wb=Workbook.create();
const names=['Resumen','Visitas','Inventario','Movimientos'];
const sheets=Object.fromEntries(names.map(n=>[n,wb.worksheets.add(n)]));
const fields={
 Visitas:[['Id','ID','text'],['Fecha','Fecha','date'],['Cliente','Cliente','text'],['Sitio','Sitio','text'],['Tecnico','Técnico','text'],['Estado','Estado','text'],['Descripcion','Trabajo','text'],['CostoMateriales','Costo materiales','number',true,true]],
 Inventario:[['Id','ID material','text'],['Descripcion','Descripción','text'],['Unidad','Unidad','text'],['SaldoInicial','Saldo inicial','number'],['Minimo','Mínimo','number'],['CostoReferencia','Costo referencia','number'],['Entradas','Entradas','number',true,true],['Consumos','Consumos','number',true,true],['Existencia','Existencia','number',true,true]],
 Movimientos:[['Id','ID','text'],['Fecha','Fecha','date'],['Tipo','Tipo','text'],['Material','Material','text'],['Cantidad','Cantidad','number'],['Visita','Visita','text',false],['CostoUnitario','Costo unitario','number'],['Importe','Importe','number',true,true]]
};
const profile={Id:'gridpilot-demo-v1',SummarySheet:'Resumen',Tables:Object.entries(fields).map(([Name,fs])=>({Name,Sheet:Name,Fields:fs.map(([Key,Header,Type,Required=true,Calculated=false])=>({Key,Header,Type,Required,Calculated}))}))};
const serial=day=>(Date.UTC(2026,9,day)-Date.UTC(1899,11,30))/86400000;
const materials=['Cable UTP','Conector RJ45','Canaleta','Cinta aislante','Tornillo','Taquete','Adaptador','Fusible','Terminal','Brida'];
for(const [name,s] of Object.entries(sheets)){
 s.showGridLines=false;
 s.getRange('A1:J28').format.font={name:'Arial',size:10,color:'#182A40'};
 s.getRange('A1:J28').format.rowHeight=24;
 s.getRange('A1:J28').format.columnWidth=18;
 s.getRange('A2').values=[[name==='Resumen'?'GridPilot · Visitas e inventario':name]];
 s.getRange('A2').format.font={name:'Arial',size:16,bold:true};
 s.getRange('A3').values=[['DEMOSTRACIÓN — datos ficticios']];
 s.getRange('A3').format.font={italic:true,color:'#566779'};
 if(name!=='Resumen'){
  const f=fields[name],last=String.fromCharCode(64+f.length),count=name==='Movimientos'?20:10;
  s.getRange(`A5:${last}5`).values=[f.map(x=>x[1])];
  s.tables.add(`A5:${last}${5+count}`,true,name);
  s.getRange(`A5:${last}5`).format={fill:'#213D61',font:{bold:true,color:'#FFFFFF'},rowHeight:32,wrapText:true};
  s.getRange(`A6:${last}${5+count}`).format.fill='#FFF8E6';
  s.freezePanes.freezeRows(5);
  s.getRange('A6:A25').setNumberFormat('@');
 }
}
sheets.Visitas.getRange('A6:H15').values=Array.from({length:10},(_,i)=>[`V${String(i+1).padStart(3,'0')}`,serial(i+1),`Cliente demo ${i+1}`,`Sitio ${i+1}`,i%2?'Ana':'Luis',i<8?'Realizada':'Programada','Mantenimiento preventivo',null]);
sheets.Inventario.getRange('A6:I15').values=Array.from({length:10},(_,i)=>[`M${String(i+1).padStart(3,'0')}`,materials[i],'pieza',20+i,10,5*(i+1),null,null,null]);
sheets.Movimientos.getRange('A6:H25').values=Array.from({length:20},(_,i)=>[`MOV${String(i+1).padStart(3,'0')}`,serial(i<10?1:i-8),i<10?'Entrada':'Consumo',`M${String(i%10+1).padStart(3,'0')}`,i<10?10:2, i<10?'':`V${String((i-10)%8+1).padStart(3,'0')}`,5*(i%10+1),null]);
// Structured table references let Excel extend and repair formulas when rows/columns change.
for(let r=6;r<=25;r++) sheets.Movimientos.getRange(`H${r}`).formulas=[[`=E${r}*G${r}`]];
for(let r=6;r<=15;r++){
 sheets.Visitas.getRange(`H${r}`).formulas=[[`=SUMIFS(Movimientos[Importe],Movimientos[Visita],A${r},Movimientos[Tipo],"Consumo")`]];
 sheets.Inventario.getRange(`G${r}:I${r}`).formulas=[[
 `=SUMIFS(Movimientos[Cantidad],Movimientos[Material],A${r},Movimientos[Tipo],"Entrada")`,
 `=SUMIFS(Movimientos[Cantidad],Movimientos[Material],A${r},Movimientos[Tipo],"Consumo")`,`=D${r}+G${r}-H${r}`]];
}
for(const name of ['Visitas','Movimientos']) sheets[name].getRange('B6:B25').setNumberFormat('dd/mm/yyyy');
for(const [name,range] of [['Visitas','H6:H15'],['Inventario','F6:F15'],['Movimientos','G6:H25']]) sheets[name].getRange(range).setNumberFormat('"MXN "#,##0.00');
for(const [name,range] of [['Visitas','H6:H15'],['Inventario','G6:I15'],['Movimientos','H6:H25']]) sheets[name].getRange(range).format.fill='#E9F0F7';
sheets.Inventario.getRange('B1:B25').format.columnWidth=25;
sheets.Visitas.getRange('G1:G25').format.columnWidth=32;
sheets.Visitas.getRange('F6:F15').dataValidation={rule:{type:'list',values:['Programada','Realizada']}};
sheets.Movimientos.getRange('C6:C25').dataValidation={rule:{type:'list',values:['Entrada','Consumo']}};
sheets.Inventario.getRange('I6:I15').conditionalFormats.addCustom('I6<E6',{fill:'#FFE0DC',font:{color:'#A52121'}});
const s=sheets.Resumen;
s.getRange('A1:A15').format.columnWidth=46;s.getRange('B1:B15').format.columnWidth=24;
s.getRange('A5:B6').values=[['Desde',serial(1)],['Hasta',serial(31)]];
s.getRange('B5:B6').setNumberFormat('dd/mm/yyyy');s.getRange('B5:B6').format.fill='#FFF8E6';
s.getRange('A8:B11').values=[['Indicador','Resultado'],['Visitas realizadas',null],['Costo de materiales del período',null],['Materiales bajo mínimo',null]];
s.getRange('A8:B8').format={fill:'#213D61',font:{bold:true,color:'#FFFFFF'}};
s.getRange('B9').formulas=[['=COUNTIFS(Visitas[Estado],"Realizada",Visitas[Fecha],">="&B5,Visitas[Fecha],"<="&B6)']];
s.getRange('B10').formulas=[['=SUMIFS(Movimientos[Importe],Movimientos[Tipo],"Consumo",Movimientos[Fecha],">="&B5,Movimientos[Fecha],"<="&B6)']];
s.getRange('B11').formulas=[['=SUMPRODUCT(--(Inventario[Existencia]<Inventario[Mínimo]))']];
s.getRange('B10').setNumberFormat('"MXN "#,##0.00');
s.getRange('A14').values=[['Amarillo: entradas editables. Azul claro: cálculos.']];
wb.recalculate();
await fs.mkdir('assets',{recursive:true}); await fs.mkdir('artifacts/previews',{recursive:true});
console.log((await wb.inspect({kind:'table',range:'Resumen!A5:B11',include:'values,formulas',tableMaxRows:10,tableMaxCols:2})).ndjson);
for(const name of names){const img=await wb.render({sheetName:name,range:name==='Resumen'?'A1:C15':`A1:${name==='Inventario'?'I':'H'}${name==='Movimientos'?25:15}`,scale:1.3});await fs.writeFile(`artifacts/previews/${name}.png`,new Uint8Array(await img.arrayBuffer()));}
await (await SpreadsheetFile.exportXlsx(wb)).save('assets/GridPilot-Demo.xlsx');
await fs.writeFile('assets/profile.json',JSON.stringify(profile,null,2));
await fs.writeFile('assets/expected.json',JSON.stringify({visits:10,materials:10,movements:20,completed:8,materialCost:550,belowMinimum:0,stock:Array.from({length:10},(_,i)=>28+i)},null,2));
