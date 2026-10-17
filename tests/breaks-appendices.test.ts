import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {strToU8,zipSync,strFromU8,unzipSync} from 'fflate';
import {readFile} from 'node:fs/promises';
import {Schema} from '@tiptap/pm/model';
import {EditorState,TextSelection} from '@tiptap/pm/state';
import {breakTransaction,blankPageTransaction} from '@prisma/writer/document-breaks';
import {importDocx} from '@prisma/writer/document-import';
import {exportDocx} from '@prisma/writer/document-export';
import {setFontLoader} from '@prisma/core/fonts';
import {defaultLayout,plainText,type FlowDocument} from '@prisma/writer/flow';
import {collectDocumentEntries} from '@prisma/writer/academic';
import {insertBlankPage} from '@prisma/pdf/page-insertion';
import {planPages} from '@prisma/writer/pagination';
Object.assign(globalThis,{DOMParser});
setFontLoader(async file=>new Uint8Array(await readFile('public/fonts/'+file)));
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run=(s:string)=>`<w:r><w:t>${s}</w:t></w:r>`;
const packageDoc=(body:string,styles='')=>zipSync({'word/document.xml':strToU8(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`),'word/styles.xml':strToU8(`<w:styles xmlns:w="${W}">${styles}</w:styles>`)});
const schema=new Schema({nodes:{doc:{content:'block+'},paragraph:{group:'block',content:'text*',attrs:{id:{default:''},spaceAfter:{default:null}}},text:{group:'inline'},pageBreak:{group:'block',atom:true},columnBreak:{group:'block',atom:true},sectionBreak:{group:'block',atom:true,attrs:{kind:{default:'nextPage'}}}}});
const p=(text='')=>schema.node('paragraph',null,text?schema.text(text):undefined);
const state=(nodes:any[],pos:number)=>{const doc=schema.node('doc',null,nodes);return EditorState.create({doc,selection:TextSelection.create(doc,pos)});};

test('page and section breaks split at the caret, preserve both halves and support undo transactions',()=>{
 for(const kind of ['pageBreak','columnBreak','sectionBreak'] as const){
  const original=state([p('AntesDespués')],6),tr=breakTransaction(original,kind);
  assert.deepEqual(tr.doc.content.content.map(n=>n.type.name),['paragraph',kind,'paragraph']);
  assert.equal(tr.doc.child(0).textContent,'Antes');assert.equal(tr.doc.child(2).textContent,'Después');
  assert.equal(original.doc.textContent,'AntesDespués');assert(tr.selection.$from.parent.isTextblock);
 }
 const end=breakTransaction(state([p('Final')],6),'sectionBreak');
 assert.equal(end.doc.lastChild?.type.name,'paragraph');
});

test('blank page insertion uses measured boundaries inside text, first page, last page and existing breaks',()=>{
 const original=state([p('PrimeraSegunda')],3),starts=[{pos:0,page:1},{pos:8,page:2}];
 const below=blankPageTransaction(original,starts,2,'below');
 assert.deepEqual(below.doc.content.content.map(n=>[n.type.name,n.textContent]),[['paragraph','Primera'],['pageBreak',''],['paragraph',''],['pageBreak',''],['paragraph','Segunda']]);
 const above=blankPageTransaction(original,starts,2,'above');assert.equal(above.doc.child(0).textContent,'');assert.equal(above.doc.child(1).type.name,'pageBreak');
 const end=blankPageTransaction(original,starts,2,'end');assert.equal(end.doc.lastChild?.textContent,'');assert.equal(end.doc.child(end.doc.childCount-2).type.name,'pageBreak');
 const explicit=state([p('Uno'),schema.node('pageBreak'),p('Dos')],2);
 const inserted=blankPageTransaction(explicit,[{pos:0,page:1},{pos:6,page:2}],2,'below');
 assert.equal(inserted.doc.content.content.filter(n=>n.type.name==='pageBreak').length,2);
 assert.equal(inserted.doc.textContent,'UnoDos');
});

test('page insertion without a page map follows explicit page boundaries',()=>{
 const doc={type:'doc',content:[{type:'paragraph',content:[{type:'text',text:'Uno'}]},{type:'pageBreak'},{type:'paragraph',content:[{type:'text',text:'Dos'}]}]};
 const result=insertBlankPage(doc,1,'below').content!;
 assert.deepEqual(result.map(n=>n.type),['paragraph','pageBreak','paragraph','pageBreak','paragraph']);
 assert.equal(plainText(result.at(-1)!),'Dos');
});

test('Word run page breaks preserve text order and repeated breaks instead of collapsing a blank page',()=>{
 const result=importDocx(packageDoc(`<w:p>${run('Uno')}<w:r><w:br w:type="page"/><w:br w:type="page"/></w:r>${run('Tres')}<w:r><w:br w:type="column"/></w:r>${run('Columna')}</w:p><w:sectPr/>`)).flow.content.content!;
 assert.deepEqual(result.map(n=>n.type),['paragraph','pageBreak','pageBreak','paragraph','columnBreak','paragraph']);
 assert.deepEqual(result.filter(n=>n.type==='paragraph').map(plainText),['Uno','Tres','Columna']);
});

test('appendix SEQ fields without heading styles become appendix entries and custom localized titles retain levels',()=>{
 const bytes=packageDoc(`<w:p><w:fldSimple w:instr='TOC \\c "Apéndice" \\h'>${run('Viejo índice')}</w:fldSimple></w:p><w:p>${run('Apéndice ')}<w:fldSimple w:instr='SEQ Apéndice \\* ALPHABETIC'>${run('A')}</w:fldSimple>${run('. Instrumento de prueba')}</w:p><w:p><w:pPr><w:pStyle w:val="Ttulo3Ap"/></w:pPr>${run('A.1 Alcance')}</w:p><w:p>${run('Véase el apéndice A para detalles.')}</w:p><w:sectPr/>`,`<w:style w:type="paragraph" w:styleId="Ttulo3Ap"><w:name w:val="Título 3 Ap"/></w:style>`);
 const {flow}=importDocx(bytes),entries=collectDocumentEntries(flow.content);
 assert.equal(entries.appendices.length,1);assert.equal(entries.appendices[0].title,'Instrumento de prueba');
 assert.equal(flow.content.content![2].attrs?.level,3);assert.equal(flow.content.content![3].type,'paragraph');
 assert.equal((flow.content.content![0].attrs?.entries as any[]).length,1);
});

test('appendix indexes export as updating TOC fields with alphabetic SEQ and internal bookmarks',async()=>{
 const flow:FlowDocument={layout:defaultLayout(),content:{type:'doc',content:[{type:'generatedIndex',attrs:{kind:'appendices'}},{type:'appendix',attrs:{id:'appendix-a',title:'Instrumento',breakBefore:true}},{type:'paragraph',content:[{type:'text',text:'Contenido A'}]},{type:'appendix',attrs:{id:'appendix-b',title:'Resultados',breakBefore:true}}]}};
 const bytes=await exportDocx(flow,'Apéndices'),files=unzipSync(bytes),xml=strFromU8(files['word/document.xml']);
 assert.match(xml,/TOC[^<]*\\c &quot;Apéndice&quot;/);assert.equal((xml.match(/SEQ Apéndice/g)||[]).length,2);assert.match(xml,/ALPHABETIC/);assert.match(xml,/w:bookmarkStart/);assert.match(strFromU8(files['word/settings.xml']),/updateFields/);
 const imported=importDocx(bytes).flow;
 assert.equal(imported.content.content![0].attrs?.kind,'appendices');assert.equal(collectDocumentEntries(imported.content).appendices.length,2);
 assert.equal(imported.content.content!.find(n=>n.attrs?.appendix)?.attrs?.breakBefore,true);
});

test('section parity, page numbering restart and explicit page breaks survive DOCX round trip',async()=>{
 const kinds=['continuous','nextPage','oddPage','evenPage'];
 const content:any[]=[{type:'paragraph',content:[{type:'text',text:'Inicio'}]}];
 for(const kind of kinds)content.push({type:'sectionBreak',attrs:{kind,id:kind,layout:{...defaultLayout(),orientation:kind==='oddPage'?'landscape':'portrait',pageNumberStart:5},linkToPrevious:false,restartNumbering:true}},{type:'paragraph',content:[{type:'text',text:kind}]});
 content.push({type:'pageBreak'},{type:'paragraph',content:[]},{type:'pageBreak'},{type:'paragraph',content:[{type:'text',text:'Tras página vacía'}]});
 const bytes=await exportDocx({layout:defaultLayout(),content:{type:'doc',content}},'Secciones');
 const imported=importDocx(bytes).flow.content.content!;
 const sections=imported.filter(n=>n.type==='sectionBreak');
 assert.deepEqual(sections.map(n=>n.attrs?.kind),kinds);
 assert(sections.every(n=>n.attrs?.restartNumbering===true&&n.attrs?.linkToPrevious===false));
 assert.equal((sections[2].attrs?.layout as any).orientation,'landscape');
 assert.equal(imported.filter(n=>n.type==='pageBreak').length,2);
 const plan=planPages([{pos:1,top:0,bottom:10},{pos:5,top:10,bottom:10,force:1,startOn:'oddPage'}],100,10,10);
 assert.equal(plan.count,3);assert.equal(plan.breaks[0].page,3);
});

test('appendix headings have one native heading style and their cached index keeps known pages',async()=>{
 const flow:FlowDocument={layout:defaultLayout(),content:{type:'doc',content:[{type:'generatedIndex',attrs:{kind:'contents'}},{type:'generatedIndex',attrs:{kind:'appendices'}},{type:'appendix',attrs:{id:'appendix-page',title:'Instrumento',page:5}}]}};
 assert.equal(collectDocumentEntries(flow.content).appendices[0].page,5);
 const files=unzipSync(await exportDocx(flow,'Apéndices con páginas'));
 const doc=new DOMParser().parseFromString(strFromU8(files['word/document.xml']),'application/xml');
 const heading=Array.from(doc.getElementsByTagNameNS(W,'p')).find(p=>Array.from(p.getElementsByTagNameNS(W,'fldSimple')).some(f=>f.getAttributeNS(W,'instr')?.includes('SEQ Apéndice')))!;
 const styles=Array.from(heading.getElementsByTagNameNS(W,'pStyle'));
 assert.equal(styles.length,1);assert.equal(styles[0].getAttributeNS(W,'val'),'Heading1');
 const index=Array.from(doc.getElementsByTagNameNS(W,'sdt')).find(s=>s.textContent?.includes('TOC \\c'))!;
 assert.match(index.textContent??'',/5/);
});

test('linked sections inherit headers in Word instead of exporting independent copies',async()=>{
 const layout={...defaultLayout(),header:'Cabecera compartida',footer:'Pie compartido'};
 const content: any[]=[{type:'paragraph',content:[{type:'text',text:'Primera sección'}]},{type:'sectionBreak',attrs:{id:'linked',kind:'nextPage',layout,linkToPrevious:true}},{type:'paragraph',content:[{type:'text',text:'Sección vinculada'}]}];
 const bytes=await exportDocx({layout,content:{type:'doc',content}},'Cabeceras');
 const parsed=new DOMParser().parseFromString(strFromU8(unzipSync(bytes)['word/document.xml']),'application/xml');
 const sections=[...Array.from(parsed.getElementsByTagNameNS(W,'sectPr'))];
 assert.equal(sections[0].getElementsByTagNameNS(W,'headerReference').length,1);
 assert.equal(sections[1].getElementsByTagNameNS(W,'headerReference').length,0);
 assert.equal(importDocx(bytes).flow.content.content!.find(n=>n.type==='sectionBreak')!.attrs?.linkToPrevious,true);
 assert.equal(importDocx(bytes).flow.layout.differentFirst,false);
});

test('first-page header settings belong to each section even when header text is linked',async()=>{
 const initial={...defaultLayout(),header:'Cabecera común',differentFirst:false};
 const next={...initial,differentFirst:true,headerFirst:'Primera de sección'};
 const content:any[]=[{type:'paragraph',content:[{type:'text',text:'Inicio'}]},{type:'sectionBreak',attrs:{id:'first-page',kind:'nextPage',layout:next,linkToPrevious:true}},{type:'paragraph',content:[{type:'text',text:'Nueva sección'}]}];
 const bytes=await exportDocx({layout:initial,content:{type:'doc',content}},'Cabeceras de sección');
 const flow=importDocx(bytes).flow;
 assert.equal(flow.layout.differentFirst,false);
 assert.equal((flow.content.content!.find(n=>n.type==='sectionBreak')!.attrs!.layout as any).differentFirst,true);
});
