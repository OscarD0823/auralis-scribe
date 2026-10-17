import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {DOMParser} from '@xmldom/xmldom';
import {Schema} from '@tiptap/pm/model';
import {EditorState,TextSelection} from '@tiptap/pm/state';
import {history,undo} from '@tiptap/pm/history';
import type {Editor} from '@tiptap/core';
import {strFromU8,strToU8,zipSync,unzipSync} from 'fflate';
import * as mupdf from 'mupdf';
import {setFontLoader} from '@prisma/core/fonts';
import {defaultLayout,buildFlowDocument,type RichNode,type FlowDocument} from '../src/flow';
import {importDocx} from '../src/document-import-core';
import {exportDocx} from '../src/document-export';
import {repeatedTableHeaderCount,setTablePagination,distributeTableColumns} from '../src/table-tools';
Object.assign(globalThis,{DOMParser});
setFontLoader(async name=>new Uint8Array(fs.readFileSync('public/fonts/'+name)));
const p=(text:string,attrs={}):RichNode=>({type:'paragraph',attrs:{fontFamily:'PrismaSans',fontSize:12,lineHeight:1,...attrs},content:[{type:'text',text}]});
const cell=(text:string,header=false):RichNode=>({type:header?'tableHeader':'tableCell',content:[p(text)]});
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
test('only contiguous header rows repeat; a first-column header and a merge into body rows do not repeat',()=>{
 assert.equal(repeatedTableHeaderCount({content:[{content:[cell('Columna',true),cell('Dato')]},{content:[cell('Fila',true),cell('Dos')]}]}),0);
 assert.equal(repeatedTableHeaderCount({content:[{content:[cell('Cabecera',true)]},{attrs:{repeatHeader:false},content:[cell('Estilo',true)]}]}),1);
 assert.equal(repeatedTableHeaderCount({content:[{attrs:{repeatHeader:true},content:[{...cell('Combinada'),attrs:{rowspan:2}}]},{content:[cell('Dato')]}]}),0);
});
test('DOCX table properties preserve explicit false flags and vertical alignment on a structural round trip',async()=>{
 const row=(props:string,align:string)=>`<w:tr><w:trPr>${props}</w:trPr><w:tc><w:tcPr><w:vAlign w:val="${align}"/></w:tcPr><w:p><w:r><w:t>Celda</w:t></w:r></w:p></w:tc></w:tr>`;
 const bytes=zipSync({'word/document.xml':strToU8(`<w:document xmlns:w="${W}"><w:body><w:tbl>${row('<w:tblHeader/>','center')}${row('<w:tblHeader w:val="false"/><w:cantSplit/>','bottom')}</w:tbl><w:sectPr/></w:body></w:document>`)});
 const flow=importDocx(bytes).flow,table=flow.content.content![0],rows=table.content!;
 assert.equal(rows[0].attrs?.repeatHeader,true);assert.equal(rows[1].attrs?.repeatHeader,false);assert.equal(rows[1].attrs?.cantSplit,true);
 assert.equal(rows[0].content![0].attrs?.verticalAlign,'center');assert.equal(rows[1].content![0].attrs?.verticalAlign,'bottom');
 const again=importDocx(await exportDocx({content:flow.content,layout:flow.layout},'Tabla')).flow.content.content![0];
 assert.equal(repeatedTableHeaderCount(again),1);assert.equal(again.content![1].attrs?.cantSplit,true);assert.equal(again.content![1].content![0].attrs?.verticalAlign,'bottom');
 const xml=new DOMParser().parseFromString(strFromU8(unzipSync(await exportDocx({content:flow.content,layout:flow.layout},'Tabla'))['word/document.xml']),'application/xml');
 const flags=Array.from(xml.getElementsByTagNameNS(W,'tblHeader'));assert.equal(flags.filter(f=>!['false','0','off'].includes(f.getAttributeNS(W,'val')??'')).length,1);
});
test('table commands preserve merged cells, all text and undo as a single transaction',()=>{
 const attrs={colspan:{default:1},rowspan:{default:1},colwidth:{default:null}};
 const schema=new Schema({nodes:{doc:{content:'table+'},text:{group:'inline'},paragraph:{content:'text*'},table:{content:'tableRow+',tableRole:'table'},tableRow:{content:'(tableCell|tableHeader)+',tableRole:'row',attrs:{repeatHeader:{default:null},cantSplit:{default:false}}},tableCell:{content:'paragraph+',tableRole:'cell',attrs},tableHeader:{content:'paragraph+',tableRole:'header_cell',attrs}}});
 const table=schema.nodeFromJSON({type:'table',content:[{type:'tableRow',content:[cell('A',true),cell('B',true)]},{type:'tableRow',content:[{...cell('Combinada'),attrs:{colspan:2}}]}]});
 const doc=schema.node('doc',undefined,[table]);let state=EditorState.create({doc,selection:TextSelection.create(doc,4),plugins:[history()]});
 const editor={get state(){return state;},view:{dispatch:(tr:any)=>state=state.apply(tr),nodeDOM:()=>null},commands:{focus:()=>true}} as unknown as Editor;
 const original=state.doc.toJSON();setTablePagination(editor,1,true);assert.equal(state.doc.firstChild!.child(1).attrs.cantSplit,true);
 assert(undo(state,tr=>state=state.apply(tr)));assert.deepEqual(state.doc.toJSON(),original);
 distributeTableColumns(editor,400);assert.deepEqual(state.doc.firstChild!.child(0).child(0).attrs.colwidth,[200]);assert.deepEqual(state.doc.firstChild!.child(1).child(0).attrs.colwidth,[200,200]);assert.equal(state.doc.textContent,'ABCombinada');
});
const pdfTexts=async(flow:FlowDocument)=>{const result=await buildFlowDocument(flow),doc=mupdf.Document.openDocument(result.bytes,'application/pdf');return Array.from({length:doc.countPages()},(_,i)=>doc.loadPage(i).toStructuredText('preserve-whitespace').asText());};
test('PDF repeats multiple table headers on continuation pages and every body token appears exactly once',async()=>{
 const rows:RichNode[]=[{type:'tableRow',attrs:{repeatHeader:true},content:[cell('CabeceraPrincipal',true),cell('Valor',true)]},{type:'tableRow',attrs:{repeatHeader:true},content:[cell('CabeceraSecundaria',true),cell('Unidad',true)]},...Array.from({length:36},(_,i)=>({type:'tableRow',content:[cell('RegistroUnico'+String(i).padStart(3,'0')),cell('Dato '+i)]}))];
 const texts=await pdfTexts({layout:{...defaultLayout(),format:'custom',customWidth:180,customHeight:100,top:10,bottom:10},content:{type:'doc',content:[{type:'table',content:rows}]}});
 assert(texts.length>2);for(const text of texts){assert.equal(text.split('CabeceraPrincipal').length-1,1);assert.equal(text.split('CabeceraSecundaria').length-1,1);}
 for(let i=0;i<36;i++)assert.equal(texts.join('').split('RegistroUnico'+String(i).padStart(3,'0')).length-1,1);
});
test('PDF keeps an ordinary row together when requested and otherwise allows its content to continue',async()=>{
 const make=(keep:boolean):FlowDocument=>({layout:{...defaultLayout(),format:'custom',customWidth:180,customHeight:100,top:10,bottom:10},content:{type:'doc',content:[p('Introducción',{spaceAfter:90}),{type:'table',content:[{type:'tableRow',content:[cell('Fila corta')]},{type:'tableRow',attrs:{cantSplit:keep},content:[{type:'tableCell',content:Array.from({length:10},(_,i)=>p('Largo'+String(i).padStart(2,'0')))}]}]}]}});
 const split=await pdfTexts(make(false)),keep=await pdfTexts(make(true));assert(split[0].includes('Largo00'));assert(!keep[0].includes('Largo00'));assert(keep[1].includes('Largo00'));
 for(const text of [split.join(''),keep.join('')])for(let i=0;i<10;i++)assert.equal(text.split('Largo'+String(i).padStart(2,'0')).length-1,1);
});
