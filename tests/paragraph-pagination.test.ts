import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {strFromU8,unzipSync} from 'fflate';
import {planPages} from '@prisma/writer/pagination';
import {defaultLayout,type FlowDocument} from '@prisma/writer/flow';
import {exportDocx} from '@prisma/writer/document-export';
import {importDocx} from '@prisma/writer/document-import';
import {readFile} from 'node:fs/promises';
import {setFontLoader} from '@prisma/core/fonts';
Object.assign(globalThis,{DOMParser});
setFontLoader(async file=>new Uint8Array(await readFile('public/fonts/'+file)));
test('heading and next line move together instead of leaving a heading at the page bottom',()=>{
 const units=[{pos:1,top:0,bottom:60},{pos:20,top:65,bottom:75,keepWithNext:true},{pos:30,top:77,bottom:90}];
 assert.deepEqual(planPages(units,100,10,10).breaks.map(b=>b.pos),[20]);
 assert.deepEqual(planPages(units.map(u=>({...u,keepWithNext:false})),100,10,10).breaks.map(b=>b.pos),[30]);
});
test('explicit breaks and oversized keep chains remain bounded and do not cause an empty page loop',()=>{
 const forced=planPages([{pos:1,top:0,bottom:70,keepWithNext:true},{pos:20,top:70,bottom:70,force:1},{pos:21,top:70,bottom:90}],100,10,10);
 assert.equal(forced.breaks[0].pos,20);
 const large=Array.from({length:2000},(_,i)=>({pos:i,top:i*20,bottom:i*20+20,keepWithNext:true}));
 assert(planPages(large,100,10,10).count<2000);
});
test('paragraph pagination options survive DOCX export, inherited styles and explicit false',async()=>{
 const flow:FlowDocument={layout:defaultLayout(),content:{type:'doc',content:[
  {type:'heading',attrs:{level:1,id:'title'},content:[{type:'text',text:'Título unido'}]},
  {type:'paragraph',attrs:{id:'together',keepNext:true,keepLines:true,breakBefore:true},content:[{type:'text',text:'Párrafo junto'}]},
  {type:'heading',attrs:{id:'separate',level:2,keepNext:false,keepLines:false},content:[{type:'text',text:'Título libre'}]},
 ]}};
 const bytes=await exportDocx(flow,'Párrafos'),xml=strFromU8(unzipSync(bytes)['word/document.xml']);
 assert.match(xml,/keepNext/);assert.match(xml,/keepLines/);
 const result=importDocx(bytes).flow.content.content??[];
 assert.equal(result[0].attrs?.keepNext,true);assert.equal(result[0].attrs?.keepLines,true);
 assert.equal(result[1].attrs?.keepNext,true);assert.equal(result[1].attrs?.keepLines,true);assert.equal(result[1].attrs?.breakBefore,true);
 assert.equal(result[2].attrs?.keepNext,false);assert.equal(result[2].attrs?.keepLines,false);
});
