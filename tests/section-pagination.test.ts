import test from 'node:test';
import assert from 'node:assert/strict';
import {defaultLayout,pageSize,mm,type PageLayout} from '../src/flow';
import {planSectionPages} from '../src/section-pagination';
import type {SectionInfo} from '../src/sections';
const portrait=defaultLayout(),landscape={...portrait,orientation:'landscape'} as PageLayout;
const section=(id:string,layout:PageLayout,kind:SectionInfo['kind']='nextPage',restartNumbering=false):SectionInfo=>({id,title:id,layout,kind,linkToPrevious:false,restartNumbering});

test('one editor can paginate portrait, landscape and custom sheets with independent sizes and margins',()=>{
 const custom={...portrait,format:'custom',customWidth:150,customHeight:180,left:15,right:10} as PageLayout;
 const result=planSectionPages([{pos:0,top:0,bottom:10},{pos:10,top:10,bottom:10,section:section('wide',landscape)},{pos:11,top:10,bottom:20},{pos:20,top:20,bottom:20,section:section('small',custom)},{pos:21,top:20,bottom:30}],portrait,1200);
 assert.equal(result.count,3);
 assert.deepEqual(result.pages.map(p=>p.layout.orientation),['portrait','landscape','portrait']);
 assert.equal(result.pages[1].width,pageSize(landscape)[0]*4/3);
 assert.equal(result.pages[2].height,pageSize(custom)[1]*4/3);
 for(let i=1;i<3;i++)assert.equal(result.pages[i].top,result.pages[i-1].top+result.pages[i-1].height+28);
 assert.equal(result.pages[2].left,(1200-result.pages[2].width)/2);
});

test('automatic overflow uses the landscape height instead of the height of the initial portrait sheet',()=>{
 const available=(pageSize(landscape)[1]-mm(landscape.top+landscape.bottom))*4/3;
 const result=planSectionPages([{pos:0,top:0,bottom:10},{pos:2,top:10,bottom:10,section:section('wide',landscape)},{pos:3,top:10,bottom:20},{pos:4,top:available+15,bottom:available+25}],portrait,1200);
 assert.equal(result.count,3);assert.equal(result.pages[2].layout.orientation,'landscape');
 assert.equal(result.breaks.at(-1)?.pos,4);
});

test('odd sections insert a blank sheet and restart visible page numbers without losing physical parity',()=>{
 const next={...landscape,pageNumberStart:5,header:'Nueva cabecera'};
 const result=planSectionPages([{pos:0,top:0,bottom:10},{pos:10,top:10,bottom:10,section:section('odd',next,'oddPage',true)},{pos:11,top:10,bottom:20},{pos:12,top:20,bottom:20,force:1}],portrait,1200);
 assert.equal(result.count,4);assert.deepEqual(result.pages.map(p=>p.number),[1,2,5,6]);
 assert.deepEqual(result.pages.map(p=>p.sectionPage),[0,1,0,1]);
 assert.equal(result.pages[1].layout.orientation,'portrait');assert.equal(result.pages[2].layout.header,'Nueva cabecera');
});

test('continuous sections stay on their sheet unless the paper geometry changes',()=>{
 const result=planSectionPages([{pos:0,top:0,bottom:10},{pos:10,top:10,bottom:10,section:section('same',portrait,'continuous')},{pos:11,top:10,bottom:20},{pos:20,top:20,bottom:20,section:section('wide',landscape,'continuous')},{pos:21,top:20,bottom:30}],portrait,1200);
 assert.equal(result.count,2);assert.equal(result.breaks.length,1);assert.equal(result.breaks[0].pos,20);
});

test('explicit empty pages and keep-with-next headings remain bounded',()=>{
 const small={...portrait,format:'custom',customWidth:100,customHeight:30,top:0,bottom:0} as PageLayout;
 const result=planSectionPages([{pos:0,top:0,bottom:100},{pos:10,top:101,bottom:116,keepWithNext:true},{pos:20,top:116,bottom:130},{pos:21,top:130,bottom:130,force:1},{pos:22,top:130,bottom:130,force:1}],small);
 assert.equal(result.count,4);assert.equal(result.breaks[0].pos,10);
});

test('a continuous section counts its shared first page before subsequent headers and numbering',()=>{
 const next={...portrait,differentFirst:true,header:'Nueva cabecera',pageNumberStart:7};
 const result=planSectionPages([{pos:0,top:0,bottom:10},{pos:1,top:10,bottom:10,force:2},{pos:2,top:10,bottom:20},{pos:3,top:20,bottom:20,section:section('continuous',next,'continuous',true)},{pos:4,top:20,bottom:20,force:1}],portrait);
 assert.equal(result.count,4);assert.equal(result.pages[3].sectionId,'continuous');
 assert.equal(result.pages[3].sectionPage,1);assert.equal(result.pages[3].number,8);
 assert.equal(result.pages[3].layout.header,'Nueva cabecera');
});
