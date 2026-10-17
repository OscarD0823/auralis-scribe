import {mm,pageSize,type PageLayout} from './flow';
import {inheritedSectionLayout,type SectionInfo} from './sections';
import type {PageGap,PageUnit} from './pagination';
import type {Node as ModelNode} from '@tiptap/pm/model';

export type EditorPage={top:number;left:number;width:number;height:number;layout:PageLayout;sectionId:string;sectionPage:number;number:number};
export type SectionPageUnit=PageUnit&{section?:SectionInfo};
const px=(mmValue:number)=>mm(mmValue)*4/3;
export function firstEditorPage(layout:PageLayout,canvasWidth=pageSize(layout)[0]*4/3):EditorPage {
 const [width,height]=pageSize(layout).map(v=>v*4/3);
 return {top:0,left:(canvasWidth-width)/2,width,height,layout,sectionId:'',sectionPage:0,number:layout.pageNumberStart??1};
}

export function editorSectionBlocks(doc:ModelNode,initial:PageLayout) {
 let section:SectionInfo={id:'',title:'Sección 1',kind:'nextPage',layout:initial,linkToPrevious:false,restartNumbering:false};
 const blocks:{node:ModelNode;pos:number;section:SectionInfo;nextSection?:SectionInfo}[]=[];
 doc.forEach((node,pos)=>{
  if(node.type.name==='sectionBreak'){
   const next={...node.attrs,layout:node.attrs.layout??initial} as SectionInfo;
   next.layout=inheritedSectionLayout(section.layout,next);
   blocks.push({node,pos,section,nextSection:next});section=next;
  }else blocks.push({node,pos,section});
 });
 return blocks;
}

/** Positions are measured before inserting page gaps, relative to the text origin. */
export function planSectionPages(units:SectionPageUnit[],initial:PageLayout,canvasWidth=pageSize(initial)[0]*4/3,gap=28) {
 const pages:EditorPage[]=[firstEditorPage(initial,canvasWidth)],breaks:PageGap[]=[];
 let active:SectionInfo={id:'',title:'',kind:'nextPage',layout:initial,linkToPrevious:false,restartNumbering:false};
 let added=0,sectionStart=0,numberOffset=(initial.pageNumberStart??1)-1;
 const origin=px(initial.top);
 const contentStart=(page:EditorPage)=>page.top+px(page.layout.top)-origin;
 const contentEnd=(page:EditorPage)=>page.top+page.height-px(page.layout.bottom)-origin;
 const append=(section:SectionInfo=active,begin=false)=>{
  if(begin){sectionStart=pages.length;if(section.restartNumbering)numberOffset=(section.layout.pageNumberStart??1)-pages.length-1;}
  const prior=pages.at(-1)!,page=firstEditorPage(section.layout,canvasWidth);
  Object.assign(page,{top:prior.top+prior.height+gap,sectionId:section.id,sectionPage:pages.length-sectionStart,number:pages.length+1+numberOffset});
  pages.push(page);return page;
 };
 const advance=(unit:PageUnit,page:EditorPage)=>{
  const space=contentStart(page)-(unit.top+added);
  if(space>0){breaks.push({pos:unit.pos,height:space,kind:unit.kind,page:pages.length});added+=space;}
 };
 const linked=units.map(u=>u.bottom);
 for(let i=units.length-2;i>=0;i--)if(units[i].keepWithNext&&!units[i+1].force&&!units[i+1].section)linked[i]=linked[i+1];
 for(const [i,unit] of units.entries()){
  if(unit.section){
   const next=unit.section,previous=active,frame=pages.at(-1)!;
   const [w,h]=pageSize(next.layout).map(v=>v*4/3);
   const sameGeometry=Math.abs(frame.width-w)<.1&&Math.abs(frame.height-h)<.1&&['top','bottom','left','right'].every(k=>frame.layout[k as 'top']===next.layout[k as 'top']);
   active=next;
   if(next.kind!=='continuous'||!sameGeometry){
    const parity=next.kind==='oddPage'?1:next.kind==='evenPage'?0:undefined;
    if(parity!=null&&(pages.length+1)%2!==parity)append(previous);
    advance(unit,append(next,true));
   }else{
    sectionStart=pages.length-1;
    if(next.restartNumbering)numberOffset=(next.layout.pageNumberStart??1)-pages.length;
   }
   continue;
  }
  let page=pages.at(-1)!;
  if(unit.force){for(let n=0;n<Math.max(1,unit.force);n++)page=append();advance(unit,page);continue;}
  const y=unit.top+added,available=page.height-px(page.layout.top+page.layout.bottom);
  const bottom=linked[i]-unit.top<=available?linked[i]:unit.bottom;
  if(bottom+added>contentEnd(page)+.5&&y>contentStart(page)+.5){page=append();advance(unit,page);}
 }
 return {pages,breaks,count:pages.length};
}
