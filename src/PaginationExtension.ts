import { Extension } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import type { EditorView } from "@tiptap/pm/view";
import { mm, pageSize, type PageLayout } from "./flow";
import {type PageGap} from "./pagination";
import {editorSectionBlocks,planSectionPages,type EditorPage,type SectionPageUnit} from './section-pagination';
import {yieldToDevice} from '@prisma/core/task-yield';
import type {Node as DocumentNode} from '@tiptap/pm/model';
import {repeatedTableHeaderCount} from './table-tools';
const key = new PluginKey<DecorationSet>("prismaPagination");
function sectionDecorations(doc:DocumentNode,initial?:PageLayout) {
 if(!initial)return [];
 const blocks=editorSectionBlocks(doc,initial);
 const canvas=Math.max(...blocks.map(b=>pageSize(b.section.layout)[0]*4/3),pageSize(initial)[0]*4/3);
 return blocks.map(({node,pos,section})=>{
  const pageWidth=pageSize(section.layout)[0]*4/3;
  const body=pageWidth-mm(section.layout.left+section.layout.right)*4/3;
  const indent=Math.max(0,Number(node.attrs.indent)||0)*24;
  const width=node.type.name==='image'?Math.min(body,Number(node.attrs.width)||body):Math.max(24,body-indent);
  const align=node.type.name==='image'?(node.attrs.align==='right'?body-width:node.attrs.align==='center'?(body-width)/2:0):0;
  const left=(canvas-pageWidth)/2+mm(section.layout.left)*4/3+align;
  return Decoration.node(pos,pos+node.nodeSize,{'data-editor-section':section.id,style:`position:relative;left:${left}px;width:${width}px;max-width:${body}px;box-sizing:border-box;`});
 });
}
export const Pagination = Extension.create<{
  layout?: PageLayout;
  onPages?: (count: number) => void;
  onGeometry?: (pages:EditorPage[])=>void;
}>({
  name: "pagination",
  addOptions: () => ({}),
  addStorage() {
    return { ...this.options, pageStarts:[{pos:0,page:1}], pageCount:1 };
  },
  addProseMirrorPlugins() {
    const settings = this.storage;
    let currentView:EditorView|undefined;
    return [
      new Plugin({
        key,
        state: {
          init: (_,state) => DecorationSet.create(state.doc,sectionDecorations(state.doc,settings.layout)),
          apply(tr, previous) {
            if(!settings.layout)return DecorationSet.empty;
            const plan = tr.getMeta(key) as PageGap[] | undefined;
            return plan
              ? DecorationSet.create(
                  tr.doc,
                  [...sectionDecorations(tr.doc,settings.layout),...plan.map((item) =>
                    Decoration.widget(
                      item.pos,
                      () => {
                        const dom = document.createElement(
                          item.kind === "table" ? "tr" : "span",
                        );
                        dom.className = "editor-page-gap";
                        dom.contentEditable = "false";
                        dom.setAttribute("aria-hidden", "true");
                        if (item.kind === "table") {
                          const cell = document.createElement("td");
                          cell.colSpan = 100;
                          cell.style.height = item.height + "px";
                          if(item.header){
                            cell.colSpan=item.header.columns;
                            const source=currentView?.nodeDOM(item.header.tablePos) as HTMLElement|null;
                            const table=source?.matches('table')?source as HTMLTableElement:source?.querySelector('table');
                            if(table){
                              const holder=document.createElement('div');holder.className='editor-repeated-header';holder.contentEditable='false';
                              const clone=table.cloneNode(false) as HTMLTableElement,body=document.createElement('tbody');
                              const columns=table.querySelector('colgroup');if(columns)clone.append(columns.cloneNode(true));
                              for(const row of Array.from(table.tBodies[0]?.rows??[]).slice(0,item.header.rows))body.append(row.cloneNode(true));
                              clone.append(body);clone.removeAttribute('id');clone.removeAttribute('data-block-id');holder.append(clone);
                              for(const el of holder.querySelectorAll<HTMLElement>('[id],a,button,input,textarea,select,[tabindex]')){el.removeAttribute('id');el.setAttribute('tabindex','-1');}
                              for(const el of holder.querySelectorAll('[data-block-id],[data-index-page]')){el.removeAttribute('data-block-id');el.removeAttribute('data-index-page');}
                              for(const el of holder.querySelectorAll('.selectedCell'))el.classList.remove('selectedCell');
                              cell.append(holder);
                            }
                          }
                          dom.append(cell);
                        } else dom.style.height = item.height + "px";
                        return dom;
                      },
                      {
                        side: -1,
                        key: `${item.pos}:${item.height.toFixed(2)}`,
                        ignoreSelection: true,
                        pageGap:true,
                      },
                    ),
                  )],
                )
              : tr.docChanged||tr.getMeta('paginationLayout')
                ? DecorationSet.create(tr.doc,[...sectionDecorations(tr.doc,settings.layout),...previous.map(tr.mapping,tr.doc).find().filter(d=>d.spec.pageGap)])
                : previous.map(tr.mapping, tr.doc);
          },
        },
        props: {
          decorations(state) {
            return key.getState(state);
          },
        },
        view(view) {
          currentView=view;
          let timer: ReturnType<typeof setTimeout>,
            running = false,
            destroyed = false,
            generation=0,
            previousLayout = settings.layout;
          let wasPreparing=false;
          async function measure(view: EditorView) {
            if (
              destroyed ||
              running ||
              !settings.layout ||
              !view.dom.isConnected
              ||view.dom.getAttribute('aria-busy')==='true'
            )
              return;
            if (view.composing) {
              schedule();
              return;
            }
            running = true;
            try {
              view.dispatch(
                view.state.tr.setMeta(key, []).setMeta("addToHistory", false),
              );
              const editorRect = view.dom.getBoundingClientRect(),
                scale =
                  editorRect.width / (view.dom.offsetWidth || editorRect.width),
                units: SectionPageUnit[] = [];
              const topOf = (rect: DOMRect) =>
                (rect.top - editorRect.top) / scale;
              const measuredDocument=view.state.doc, measuredLayout=settings.layout, measuredGeneration=generation;
              const blocks=editorSectionBlocks(measuredDocument,measuredLayout);
              let sliceStarted=performance.now();
              for (const {node,pos,section,nextSection} of blocks) {
                if(performance.now()-sliceStarted>12){
                  await yieldToDevice();
                  if(destroyed)return;
                  if(view.state.doc!==measuredDocument||settings.layout!==measuredLayout||generation!==measuredGeneration){schedule();return;}
                  sliceStarted=performance.now();
                }
                const dom = view.nodeDOM(pos) as HTMLElement | null;
                if (!dom?.getBoundingClientRect) continue;
                const rect = dom.getBoundingClientRect();
                if (pos > 0 && node.attrs.breakBefore)
                  units.push({
                    pos,
                    top: topOf(rect),
                    bottom: topOf(rect),
                    force: 1,
                  });
                if (
                  node.type.name === "pageBreak" ||
                  node.type.name === "sectionBreak"
                ) {
                  units.push({
                    pos: pos + node.nodeSize,
                    top: topOf(rect) + rect.height / scale,
                    bottom: topOf(rect) + rect.height / scale,
                    force: 1,
                    ...(nextSection?{section:nextSection}:{}),
                  });
                  continue;
                }
                if (node.type.name === "table") {
                  const table=dom.matches('table')?dom as HTMLTableElement:dom.querySelector('table');
                  const rows = Array.from(table?.tBodies[0]?.rows??[]);
                  const shell={type:'table',content:Array.from({length:node.childCount},(_,i)=>{const row=node.child(i);return {type:'tableRow',attrs:row.attrs,content:Array.from({length:row.childCount},(_,j)=>({type:row.child(j).type.name,attrs:row.child(j).attrs}))};})};
                  const headerCount=repeatedTableHeaderCount(shell);
                  const headerHeight=rows.slice(0,headerCount).reduce((sum,row)=>sum+row.getBoundingClientRect().height/scale,0);
                  const columns=Math.max(1,...shell.content.map(row=>row.content.reduce((sum,c)=>sum+(Number(c.attrs.colspan)||1),0)));
                  node.forEach((row, offset, i) => {
                    const box = rows[i]?.getBoundingClientRect();
                    if (box)
                      units.push({
                        pos: pos + 1 + offset,
                        top: topOf(box),
                        bottom: topOf(box) + box.height / scale,
                        kind: "table",
                        keepWithNext:i<headerCount&&i<node.childCount-1,
                        ...(i>=headerCount&&headerCount?{repeatHeader:{tablePos:pos,rows:headerCount,columns,height:headerHeight}}:{}),
                      });
                  });
                  continue;
                }
                if (
                  node.isTextblock ||
                  node.type.name === "blockquote" ||
                  node.type.name === "bulletList" ||
                  node.type.name === "orderedList"
                ) {
                  const lines = new Map<number, SectionPageUnit>();
                  const walker = document.createTreeWalker(
                    dom,
                    NodeFilter.SHOW_TEXT,
                  );
                  while (walker.nextNode()) {
                    const text = walker.currentNode as Text;
                    if (
                      text.parentElement?.closest(
                        "[contenteditable=false],.formatting-mark,.academic-note",
                      )
                    )
                      continue;
                    for (const word of text.data.matchAll(/\S+\s*/gu)) {
                      if(performance.now()-sliceStarted>12){
                        await yieldToDevice();
                        if(destroyed)return;
                        if(view.state.doc!==measuredDocument||settings.layout!==measuredLayout||generation!==measuredGeneration){schedule();return;}
                        sliceStarted=performance.now();
                      }
                      // Long tokens can wrap inside a word; measure their characters as well.
                      const offsets =
                        word[0].length > 40
                          ? Array.from(
                              { length: word[0].length },
                              (_, i) => word.index + i,
                            )
                          : [word.index];
                      for (const offset of offsets) {
                        const range = document.createRange();
                        range.setStart(text, offset);
                        range.setEnd(text, Math.min(text.length, offset + 1));
                        const box = range.getBoundingClientRect();
                        if (!box.height) continue;
                        const y = topOf(box),
                          line = Math.round(y * 2) / 2,
                          from = view.posAtDOM(text, offset);
                        const previous = lines.get(line);
                        if (!previous || from < previous.pos)
                          lines.set(line, {
                            pos: from,
                            top: y,
                            bottom: y + box.height / scale,
                          });
                      }
                    }
                  }
                  if (!lines.size)
                    units.push({
                      pos,
                      top: topOf(rect),
                      bottom: topOf(rect) + rect.height / scale,
                    });
                  else {
                    const measured=[...lines.values()].sort((a,b)=>a.top-b.top);
                    const keepLines=node.attrs.keepLines??node.type.name==='heading';
                    const keepNext=node.attrs.keepNext??node.type.name==='heading';
                    const available=(pageSize(section.layout)[1]-mm(section.layout.top+section.layout.bottom))*4/3;
                    if(keepLines&&measured.at(-1)!.bottom-measured[0].top<=available)
                      units.push({pos,top:measured[0].top,bottom:measured.at(-1)!.bottom,keepWithNext:!!keepNext});
                    else {
                      measured[measured.length-1].keepWithNext=!!keepNext;
                      units.push(...measured);
                    }
                  }
                } else
                  units.push({
                    pos,
                    top: topOf(rect),
                    bottom: topOf(rect) + rect.height / scale,
                  });
              }
              const canvas=Math.max(...blocks.map(b=>pageSize(b.section.layout)[0]*4/3),pageSize(measuredLayout)[0]*4/3);
              const plan = planSectionPages(units,measuredLayout,canvas);
              view.dispatch(
                view.state.tr
                  .setMeta(key, plan.breaks)
                  .setMeta("addToHistory", false),
              );
              settings.pageStarts=[{pos:0,page:1},...plan.breaks.map(b=>({pos:b.pos,page:b.page??1}))];
              settings.pageCount=plan.count;
              (settings as typeof settings & {measuredDoc?:DocumentNode}).measuredDoc=measuredDocument;
              const pageMap=new Map<string,number>();
              view.state.doc.descendants((node,pos)=>{
                if(node.attrs.id){
                  let low=0,high=settings.pageStarts.length;
                  while(low<high){const mid=(low+high)>>>1;if(settings.pageStarts[mid].pos<=pos)low=mid+1;else high=mid;}
                  const physical=settings.pageStarts[Math.max(0,low-1)].page;
                  pageMap.set(String(node.attrs.id),plan.pages[physical-1]?.number??physical);
                }
              });
              for(const cell of view.dom.querySelectorAll<HTMLElement>('[data-index-page]')){
                const page=pageMap.get(cell.dataset.indexPage??'');
                if(page)cell.textContent=String(page);
              }
              settings.onPages?.(plan.count);
              settings.onGeometry?.(plan.pages);
            } finally {
              running = false;
            }
          }
          function schedule() {
            generation++;
            clearTimeout(timer);
            timer = setTimeout(() => void measure(view), 120);
          }
          // Page gaps change the editor's height. Observing those changes would
          // repeatedly remove/rebuild every gap, especially costly on long docs.
          let measuredWidth=view.dom.getBoundingClientRect().width;
          const resize = new ResizeObserver(()=>{
            const width=view.dom.getBoundingClientRect().width;
            if(Math.abs(width-measuredWidth)>.5){measuredWidth=width;schedule();}
          });
          resize.observe(view.dom);
          const imageLoaded=(event:Event)=>{if(!(event.target as HTMLElement).closest?.('.editor-repeated-header'))schedule();};
          view.dom.addEventListener('load',imageLoaded,true);
          // Font loads can change line wraps without a document transaction.
          document.fonts?.addEventListener("loadingdone", schedule);
          schedule();
          return {
            update(next, old) {
              if(next.dom.getAttribute('aria-busy')==='true'){wasPreparing=true;clearTimeout(timer);return;}
              const changed = previousLayout !== settings.layout;
              previousLayout = settings.layout;
              if (!running && (next.state.doc !== old.doc || changed||wasPreparing))schedule();
              wasPreparing=false;
            },
            destroy() {
              destroyed = true;
              clearTimeout(timer);
              resize.disconnect();
              view.dom.removeEventListener('load',imageLoaded,true);
              document.fonts?.removeEventListener("loadingdone", schedule);
            },
          };
        },
      }),
    ];
  },
});
