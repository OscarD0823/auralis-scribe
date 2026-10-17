import type {Editor} from '@tiptap/core';
import {TableMap} from '@tiptap/pm/tables';
import type {RichNode} from './flow';

export function repeatsTableHeader(row:RichNode):boolean {
 return row.attrs?.repeatHeader==null
  ?!!row.content?.length&&row.content.every(cell=>cell.type==='tableHeader')
  :row.attrs.repeatHeader===true;
}
/** Only a contiguous prefix repeats. A merged cell must stay within that prefix. */
export function repeatedTableHeaderCount(table:RichNode):number {
 const rows=table.content??[];let count=0;
 while(count<rows.length&&repeatsTableHeader(rows[count]))count++;
 for(let i=0;i<count;i++)if(rows[i].content?.some(cell=>i+(Number(cell.attrs?.rowspan)||1)>count))return 0;
 return count;
}
export function selectedTable(editor:Editor) {
 const {selection}=editor.state;
 const node=editor.state.doc.nodeAt(selection.from);
 if(node?.type.name==='table')return {node,pos:selection.from};
 for(let d=selection.$from.depth;d>0;d--)if(selection.$from.node(d).type.name==='table')return {node:selection.$from.node(d),pos:selection.$from.before(d)};
 throw new Error('Coloca el cursor en una celda de la tabla.');
}
export function setTablePagination(editor:Editor,repeatRows:number,cantSplit?:boolean) {
 const {node,pos}=selectedTable(editor);
 if(!Number.isInteger(repeatRows)||repeatRows<0||repeatRows>node.childCount)throw new Error('El número de filas de encabezado no es válido.');
 const tr=editor.state.tr;
 node.forEach((row,offset,i)=>tr.setNodeMarkup(pos+1+offset,undefined,{...row.attrs,repeatHeader:i<repeatRows,...(cantSplit==null?{}:{cantSplit})}));
 editor.view.dispatch(tr);editor.commands.focus();
}
export function distributeTableColumns(editor:Editor,availableWidth?:number) {
 const {node,pos}=selectedTable(editor),columns=TableMap.get(node).width;
 if(!columns)return;
 const dom=editor.view.nodeDOM(pos) as HTMLElement|null;
 const width=Math.max(24,Math.round((availableWidth??dom?.clientWidth??600)/columns));
 const tr=editor.state.tr;
 node.forEach((row,offset)=>row.forEach((cell,cellOffset)=>{
  tr.setNodeMarkup(pos+offset+cellOffset+2,undefined,{...cell.attrs,colwidth:Array(Math.max(1,Number(cell.attrs.colspan)||1)).fill(width)});
 }));
 editor.view.dispatch(tr);editor.commands.focus();
}
