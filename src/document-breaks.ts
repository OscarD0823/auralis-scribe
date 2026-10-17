import type {Editor} from '@tiptap/core';
import {Selection, type EditorState} from '@tiptap/pm/state';

/** A block break splits the paragraph at the caret, preserving both halves. */
export function breakTransaction(state:EditorState, type:'pageBreak'|'columnBreak'|'sectionBreak', attrs:Record<string,unknown>={}) {
  // Sections are document-level boundaries. Inside a table/list, place them
  // after the containing object rather than hiding one inside a cell/item.
  const tr=state.selection.$from.depth>1
    ? state.tr.insert(state.selection.$from.after(1),state.schema.nodes[type].create(attrs))
    : state.tr.replaceSelectionWith(state.schema.nodes[type].create(attrs));
  if (!tr.doc.lastChild?.isTextblock)
    tr.insert(tr.doc.content.size,state.schema.nodes.paragraph.create());
  const next=state.selection.$from.depth>1?state.selection.$from.after(1)+1:tr.selection.to;
  tr.setSelection(Selection.near(tr.doc.resolve(next),1));
  return tr.scrollIntoView();
}
export function insertDocumentBreak(editor:Editor, type:'pageBreak'|'columnBreak'|'sectionBreak', attrs:Record<string,unknown>={}) {
  editor.view.dispatch(breakTransaction(editor.state,type,attrs));
  editor.commands.focus();
}

export type PageStart={pos:number;page:number};
/** Insert at the measured page boundary, including boundaries inside a paragraph. */
export function blankPageTransaction(state:EditorState, starts:PageStart[], count:number, placement:string) {
  const current=starts.filter(s=>s.pos<=state.selection.from).at(-1)?.page??1;
  const target=placement==='above'?current:current+1;
  const boundary=starts.find(s=>s.page>=target);
  const end=placement==='end'||!boundary||target>count;
  let pos=end?state.doc.content.size:boundary.pos;
  const nested=state.doc.resolve(pos);
  if(nested.depth>1)pos=nested.before(1);
  const paragraph=state.schema.nodes.paragraph.create({id:crypto.randomUUID(),spaceAfter:0});
  const pageBreak=()=>state.schema.nodes.pageBreak.create();
  const $pos=state.doc.resolve(pos);
  const alreadyBroken=$pos.nodeBefore?.type.name==='pageBreak'||
    ($pos.nodeBefore?.type.name==='sectionBreak'&&$pos.nodeBefore.attrs.kind!=='continuous');
  const followingBreak=$pos.nodeAfter?.attrs.breakBefore===true;
  const nodes=pos===0?[paragraph,pageBreak()]:end?[pageBreak(),paragraph]:alreadyBroken?[paragraph,pageBreak()]:followingBreak?[pageBreak(),paragraph]:[pageBreak(),paragraph,pageBreak()];
  const tr=state.tr.insert(pos,nodes);
  tr.setSelection(Selection.near(tr.doc.resolve(pos+(nodes[0].type.name==='paragraph'?1:nodes[0].nodeSize+1)),1));
  return tr.scrollIntoView();
}
