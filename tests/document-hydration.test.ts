import test from 'node:test';
import assert from 'node:assert/strict';
import {Schema} from '@tiptap/pm/model';
import {EditorState} from '@tiptap/pm/state';
import {appendDocumentBatch} from '@prisma/writer/document-hydration';
test('loading batches removes temporary trailing paragraphs without losing original text or breaks',()=>{
 const schema=new Schema({nodes:{doc:{content:'block+'},paragraph:{group:'block',content:'text*'},pageBreak:{group:'block',atom:true},text:{}}});
 const p=(text:string)=>schema.node('paragraph',null,text?schema.text(text):undefined),br=()=>schema.node('pageBreak');
 const nodes=Array.from({length:1000},(_,i)=>i%3===1?br():p('Contenido '+i));
 let state=EditorState.create({schema,doc:schema.node('doc',null,nodes.slice(0,2))}),position=nodes[0].nodeSize+nodes[1].nodeSize;
 for(let offset=2;offset<nodes.length;offset+=23){
  // Simulate the editor's trailing-node plugin after an intermediate page break.
  state=state.apply(state.tr.insert(state.doc.content.size,p('')));
  const batch=appendDocumentBatch(state.tr,nodes.slice(offset,offset+23),position);
  assert.equal(batch.transaction.getMeta('preventUpdate'),true);
  assert.equal(batch.transaction.getMeta('addToHistory'),false);
  state=state.apply(batch.transaction);position=batch.position;
 }
 assert.deepEqual(state.doc.toJSON(),schema.node('doc',null,nodes).toJSON());
});
