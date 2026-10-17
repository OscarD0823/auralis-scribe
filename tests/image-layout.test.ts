import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {Document,Paragraph,ImageRun,Packer} from 'docx';
import {importDocx} from '../src/document-import-core';
import type {RichNode} from '../src/flow';
Object.assign(globalThis,{DOMParser});
test('DOCX illustration frames preserve height before deferred image decoding',async()=>{
  const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a+AAAAABJRU5ErkJggg==','base64');
  const bytes=await Packer.toBuffer(new Document({sections:[{children:[new Paragraph({children:[new ImageRun({type:'png',data:png,transformation:{width:100,height:200}})]})]}]}));
  const result=importDocx(bytes),images:RichNode[]=[];
  const walk=(node:RichNode)=>{if(node.type==='image')images.push(node);node.content?.forEach(walk);};walk(result.flow.content);
  assert.equal(images.length,1);
  assert.equal(images[0].attrs?.width,100);
  assert.equal(images[0].attrs?.height,200);
  assert.match(String(images[0].attrs?.src),/^data:image\/png;base64,/);
});
