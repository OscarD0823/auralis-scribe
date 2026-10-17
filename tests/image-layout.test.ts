import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {Document,Paragraph,ImageRun,Packer,TextRun,SimpleField,Bookmark,InternalHyperlink} from 'docx';
import {importDocx} from '../src/document-import-core';
import {plainText,type RichNode} from '../src/flow';
import {collectDocumentEntries} from '../src/academic';
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

test('caption and illustration bookmarks remain valid internal link targets after importing mixed runs',async()=>{
 const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a+AAAAABJRU5ErkJggg==','base64');
 const bytes=await Packer.toBuffer(new Document({sections:[{children:[
  new Paragraph({children:[new Bookmark({id:'caption',children:[new TextRun('Figura '),new SimpleField('SEQ Figura \\* ARABIC','1'),new TextRun(' Ejemplo')]})]}),
  new Paragraph({children:[new Bookmark({id:'drawing',children:[new ImageRun({type:'png',data:png,transformation:{width:120,height:80}})]}),new TextRun('Nota. Fuente de prueba.')]}),
  new Paragraph({children:[new InternalHyperlink({anchor:'caption',children:[new TextRun('Ir al título')]}),new InternalHyperlink({anchor:'drawing',children:[new TextRun('Ir al dibujo')]})]}),
 ]}]}));
 const nodes=importDocx(bytes).flow.content.content!;
 const links=nodes.at(-1)!.content!.map(n=>n.marks!.find(m=>m.type==='link')!.attrs!.href);
 assert.deepEqual(links,['#'+nodes[0].attrs!.id,'#'+nodes[1].attrs!.id]);
});

test('a figure and its note in the same Word paragraph retain order and stay in the generated figure index',async()=>{
 const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a+AAAAABJRU5ErkJggg==','base64');
 const bytes=await Packer.toBuffer(new Document({sections:[{children:[
  new Paragraph({children:[new TextRun('Figura '),new SimpleField('SEQ Figura \\* ARABIC','1'),new TextRun(' Diagrama de ejemplo')]}),
  new Paragraph({children:[new ImageRun({type:'png',data:png,transformation:{width:120,height:80}}),new TextRun({text:'Nota. Elaboración propia.',break:1})]}),
 ]}]}));
 const flow=importDocx(bytes).flow,entries=collectDocumentEntries(flow.content);
 assert.equal(entries.figures.length,1);
 assert.equal(flow.content.content![0].type,'caption');
 assert.equal(flow.content.content![1].type,'image');
 assert.equal(flow.content.content![0].attrs?.note,'Elaboración propia.');
 assert.equal(flow.content.content![0].attrs?.title,'Diagrama de ejemplo');
});

test('mixed Word runs preserve text before, between and after drawings without duplicating image nodes',async()=>{
 const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a+AAAAABJRU5ErkJggg==','base64');
 const drawing=()=>new ImageRun({type:'png',data:png,transformation:{width:50,height:40}});
 const bytes=await Packer.toBuffer(new Document({sections:[{children:[new Paragraph({children:[new TextRun('Antes'),drawing(),new TextRun('Entre'),drawing(),new TextRun('Después')]})]}]}));
 const nodes=importDocx(bytes).flow.content.content!;
 assert.deepEqual(nodes.map(n=>n.type),['paragraph','image','paragraph','image','paragraph']);
 assert.deepEqual(nodes.filter(n=>n.type==='paragraph').map(plainText),['Antes','Entre','Después']);
});
