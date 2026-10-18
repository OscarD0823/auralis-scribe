import test from 'node:test';import assert from 'node:assert/strict';import {DOMParser} from '@xmldom/xmldom';import {zipSync,strToU8} from 'fflate';
import {importDocx} from '../src/document-import-core';import {plainText} from '../src/flow';
Object.assign(globalThis,{DOMParser});
const w='http://schemas.openxmlformats.org/wordprocessingml/2006/main',mc='http://schemas.openxmlformats.org/markup-compatibility/2006';
const load=(body:string)=>importDocx(zipSync({'word/document.xml':strToU8(`<w:document xmlns:w="${w}" xmlns:mc="${mc}"><w:body>${body}<w:sectPr/></w:body></w:document>`)}));
test('inline content controls preserve their displayed values without claiming interactive controls',()=>{
 const result=load('<w:p><w:r><w:t>Cliente </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Nombre"/></w:sdtPr><w:sdtContent><w:r><w:t>Nombre de prueba</w:t></w:r></w:sdtContent></w:sdt></w:p>');
 assert(plainText(result.flow.content).includes('Cliente Nombre de prueba'));assert(result.warnings.some(w=>w.includes('controles de contenido')&&w.includes('no son interactivos')));
});
test('a Word text box in AlternateContent is read once instead of disappearing or duplicating its fallback',()=>{
 const box='<w:r><w:pict><w:txbxContent><w:p><w:r><w:t>CUADRO_UNICO</w:t></w:r></w:p><w:p><w:r><w:t>Segunda línea</w:t></w:r></w:p></w:txbxContent></w:pict></w:r>';
 const result=load(`<w:p><mc:AlternateContent><mc:Choice Requires="wps">${box}</mc:Choice><mc:Fallback>${box}</mc:Fallback></mc:AlternateContent></w:p>`),text=plainText(result.flow.content);
 assert.equal(text.split('CUADRO_UNICO').length-1,1);assert(text.includes('Segunda línea'));assert(result.warnings.some(w=>w.includes('cuadros de texto')&&w.includes('flujo principal')));
});
test('native math, charts and dynamic fields report reconstruction losses explicitly',()=>{
 const result=load('<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath><w:r><w:drawing><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/></w:drawing></w:r><w:fldSimple w:instr="MERGEFIELD Nombre"><w:r><w:t>Persona</w:t></w:r></w:fldSimple></w:p>');
 assert(result.warnings.some(w=>w.includes('1 ecuación nativa')));assert(result.warnings.some(w=>w.includes('1 gráfico nativo')));assert(result.warnings.some(w=>w.includes('campos dinámicos')));assert(plainText(result.flow.content).includes('Persona'));
});

test('AlternateContent uses the supported embedded image fallback once',()=>{
 const image='<w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rImage"/></w:drawing></w:r>';
 const result=importDocx(zipSync({
  'word/document.xml':strToU8(`<w:document xmlns:w="${w}" xmlns:mc="${mc}"><w:body><w:p><mc:AlternateContent><mc:Choice Requires="unsupported"><w:r><w:t>UNSUPPORTED_BRANCH</w:t></w:r></mc:Choice><mc:Fallback>${image}</mc:Fallback></mc:AlternateContent></w:p><w:sectPr/></w:body></w:document>`),
  'word/_rels/document.xml.rels':strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fallback.png"/></Relationships>'),
  'word/media/fallback.png':new Uint8Array([137,80,78,71]),
 }));
 const images:any[]=[];const walk=(node:any)=>{if(node.type==='image')images.push(node);node.content?.forEach(walk);};walk(result.flow.content);
 assert.equal(images.length,1);assert(!plainText(result.flow.content).includes('UNSUPPORTED_BRANCH'));assert(images[0].attrs.src.startsWith('data:image/png;base64,'));
});
