import test from 'node:test';
import assert from 'node:assert/strict';
import {initialDocumentPreview} from '@prisma/writer/document-preview';
import {defaultLayout,type FlowDocument} from '@prisma/writer/flow';
test('initial preview stops at the first page break and excludes original files and metadata',()=>{
 const flow={layout:defaultLayout(),content:{type:'doc',content:[{type:'paragraph',content:[{type:'text',text:'Primera página'}]},{type:'pageBreak'},{type:'paragraph',content:[{type:'text',text:'Página siguiente'}]}]},officeOriginal:{private:'original'},references:[{private:'references'}]} as unknown as FlowDocument;
 const preview=initialDocumentPreview(flow);
 assert.equal(preview.content.content?.length,1);assert.equal(preview.officeOriginal,undefined);assert.equal(preview.references,undefined);
 assert.equal(preview.content.content?.[0].content?.[0].text,'Primera página');
});
test('preview has bounded text, nodes and images for documents of any length',()=>{
 const preview=initialDocumentPreview({layout:defaultLayout(),content:{type:'doc',content:Array.from({length:10000},()=>({type:'paragraph',content:[{type:'text',text:'a'.repeat(1000)}]}))}});
 assert(preview.content.content!.length<=18);
 assert(JSON.stringify(preview).length<10000);
});
