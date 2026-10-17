import {importDocx,embeddedDocxFonts} from './document-import-core';
import {deviceFonts,initializeDeviceFonts} from '@prisma/core/device-fonts';
import {validateLocalFonts,type LocalFont} from '@prisma/core/local-fonts';
import {attachOfficeOriginal} from '@prisma/core/office-engines';
import {initialDocumentPreview} from './document-preview';
import type {FlowDocument} from './flow';
export {importDocx,importText} from './document-import-core';
type Result=ReturnType<typeof importDocx>;
async function readDocument(bytes:Uint8Array,available:LocalFont[],onPreview?:(flow:FlowDocument)=>void):Promise<Result>{
 if(typeof Worker==='undefined'){
  const result=importDocx(bytes,available);onPreview?.(initialDocumentPreview(result.flow));await attachOfficeOriginal(result.flow,bytes);return result;
 }
 // One worker per import: terminate on success/error to release XML trees and ZIP buffers.
 // Creation errors are reported instead of repeating an expensive parse on the UI thread.
 const worker=new Worker(new URL('./document-import.worker.ts',import.meta.url),{type:'module'});
 return new Promise((resolve,reject)=>{
  const finish=()=>worker.terminate();
  worker.onmessage=({data})=>{
    if(data.preview){try{onPreview?.(data.preview);}catch(error){finish();reject(error);}return;}
    finish();data.error?reject(new Error(data.error)):resolve(data.result);
  };
  worker.onerror=()=>{finish();reject(new Error('No se pudo leer el documento en segundo plano. Intenta abrirlo de nuevo.'));};
  const copy=bytes.slice();
  try{worker.postMessage({bytes:copy,available},[copy.buffer]);}catch(error){finish();reject(error);}
 });
}
export async function importOfficeDocx(bytes:Uint8Array,onPreview?:(flow:FlowDocument)=>void){
 await initializeDeviceFonts();
 const embedded=await embeddedDocxFonts(bytes);
 const available=[...embedded.fonts,...deviceFonts().filter(face=>!embedded.fonts.some(f=>f.value===face.value&&f.style===face.style))];
 const result=await readDocument(bytes,available,onPreview);
 await validateLocalFonts(result.flow.fonts??[]);
 result.warnings.push(...embedded.warnings);
 result.warnings=result.warnings.filter(w=>!/historial de cambios|notas al pie o finales/.test(w));
 result.warnings.push('El original de DOCX se conserva. Los cambios de texto y notas mantienen sus partes; cambios de diseño o estructura requieren reconstrucción.');
 if(result.flow.importReport)result.flow.importReport.warnings=result.warnings;
 return result;
}
