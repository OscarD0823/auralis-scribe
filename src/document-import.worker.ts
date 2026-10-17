import {DOMParser} from '@xmldom/xmldom';
import {importDocx} from './document-import-core';
import {attachOfficeOriginal} from '@prisma/core/office-engines';
import {initialDocumentPreview} from './document-preview';

Object.assign(globalThis,{DOMParser});
self.onmessage=async({data})=>{
  try {
    const result=importDocx(data.bytes,data.available);
    self.postMessage({preview:initialDocumentPreview(result.flow)});
    await attachOfficeOriginal(result.flow,data.bytes);
    self.postMessage({result});
  } catch(error) {
    self.postMessage({error:error instanceof Error?error.message:'No se pudo leer el DOCX.'});
  }
};
