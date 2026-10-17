import type {FlowDocument,RichNode} from './flow';

/** Bounded initial-page view. Never sends the original ZIP or private metadata. */
export function initialDocumentPreview(flow:FlowDocument):FlowDocument {
  let nodes=400,characters=6000,imageBytes=3_000_000;
  const copy=(node:RichNode):RichNode|undefined=>{
    if(--nodes<0||characters<=0)return;
    if(node.type==='image'){
      const src=String(node.attrs?.src??'');
      if(src.length>imageBytes)return;
      imageBytes-=src.length;
    }
    const text=node.text?.slice(0,characters);characters-=text?.length??0;
    return {...node,text,content:node.content?.map(copy).filter((n):n is RichNode=>!!n)};
  };
  const content:RichNode[]=[];
  for(const node of flow.content.content??[]){
    if(content.length>=18||node.type==='pageBreak'||(content.length>0&&node.attrs?.breakBefore))break;
    const block=copy(node);if(block)content.push(block);
  }
  return {layout:{...flow.layout},language:flow.language,content:{type:'doc',content}};
}
