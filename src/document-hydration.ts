import {Fragment,type Node} from '@tiptap/pm/model';
import type {Transaction} from '@tiptap/pm/state';
/** Replace editor-generated trailing placeholders instead of importing them. */
export function appendDocumentBatch(transaction:Transaction,nodes:Node[],position:number){
  const fragment=Fragment.fromArray(nodes);
  return {transaction:transaction.replaceWith(position,transaction.doc.content.size,fragment).setMeta('addToHistory',false).setMeta('preventUpdate',true).setMeta('auralisHydrating',true),position:position+fragment.size};
}
