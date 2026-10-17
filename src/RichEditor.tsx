import { useEffect, useRef, useState } from "react";
import {type Node as ModelNode} from '@tiptap/pm/model';
import {yieldToDevice} from '@prisma/core/task-yield';
import {appendDocumentBatch} from './document-hydration';
import { captureSelection, restoreSelection } from "./document-selection";
import { EditorContent, useEditor, type Editor } from "@tiptap/react";
import { Node, Extension } from "@tiptap/core";
import StarterKit from "@tiptap/starter-kit";
import { TextStyleKit } from "@tiptap/extension-text-style";
import TextAlign from "@tiptap/extension-text-align";
import Image from "@tiptap/extension-image";
import { TableKit } from "@tiptap/extension-table";
import Placeholder from "@tiptap/extension-placeholder";
import Superscript from "@tiptap/extension-superscript";
import Subscript from "@tiptap/extension-subscript";
import type { RichNode } from "./flow";
import type { PageLayout } from "./flow";
import { Pagination } from "./PaginationExtension";
import {
  Citation,
  Caption,
  Appendix,
  GeneratedIndex,
  ColumnBreak,
  SectionBreak,
  Annotation,
  Media,
  DocumentAttributes,
  FormattingMarks,
  SpellingMarks,
} from "./document-extensions";
export type RichEditorInstance = Editor;
const PageBreak = Node.create({
  name: "pageBreak",
  group: "block",
  atom: true,
  parseHTML: () => [{ tag: "div[data-page-break]" }],
  renderHTML: () => [
    "div",
    { "data-page-break": "true", class: "page-break" },
    "Salto de página",
  ],
});
const Indent = Extension.create({
  name: "indent",
  addGlobalAttributes() {
    return [
      {
        types: ["paragraph", "heading"],
        attributes: {
          indent: {
            default: 0,
            parseHTML: (el) => Number(el.dataset.indent) || 0,
            renderHTML: (attrs) => ({
              "data-indent": attrs.indent,
              style: `margin-left:${Number(attrs.indent) * 18}pt`,
            }),
          },
          lineHeight: {
            default: null,
            parseHTML: (el) => el.style.lineHeight,
            renderHTML: (attrs) =>
              attrs.lineHeight
                ? { style: `line-height:${attrs.lineHeight}` }
                : {},
          },
        },
      },
    ];
  },
});
export function RichEditor({
  content,
  onChange,
  onFocus,
  inline = false,
  style,
  autoFocus = false,
  showMarks = false,
  autoCorrect = false,
  positioned = false,
  searchQuery = "",
  pageLayout,
  onPages,
}: {
  content: RichNode;
  onChange: (node: RichNode) => void;
  onFocus: (editor: Editor) => void;
  inline?: boolean;
  style?: React.CSSProperties;
  autoFocus?: boolean;
  showMarks?: boolean;
  autoCorrect?: boolean;
  positioned?: boolean;
  searchQuery?: string;
  pageLayout?: PageLayout;
  onPages?: (count: number) => void;
}) {
  const lastExternal = useRef("");
  const hydrating=useRef(!inline&&(content.content?.length??0)>300);
  const [preparing,setPreparing]=useState(hydrating.current?80:0);
  const initial=useRef(hydrating.current?{...content,content:content.content?.slice(0,80)}:content);
  const editor = useEditor({
    immediatelyRender: false,
    extensions: [
      StarterKit.configure({
        link: { openOnClick: false, protocols: ["https", "http", "mailto"] },
      }),
      TextStyleKit,
      TextAlign.configure({ types: ["heading", "paragraph"] }),
      Image.configure({
        allowBase64: true,
        HTMLAttributes: {loading:'lazy',decoding:'async'},
        resize: {
          enabled: true,
          alwaysPreserveAspectRatio: true,
          minWidth: 24,
          minHeight: 24,
        },
      }),
      TableKit.configure({ table: { resizable: true } }),
      Placeholder.configure({
        placeholder: inline
          ? "Escribe aquí…"
          : "Empieza a escribir tu documento…",
      }),
      Superscript,
      Subscript,
      PageBreak,
      Indent,
      Citation,
      Caption,
      Appendix,
      GeneratedIndex,
      ColumnBreak,
      SectionBreak,
      Annotation,
      ...Media,
      DocumentAttributes,
      SpellingMarks,
      Pagination.configure({ layout: pageLayout, onPages }),
      FormattingMarks.configure({
        enabled: showMarks,
        autoCorrect,
        query: searchQuery,
      }),
    ],
    content:initial.current,
    editable:!hydrating.current,
    autofocus: autoFocus&&!hydrating.current ? (inline ? "end" : "start") : false,
    onCreate: ({ editor }) => {
      if (autoFocus && !editor.isDestroyed&&!hydrating.current) onFocus(editor);
    },
    onUpdate: ({ editor }) => {if(!hydrating.current)onChange(editor.getJSON());},
    onFocus: ({ editor }) => {if(!hydrating.current)onFocus(editor);},
    onSelectionUpdate: ({ editor }) => {if(!hydrating.current)onFocus(editor);},
    editorProps: {
      attributes: {
        class: inline ? "inline-prose" : "document-prose",
        "aria-label": inline
          ? "Editar texto del documento"
          : "Contenido del documento",
        spellcheck: "true",
        lang: "es",
      },
      transformPastedHTML: (html) => {
        const doc = new DOMParser().parseFromString(html, "text/html");
        for (const img of doc.querySelectorAll("img"))
          if (!/^data:image\/(png|jpeg);base64,/i.test(img.src)) img.remove();
        for (const el of doc.querySelectorAll("script,iframe,object,embed"))
          el.remove();
        return doc.body.innerHTML;
      },
    },
  });
  useEffect(() => {
    if (!editor || editor.isDestroyed) return;
    const document=editor.schema.nodeFromJSON(content);
    const incoming = JSON.stringify(document.toJSON());
    if (incoming === lastExternal.current) return;
    lastExternal.current = incoming;
    if (
      editor &&
      !editor.isDestroyed &&
      JSON.stringify(editor.getJSON()) !== incoming
    ) {
      const selection = captureSelection(editor.state.selection);
      if(!inline&&document.childCount>300){
        lastExternal.current='';
        let stopped=false;
        hydrating.current=true;
        editor.setEditable(false,false);
        editor.view.dom.setAttribute('aria-busy','true');
        const blocks:ModelNode[]=[];document.forEach(node=>blocks.push(node));
        const first=Math.min(80,blocks.length);
        let insertion=blocks.slice(0,first).reduce((sum,node)=>sum+node.nodeSize,0);
        editor.commands.setContent({...content,content:content.content?.slice(0,first)},{emitUpdate:false});
        setPreparing(first);
        void(async()=>{
          // Yield before every insertion so the first page can paint and the
          // loading animation remains responsive. Never emit a partial model.
          for(let offset=first;offset<blocks.length;offset+=80){
            await yieldToDevice();
            if(stopped||editor.isDestroyed)return;
            const batch=appendDocumentBatch(editor.state.tr,blocks.slice(offset,offset+80),insertion);
            editor.view.dispatch(batch.transaction);insertion=batch.position;
            setPreparing(Math.min(blocks.length,offset+80));
          }
          if(stopped||editor.isDestroyed)return;
          editor.view.dom.removeAttribute('aria-busy');
          editor.setEditable(true,false);
          lastExternal.current=incoming;
          hydrating.current=false;setPreparing(0);
          editor.view.dispatch(editor.state.tr.setSelection(restoreSelection(selection,editor.state.doc)).setMeta('addToHistory',false));
          if(autoFocus)onFocus(editor);
        })().catch(()=>{
          if(stopped||editor.isDestroyed)return;
          // Restore the complete source before enabling editing if a batch fails.
          editor.commands.setContent(content,{emitUpdate:false});
          lastExternal.current=incoming;hydrating.current=false;setPreparing(0);
          editor.view.dom.removeAttribute('aria-busy');editor.setEditable(true,false);
          if(autoFocus)onFocus(editor);
        });
        return()=>{stopped=true;};
      }
      editor.commands.setContent(content, { emitUpdate: false });
      hydrating.current=false;setPreparing(0);
      editor.view.dom.removeAttribute('aria-busy');editor.setEditable(true,false);
      editor.view.dispatch(
        editor.state.tr.setSelection(
          restoreSelection(selection, editor.state.doc),
        ),
      );
    }
  }, [content, editor]);
  useEffect(() => {
    if (!editor || editor.isDestroyed) return;
    const extension = editor.extensionManager.extensions.find(
      (e) => e.name === "formattingMarks",
    );
    if (extension) {
      Object.assign(
        (
          editor.storage as unknown as Record<
            string,
            { enabled: boolean; autoCorrect: boolean; query: string }
          >
        ).formattingMarks,
        { enabled: showMarks, autoCorrect, query: searchQuery },
      );
      editor.view.dispatch(editor.state.tr);
    }
  }, [editor, showMarks, autoCorrect, searchQuery]);
  useEffect(() => {
    if (!editor || editor.isDestroyed) return;
    const storage = editor.storage as unknown as {
      pagination: { layout?: PageLayout; onPages?: (count: number) => void };
    };
    Object.assign(storage.pagination, { layout: pageLayout, onPages });
    editor.view.dispatch(editor.state.tr.setMeta("paginationLayout", true));
    window.dispatchEvent(new Event("resize"));
  }, [editor, pageLayout, onPages]);
  return (
    <div
      className={
        positioned
          ? "pdf-document-editor"
          : inline
            ? "inline-editor"
            : "rich-editor"
      }
      style={style}
    >
      {preparing>0&&<div className="editor-preparation" role="status">Preparando el documento para editar · {preparing.toLocaleString()} de {(content.content?.length??0).toLocaleString()} elementos. Puedes ver la primera página; la edición se habilitará al terminar.</div>}
      <EditorContent editor={editor} />
    </div>
  );
}
