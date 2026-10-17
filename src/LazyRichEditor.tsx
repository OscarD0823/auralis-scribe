import { lazy, Suspense, type ComponentProps } from "react";
import {DocumentLoading} from '@prisma/core/DocumentLoading';
const LoadedEditor = lazy(() =>
  import("./RichEditor").then((module) => ({ default: module.RichEditor })),
);
export type { RichEditorInstance } from "./RichEditor";
export function RichEditor(props: ComponentProps<typeof LoadedEditor>) {
  return (
    <Suspense
      fallback={
        <DocumentLoading label="Preparando escritura…" compact/>
      }
    >
      <LoadedEditor {...props} />
    </Suspense>
  );
}
