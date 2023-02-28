/*!
 * Copyright (C) Microsoft Corporation. All rights reserved.
 */
/* istanbul ignore file */
/* eslint-disable */
/* eslint-disable prettier/prettier */
import monaco from 'monaco-editor/esm/vs/editor/editor.api';
import * as React from 'react';
import { IStyle, ITheme } from '@fluentui/react/lib/Styling';
import { IStyleFunctionOrObject } from '@fluentui/react/lib/Utilities';

import { SyntaxParameters } from './PowerFxSyntaxTypes';

export const reactFormulaEditorId: string = 'reactFormulaEditor';
export type FormulaCursorPosition = {
  cursorStartIndex: number;
  cursorEndIndex: number;
};

// Unfortunately the type definitions for monaco are old. Once the react formula bar is the only
// formula bar this can be removed.
export enum ErrorSeverity {
  Info = 1,
  Warning = 4,
  Error = 8
}

export interface FormulaEditorError {
  startIndex: number;
  endIndex: number;
  severity: ErrorSeverity;
  message: string;
  messageId?: string;
}

export interface FormulaEditorStyles {
  root: IStyle;
  editorBoundary: IStyle;
}

export type FormulaEditorStyleProps = Required<Pick<FormulaEditorProps, 'theme' | 'locked'>> &
  Pick<FormulaEditorProps, 'className'> &
  FormulaEditorState;

export interface FormulaEditorState {
  focused: boolean;
}

export interface FormulaEditorProps {
  /** The initial value for the editor. */
  defaultValue: string;
  /** Whether or not the editor should be locked and editing should be prevented. */
  locked?: boolean;

  /** When set, editor will lose focus when the enter key is pressed,
   * shift + enter can still be used to create multiple lines.
   */
  leaveOnEnterKey?: boolean;

  // Design props
  /** The width of the editor. Either a number (of pixels) or a CSS style string value. */
  width?: string | number;
  /** The height of the editor. Either a number (of pixels) or a CSS style string value. */
  height?: string | number;
  /** Styling overrides for the component */
  styles?: IStyleFunctionOrObject<FormulaEditorStyleProps, FormulaEditorStyles>;
  /** The current fabric theme object */
  theme?: ITheme;
  /** A classname to provide custom styling to the root element. */
  className?: string;
  /** Whether to give keyboard focus to the Monaco editor on mount. */
  focusOnMount?: boolean;
  /** An event emitted when the editor is focused. */
  syntaxParameters: SyntaxParameters;
  /** A function to provide completion items for the given position and document */
  completionProvider?: monaco.languages.CompletionItemProvider['provideCompletionItems'];
  /** A function to provide parameter hints for the given position and document */
  signatureHelpProvider?: monaco.languages.SignatureHelpProvider['provideSignatureHelp'];
  /** Monaco interface {monaco.editor.IEditorConstructionOptions}. */
  monacoEditorOptions?: monaco.editor.IEditorConstructionOptions;

  // Callback props
  /** An event emitted when the content of the current editor has changed. */
  onChange?: (newValue: string, e: monaco.editor.IModelContentChangedEvent) => void;
  /** An event emitted when the editor has mounted. */
  onEditorDidMount?: (monacoEditor: monaco.editor.ICodeEditor, monacoEnv: typeof monaco) => void;
  /** An event emitted when the editor loses focus. */
  onBlur?: (value: string) => void;
  /** An event emitted when the editor is focused. */
  onFocus?: (cursorPosition: FormulaCursorPosition, value: string) => void;
  onDidChangeCursorPosition?: (cursorPosition: FormulaCursorPosition, value: string) => void;
  onKeyDown?: (event: monaco.IKeyboardEvent) => void;
  onKeyUp?: (event: monaco.IKeyboardEvent) => void;
  onContainerKeyDown?: (event: React.KeyboardEvent<HTMLDivElement>) => void;
  isReadOnly?: boolean;
  key?: string;
}

export class PowerFxFormulaEditorError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PowerFxFormulaEditorError';
  }
}
