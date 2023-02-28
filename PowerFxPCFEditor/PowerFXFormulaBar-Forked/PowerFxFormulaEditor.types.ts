/*!
 * Copyright (C) Microsoft Corporation. All rights reserved.
 */
/* istanbul ignore file */
/* eslint-disable */
/* eslint-disable prettier/prettier */
import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';
import { IIconStyles } from '@fluentui/react/lib/Icon';
import { IStyle, ITheme } from '@fluentui/react/lib/Styling';
import { IStyleFunctionOrObject } from '@fluentui/react/lib/Utilities';

export const oneLineHeight: number = 19;

export interface PowerFxFormulaEditorStyles {
  root: IStyle;
  formulaEditor: IStyle;
  subComponentStyles?: PowerFxFormulaEditorSubComponentStyles;
}

export type PowerFxFormulaEditorSubComponentStyles = {
  expandEditorIcon: IStyleFunctionOrObject<{}, IIconStyles>;
};

export type PowerFxFormulaEditorStyleProps = Required<
  Pick<PowerFxFormulaEditorProps, 'theme' | 'showExpandEditorButton'>
>;

export interface PowerFxFormulaEditorProps {
  /** get the document uri */
  getDocumentUriAsync: () => Promise<string>;

  /** The initial value for the editor. */
  defaultValue: string;

  /** Send and receive messages to and from language server */
  messageProcessor: MessageProcessor;

  /** Monaco interface {monaco.editor.IEditorConstructionOptions}. */
  monacoEditorOptions?: monaco.editor.IEditorConstructionOptions;

  /** Styling overrides for the component */
  styles?: IStyleFunctionOrObject<PowerFxFormulaEditorStyleProps, PowerFxFormulaEditorStyles>;

  /** The current fabric theme object */
  theme?: ITheme;

  /** An event emitted when the content of the current editor has changed. */
  onChange?: (newValue: string) => void;

  /** An event emitted when the editor loses focus. */
  onBlur?: (value: string) => void;
  onKeyDown?: (event: monaco.IKeyboardEvent) => void;
  onKeyUp?: (event: monaco.IKeyboardEvent) => void;

  /** An event emitted when the editor has mounted. */
  onEditorDidMount?: (monacoEditor: monaco.editor.ICodeEditor, monacoEnv: typeof monaco) => void;

  /** The minimum number of text lines to be rendered in the TextArea */
  minLineCount: number;

  /** The maximum number of text lines to be rendered in the TextArea */
  maxLineCount: number;

  /**
   * Show chevron icon button to expand/collapse formula bar editor
   * Note: If you wish the expanded editor to be shown on top of other UI instead of pushing them down, you can set the
   * host with fixed height, larger z-index and overflow: true
   */
  showExpandEditorButton?: boolean;

  /** Optional language server protocol configuration */
  lspConfig?: {
    disableDidOpenNotification?: boolean;
    disableDidChangeNotification?: boolean;
    disableCompletionRequest?: boolean;
    enableSignatureHelpRequest?: boolean;
  };

  /** Whether to give keyboard focus to the editor on mount. */
  editorFocusOnMount?: boolean;

  /** Aria label for the expand editor button */
  expandEditorButtonAriaLabel?: string;

  /** This function is called when a percieved error is throw in the component */
  onError?: (error: Error) => void;

  width?: number;

  key?: string;

  isReadOnly?: boolean;
}

export interface IDisposable {
  dispose(): void;
}

export interface MessageProcessor {
  /**
   * Add the callback method which will be invoked whenever any new message is received from language server.
   *
   * @param listener the listener callback method
   * @returns the IDisposable object that can dispose the listener
   */
  addListener: (listener: (payload: string) => void) => IDisposable;

  /**
   * Send payload to the language server.
   */
  sendAsync: (payload: string) => Promise<void>;
}
