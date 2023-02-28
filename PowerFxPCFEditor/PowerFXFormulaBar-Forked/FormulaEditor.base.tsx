/*!
 * Copyright (C) Microsoft Corporation. All rights reserved.
 */
/* istanbul ignore file */
/* eslint-disable */
/* eslint-disable prettier/prettier */
//import monaco from 'monaco-editor/esm/vs/editor/editor.api';
//import monaco from 'monaco-editor';

import * as React from 'react';
import * as monaco from 'monaco-editor';
import loader from '@monaco-editor/loader';
import MonacoEditor, { Monaco } from '@monaco-editor/react';

import { classNamesFunction } from '@fluentui/react/lib/Utilities';

import {
  FormulaEditorProps,
  FormulaEditorState,
  FormulaEditorStyleProps,
  FormulaEditorStyles
} from './FormulaEditor.types';
import { addProvidersForModel, ensureLanguageRegistered, languageName } from './PowerFxSyntax';
import { ensureThemeSetup, themeName } from './PowerFxTheme';
export const editorFontFamily = "'Menlo', 'Consolas', monospace,sans-serif";
export const editorFontSize = 14;

/** The default option configuration that we use for all the Monaco code editors inside PowerFx. */
const defaultEditorOptions: monaco.editor.IEditorConstructionOptions & monaco.editor.IGlobalEditorOptions = {
  fontSize: editorFontSize,
  lineDecorationsWidth: 4,
  scrollbar: {
    vertical: 'auto',
    verticalScrollbarSize: 8,
    horizontal: 'auto',
    horizontalScrollbarSize: 8
  },
  // This fixes the first time render bug, and handles additional resizes.
  automaticLayout: true,
  contextmenu: false,
  // Don't show a border above and below the current line in the editor.
  renderLineHighlight: 'none',
  lineNumbers: 'off',
  wordWrap: 'on',
  autoClosingBrackets: 'never',
  quickSuggestions: true,
  scrollBeyondLastLine: false,
  // Don't show the minimap (the scaled down thumbnail view of the code)
  minimap: { enabled: false },
  selectionClipboard: false,
  // Don't add a margin on the left to render special editor symbols
  glyphMargin: false,
  revealHorizontalRightPadding: 24,
  find: {
    seedSearchStringFromSelection: 'never',
    autoFindInSelection: 'never'
  },
  suggestOnTriggerCharacters: true,
  codeLens: false,
  // Don't allow the user to collapse the curly brace sections
  folding: false,
  formatOnType: true,
  fontFamily: editorFontFamily,
  wordBasedSuggestions: false,
  // This option helps to fix some of the overflow issues when using the suggestion widget in grid rows
  // NOTE: This doesn't work when it's hosted inside Fluent Callout control
  // More details in https://github.com/microsoft/monaco-editor/issues/2503
  fixedOverflowWidgets: true
};

const getClassNames = classNamesFunction<FormulaEditorStyleProps, FormulaEditorStyles>();

export class FormulaEditorBase extends React.Component<FormulaEditorProps, FormulaEditorState> {
  private _model: monaco.editor.IModel | null = null;
  private _editor: monaco.editor.ICodeEditor | null = null;
  private _boundaryRef: React.RefObject<HTMLDivElement>;
  /** Manually resetting the editor text on context change calls onChange. We need to ignore the call to onChange in that case. */
  private _ignoreOnChange: boolean = false;
  private _subscriptions: monaco.IDisposable[] = [];
  public constructor(props: FormulaEditorProps) {
    super(props);
    this.state = { focused: false };
    this._boundaryRef = React.createRef<HTMLDivElement>();
    //Required for initial loading of monaco loader js which fails in custom page
    // Reference - https://github.com/suren-atoyan/monaco-loader
    loader.config({ monaco });
    loader.init().then((monaco) => {
      console.info("monaco instance initialized");
    }).catch((error) => { console.error(error) });
  }

  public render(): JSX.Element {
    const {
      defaultValue,
      width,
      height = 30,
      locked,
      className,
      styles,
      theme,
      monacoEditorOptions,
      key
    } = this.props;
    const { focused } = this.state;

    const classNames = getClassNames(styles, { className, theme: theme!, locked: !!locked, focused });
    return (
      <div
        className={classNames.root}
        onKeyDown={this._onContainerKeyDown}
        ref={this._boundaryRef}
        role="presentation"
        tabIndex={0}
      >
        <div className={classNames.editorBoundary}>
          {
            <MonacoEditor
              language={languageName}
              theme={themeName}
              width={width}
              height={height}
              options={{ ...defaultEditorOptions, ...monacoEditorOptions }}
              defaultValue={defaultValue}
              onMount={this._onEditorDidMount}
              key={key}
            />
          }
        </div>
      </div>
    );
  }

  public componentDidMount() {
    if (this._boundaryRef.current) {
      const monacoInputTextarea = this._boundaryRef.current.querySelector('textarea');
      if (monacoInputTextarea) {
        // The editor MUST be manually focused, or the user will end up in an
        // inescapable element in keyboard navigation.
        monacoInputTextarea.tabIndex = -1;
      }
    }
  }

  public componentDidUpdate(prevProps: FormulaEditorProps) {
    // Determine whether the editor value needs to be reset.
    const lockedStateHasChanged = prevProps.locked !== this.props.locked;

    if (lockedStateHasChanged && this._editor) {
      this._editor.updateOptions({
        readOnly: this.props.locked
      });
    }
    this._editor?.updateOptions({
      readOnly: this.props.isReadOnly
    });
  }

  public componentWillUnmount() {
    this._disposeEditorSubscriptions();
  }

  private _getCurrentCursorPosition = () => {
    const editor = this._editor!;
    const selection = editor.getSelection()!;
    const startPosition = selection.getStartPosition();
    const endPosition = selection.getEndPosition();

    const cursorStartIndex = this._model!.getOffsetAt(startPosition);
    const cursorEndIndex = this._model!.getOffsetAt(endPosition);

    return {
      position: editor.getPosition(), // the start position of the cursor for use in intellisense
      cursorStartIndex,
      cursorEndIndex
    };
  };

  private _onEditorDidMount = (
    monacoEditor: monaco.editor.IStandaloneCodeEditor,
    monacoParam: typeof monaco
  ) => {
    this._model = monacoParam.editor.createModel(this.props.defaultValue);
    monacoEditor.setModel(this._model);
    this._editor = monacoEditor;

    ensureLanguageRegistered(monacoParam, this.props.syntaxParameters);
    ensureThemeSetup(monacoParam);
    monacoParam.editor.setModelLanguage(this._editor.getModel()!, languageName);
    addProvidersForModel(this._model, this._provideCompletionItems, this._provideSignatureHelp);

    this._disposeEditorSubscriptions();

    this._subscriptions.push(monacoEditor.onDidBlurEditorWidget(this._onBlur));
    this._subscriptions.push(monacoEditor.onDidFocusEditorWidget(this._onFocus));
    this._subscriptions.push(monacoEditor.onDidChangeCursorPosition(this._onDidChangeCursorPosition));
    this._subscriptions.push(monacoEditor.onKeyDown(this._onKeyDown));
    this._subscriptions.push(monacoEditor.onKeyUp(this._onKeyUp));
    this._subscriptions.push(this._model.onDidChangeContent(this._onChange));

    if (this.props.focusOnMount) {
      monacoEditor.focus();
      // By default the cursor will move to the beginning. This looks bad and is annoying, so move it to the end.
      this._moveCursorToEnd(this._editor, this._model);
    }
    if (this.props.onEditorDidMount) {
      this.props.onEditorDidMount(monacoEditor, monacoParam);
    }
  };

  /**
   * Moves the cursor to the end of the content.
   */
  private _moveCursorToEnd(monacoEditor: monaco.editor.ICodeEditor, model: monaco.editor.ITextModel): void {
    const numberOfLines = model.getLineCount();
    const numberOfColumnsOnLastLine = model.getLineMaxColumn(numberOfLines);
    monacoEditor.setPosition({
      column: numberOfColumnsOnLastLine,
      lineNumber: numberOfLines
    });
  }

  private _onChange = (e: monaco.editor.IModelContentChangedEvent) => {
    if (this._ignoreOnChange || !this.props.onChange || !this._model) {
      return;
    }

    this.props.onChange(this._model.getValue(), e);
  };

  private _onContainerKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    if (
      (event.key === 'Enter' || event.key === ' ') &&
      this._editor &&
      event.target === this._boundaryRef.current
    ) {
      // When focus is on the boundary allow users enter the editor region.
      this._editor.focus();
      event.preventDefault();
      event.stopPropagation();
    }

    if (event.key === 'ArrowUp' && event.target !== this._boundaryRef.current) {
      // This is being done because of a conflict when the control is used in
      // Fabric Focus Zones that use arrow keys to move focus. If the editor
      // has focus we don't want fabric handling it.
      event.stopPropagation();
    }
    return true;
  };

  private _provideCompletionItems: monaco.languages.CompletionItemProvider['provideCompletionItems'] = async (
    ...rest
  ) => {
    if (!this.props.completionProvider) {
      return {
        incomplete: false,
        suggestions: []
      };
    }

    return this.props.completionProvider(...rest);
  };

  private _provideSignatureHelp: monaco.languages.SignatureHelpProvider['provideSignatureHelp'] = async (
    ...rest
  ) => {
    // todo: test dispose
    if (!this.props.signatureHelpProvider) {
      return {
        value: {
          signatures: [],
          activeSignature: 0,
          activeParameter: 0
        },
        dispose: () => {
          return;
        }
      };
    }

    return this.props.signatureHelpProvider(...rest);
  };

  private _onBlur = () => {
    this.setState({ focused: false });
    if (this.props.onBlur) {
      let blurValue = '';
      if (this._editor) {
        blurValue = this._editor.getValue();
      }
      this.props.onBlur(blurValue);
    }
  };

  private _onFocus = () => {
    this.setState({ focused: true });
    if (this.props.onFocus) {
      let focusValue = '';
      if (this._editor) {
        focusValue = this._editor.getValue();
      }
      this.props.onFocus(this._getCurrentCursorPosition(), focusValue);
    }
  };

  private _onKeyDown = (event: monaco.IKeyboardEvent) => {
    if (
      this._boundaryRef.current &&
      (event.code === 'Escape' || (this.props.leaveOnEnterKey && event.code === 'Enter' && !event.shiftKey))
    ) {
      this._boundaryRef.current.focus();
      event.preventDefault();
      event.stopPropagation();
      return;
    }

    if (this.props.onKeyDown) {
      this.props.onKeyDown(event);
    }
  };

  private _onKeyUp = (event: monaco.IKeyboardEvent) => {
    if (this.props.onKeyUp) {
      this.props.onKeyUp(event);
    }
  };

  private _onDidChangeCursorPosition = () => {
    // Provide current position and value to the subscriber in order to guarantee
    // that the value they use is valid for the current position if the user is actively typing
    if (this.props.onDidChangeCursorPosition) {
      this.props.onDidChangeCursorPosition(this._getCurrentCursorPosition(), this._editor!.getValue());
    }
  };

  private _disposeEditorSubscriptions = () => {
    this._subscriptions.forEach(element => element.dispose());
    this._subscriptions = [];
  };
}
