/*!
 * Copyright (C) Microsoft Corporation. All rights reserved.
 */
/* eslint-disable */
/* istanbul ignore file */
import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';
import * as React from 'react';
import {
  CompletionItemKind,
  CompletionTriggerKind,
  Diagnostic,
  DiagnosticSeverity,
  PublishDiagnosticsParams
} from 'vscode-languageserver-protocol';

import { IconButton } from '@fluentui/react/lib/Button';
import { classNamesFunction } from '@fluentui/react/lib/Utilities';

import { FormulaEditor } from './FormulaEditor';
import {
  IDisposable,
  oneLineHeight,
  PowerFxFormulaEditorProps,
  PowerFxFormulaEditorStyleProps,
  PowerFxFormulaEditorStyles
} from './PowerFxFormulaEditor.types';
import { PowerFxLanguageClient, PublishTokensParams, TokenResultType } from './PowerFxLanguageClient';
//import { HighlightedName, NameKind } from '@microsoft/power-fx-formulabar/src/PowerFxSyntaxTypes';
//import { PowerFxFormulaEditorError } from '@microsoft/power-fx-formulabar/src/FormulaEditor.types';

interface PowerFxFormulaEditorState {
  expanded: boolean;
  height: number;
}

export enum NameKind {
  HostSymbol,
  Variable,
  Function
}

/** Defines a name that should be highlighted in the editor. */
export interface HighlightedName {
  readonly name: string;
  readonly kind: NameKind;
}


const formulaEditorMinHeight: number = 30;
const minExpandedEditorHeight: number = 7 * oneLineHeight;

const getClassNames = classNamesFunction<
  PowerFxFormulaEditorStyleProps,
  Required<PowerFxFormulaEditorStyles>
>();

export class PowerFxFormulaEditor extends React.Component<
  PowerFxFormulaEditorProps,
  PowerFxFormulaEditorState
>
{
  public static defaultProps = {
    editorFocusOnMount: true
  };
  private _editor: monaco.editor.ICodeEditor | null = null;
  private _model: monaco.editor.IModel | null = null;
  private _monaco: typeof monaco | null = null;
  private _languageClient: PowerFxLanguageClient | null = null;
  private _listenerDisposable: IDisposable | null = null;
  private _version: number = 0;
  private _normalizedCompletionLookup: { [lowercase: string]: string } = {};
  private _onNamesChanged: (names: HighlightedName[]) => void = () => null;

  constructor(props: PowerFxFormulaEditorProps) {
    super(props);

    this.state = {
      expanded: false,
      height: this._getHeight()
    };

    this._listenerDisposable = this.props.messageProcessor.addListener((payload: string): void => {
      this._languageClient?.onDataReceivedFromLanguageServer(payload);
    });
  }

  public componentWillUnmount() {
    this._listenerDisposable?.dispose();
    this._listenerDisposable = null;
  }

  public render() {
    const {
      styles,
      theme,
      defaultValue,
      onBlur,
      onKeyDown,
      onKeyUp,
      monacoEditorOptions,
      showExpandEditorButton,
      expandEditorButtonAriaLabel,
      editorFocusOnMount,
      width,
      key,
      isReadOnly
    } = this.props;
    const { height, expanded } = this.state;
    const classNames = getClassNames(styles, {
      theme: theme!,
      showExpandEditorButton: !!showExpandEditorButton
    });
    const expandEditorIconStyles = classNames.subComponentStyles?.expandEditorIcon({});
    return (
      <div id="root" className={classNames.root}>
        <div id="formulaEditor" className={classNames.formulaEditor}>
          <FormulaEditor
            defaultValue={defaultValue}
            monacoEditorOptions={monacoEditorOptions}
            height={height}
            width={width}
            isReadOnly={isReadOnly}
            key={this.props.key}
            // Might be nice for this be true, but it interferes with the selection widget.
            leaveOnEnterKey={false}
            onEditorDidMount={this._editorDidMount}
            syntaxParameters={{
              highlightedNames: [],
              useSemicolons: false,
              subscribeToNames: namesChanged => {
                this._onNamesChanged = namesChanged;
              },
              getNormalizedCompletionLookup: () => this._normalizedCompletionLookup
            }}
            focusOnMount={editorFocusOnMount}
            onChange={this._onContentChanged}
            onBlur={onBlur}
            onKeyDown={onKeyDown}
            onKeyUp={onKeyUp}
            completionProvider={(
              model: monaco.editor.ITextModel,
              position: monaco.Position,
              context: monaco.languages.CompletionContext
            ) =>
              // Note: we call model methods up-front to avoid disposed model issue
              // since there are other async operations inside _provideCompletionAsync
              this._provideCompletionAsync(
                model.getValue(),
                position,
                model.getWordAtPosition(position),
                context.triggerKind,
                context.triggerCharacter
              )
            }
            signatureHelpProvider={(
              model: monaco.editor.ITextModel,
              position: monaco.Position,
              token: monaco.CancellationToken,
              context: monaco.languages.SignatureHelpContext
            ) => this._provideSignatureHelpAsync(model.getValue(), position, context)}
          />
        </div>

        {showExpandEditorButton && (
          <IconButton
            styles={expandEditorIconStyles}
            iconProps={{ iconName: expanded ? 'ChevronUp' : 'ChevronDown' }}
            onClick={() => {
              this.setState({ expanded: !expanded });
              this.updateHeight(!expanded);
            }}
            ariaLabel={expandEditorButtonAriaLabel}
          />
        )}
      </div>
    );
  }

  public updateHeight = (overrideExpand?: boolean) => {
    const expanded = overrideExpand ?? this.state.expanded;
    const { height: prevHeight } = this.state;

    const calculatedHeight = this._getHeight();
    const height = !this.props.showExpandEditorButton
      ? calculatedHeight
      : expanded
        ? // when expanded, set the height to full height of the text (upto maxLineCount) but no shorter than 7 lines tall
        Math.max(minExpandedEditorHeight, calculatedHeight)
        : // when collapsed, set the height to minLineCount tall but no shorter than formulaEditorMinHeight
        Math.max(formulaEditorMinHeight, this.props.minLineCount * oneLineHeight);

    if (height !== prevHeight) {
      this.setState({ height });
    }
  };

  private _getHeight(): number {
    const { minLineCount, maxLineCount } = this.props;

    if (minLineCount === maxLineCount) {
      return Math.max(formulaEditorMinHeight, maxLineCount * oneLineHeight);
    }

    let height: number = oneLineHeight;
    // TODO: A better solution is to use editor.onDidContentSizeChange and editor.getContentHeight() added in monaco-editor 0.20.0 to
    // automatically set the height
    if (this._editor && this._model && !this._model.isDisposed()) {
      const lineCount = this._model.getLineCount() || 1;

      try {
        if (lineCount <= 1 && height > 200) {
          // recent version of monaco editor 0.27.0 returns inconsistent
          // height such as getTopForLineNumber(1) === 0 and getTopForLineNumber(2) === 703
          // when the model has no content or a single line of content without newline.
          // allow this behavior to be trackable
          // throw new PowerFxFormulaEditorError(
          //   `fx editor computed invalid height ${height} for lineCount ${lineCount}`
          // );
        }
      } catch (error) {
        this.props?.onError?.(error as Error);
      }

      height = this._editor.getTopForLineNumber(lineCount + 1) + oneLineHeight;
    }

    return Math.min(
      Math.max(formulaEditorMinHeight, maxLineCount * oneLineHeight),
      Math.max(formulaEditorMinHeight, minLineCount * oneLineHeight, height)
    );
  }

  private _editorDidMount = (monacoEditor: monaco.editor.ICodeEditor, monacoParam: typeof monaco): void => {
    this._editor = monacoEditor;
    this._model = monacoEditor.getModel();
    this._monaco = monacoParam;
    if (!this._model || !this._monaco || this._model.isDisposed()) {
      return;
    }

    // Create PowerFx language client
    const { messageProcessor, getDocumentUriAsync, lspConfig, onEditorDidMount } = this.props;
    this._languageClient = new PowerFxLanguageClient(
      getDocumentUriAsync,
      messageProcessor.sendAsync,
      this._handleDiagnosticsNotification,
      this._handleTokensNotification
    );

    if (!lspConfig?.disableDidOpenNotification) {
      this._languageClient.notifyDidOpenAsync(this._model.getValue());
    }

    // Update height
    this.updateHeight();

    // Notify caller on editor did mount event
    if (onEditorDidMount) {
      onEditorDidMount(monacoEditor, monacoParam);
    }
  };

  private _handleDiagnosticsNotification = (params: PublishDiagnosticsParams): void => {
    if (!this._monaco || !this._model) {
      return;
    }

    const markers = params.diagnostics.map((error: Diagnostic) => {
      const { start, end } = error.range;
      return {
        severity: this._getMarkerSeverity(error.severity),
        message: error.message,
        startLineNumber: start.line,
        startColumn: start.character,
        endLineNumber: end.line,
        endColumn: end.character + 1 // end character is included
      };
    });

    this._monaco.editor.setModelMarkers(this._model, '', markers);
  };

  private _provideCompletionAsync = async (
    currentText: string,
    position: monaco.Position,
    currentWordPosition: monaco.editor.IWordAtPosition | null,
    triggerKind: monaco.languages.CompletionTriggerKind,
    triggerCharacter?: string
  ): Promise<monaco.languages.CompletionList> => {
    const { lspConfig } = this.props;
    if (lspConfig?.disableCompletionRequest) {
      return { suggestions: [] };
    }

    const result = await this._languageClient?.requestProvideCompletionItemsAsync(
      currentText,
      position.lineNumber - 1,
      position.column - 1,
      this._getCompletionTriggerKind(triggerKind),
      triggerCharacter
    );
    if (!result) {
      return { suggestions: [] };
    }

    const range = currentWordPosition
      ? {
        startColumn: currentWordPosition.startColumn,
        endColumn: currentWordPosition.endColumn,
        startLineNumber: position.lineNumber,
        endLineNumber: position.lineNumber
      }
      : {
        startColumn: position.column,
        endColumn: position.column,
        startLineNumber: position.lineNumber,
        endLineNumber: position.lineNumber
      };

    const suggestions: monaco.languages.CompletionItem[] = result.items.map(item => {
      const label = item.label;
      this._normalizedCompletionLookup[label.toLowerCase()] = label;
      return {
        label,
        documentation: item.documentation,
        detail: item.detail,
        kind: this._getCompletionKind(item.kind),
        range,
        insertText: item.label
      };
    });

    return {
      incomplete: !currentWordPosition,
      suggestions
    };
  };

  private _handleTokensNotification = (params: PublishTokensParams): void => {
    if (!this._monaco || !this._model) {
      return;
    }

    const names: HighlightedName[] = [];
    for (const [key, value] of Object.entries(params.tokens)) {
      this._normalizedCompletionLookup[key.toLowerCase()] = key;
      switch (value) {
        case TokenResultType.Function:
          names.push({ name: key, kind: NameKind.Function });
          break;

        case TokenResultType.Variable:
          names.push({ name: key, kind: NameKind.Variable });
          break;

        case TokenResultType.HostSymbol:
          names.push({ name: key, kind: NameKind.HostSymbol });
          break;

        default:
          break;
      }
    }

    this._onNamesChanged?.(names);
  };

  private _provideSignatureHelpAsync = async (
    currentText: string,
    position: monaco.Position,
    context: monaco.languages.SignatureHelpContext
  ): Promise<monaco.languages.SignatureHelpResult> => {
    const noResult: monaco.languages.SignatureHelpResult = {
      value: {
        signatures: [],
        activeSignature: 0,
        activeParameter: 0
      },
      dispose: () => {
        return;
      }
    };
    const { lspConfig } = this.props;
    if (!lspConfig?.enableSignatureHelpRequest) {
      return noResult;
    }

    const result = await this._languageClient?.requestProvideSignatureHelpAsync(
      currentText,
      position.lineNumber - 1,
      position.column - 1
    );
    if (!result || result.signatures.length === 0) {
      return noResult;
    }

    return {
      value: {
        signatures: result.signatures.map(item => ({
          label: item.label,
          documentation: item.documentation,
          parameters: item.parameters || [],
          activeParameter: item.activeParameter
        })),
        activeSignature: result.activeSignature || 0,
        activeParameter: result.activeParameter || 0
      },
      dispose: () => {
        return;
      }
    };
  };

  private _onContentChanged = async (value: string): Promise<void> => {
    const { onChange, lspConfig } = this.props;
    if (onChange) {
      onChange(value);
    }

    if (!lspConfig?.disableDidChangeNotification) {
      this._languageClient?.notifyDidChangeAsync(value, this._version++);
    }

    // Update height
    this.updateHeight();
  };

  private _getCompletionTriggerKind = (
    kind: monaco.languages.CompletionTriggerKind
  ): CompletionTriggerKind => {
    switch (kind) {
      case monaco.languages.CompletionTriggerKind.Invoke:
        return CompletionTriggerKind.Invoked;
      case monaco.languages.CompletionTriggerKind.TriggerCharacter:
        return CompletionTriggerKind.TriggerCharacter;
      case monaco.languages.CompletionTriggerKind.TriggerForIncompleteCompletions:
        return CompletionTriggerKind.TriggerForIncompleteCompletions;
      default:
        throw new Error('Unknown trigger kind!');
    }
  };

  private _getCompletionKind = (kind?: CompletionItemKind): monaco.languages.CompletionItemKind => {
    switch (kind) {
      case CompletionItemKind.Text:
        return monaco.languages.CompletionItemKind.Text;
      case CompletionItemKind.Method:
        return monaco.languages.CompletionItemKind.Method;
      case CompletionItemKind.Function:
        return monaco.languages.CompletionItemKind.Function;
      case CompletionItemKind.Constructor:
        return monaco.languages.CompletionItemKind.Constructor;
      case CompletionItemKind.Field:
        return monaco.languages.CompletionItemKind.Field;
      case CompletionItemKind.Variable:
        return monaco.languages.CompletionItemKind.Variable;
      case CompletionItemKind.Class:
        return monaco.languages.CompletionItemKind.Class;
      case CompletionItemKind.Interface:
        return monaco.languages.CompletionItemKind.Interface;
      case CompletionItemKind.Module:
        return monaco.languages.CompletionItemKind.Module;
      case CompletionItemKind.Property:
        return monaco.languages.CompletionItemKind.Property;
      case CompletionItemKind.Unit:
        return monaco.languages.CompletionItemKind.Unit;
      case CompletionItemKind.Value:
        return monaco.languages.CompletionItemKind.Value;
      case CompletionItemKind.Enum:
        return monaco.languages.CompletionItemKind.Enum;
      case CompletionItemKind.Keyword:
        return monaco.languages.CompletionItemKind.Keyword;
      case CompletionItemKind.Snippet:
        return monaco.languages.CompletionItemKind.Snippet;
      case CompletionItemKind.Color:
        return monaco.languages.CompletionItemKind.Color;
      case CompletionItemKind.File:
        return monaco.languages.CompletionItemKind.File;
      case CompletionItemKind.Reference:
        return monaco.languages.CompletionItemKind.Reference;
      case CompletionItemKind.Folder:
        return monaco.languages.CompletionItemKind.Folder;
      case CompletionItemKind.EnumMember:
        return monaco.languages.CompletionItemKind.EnumMember;
      case CompletionItemKind.Constant:
        return monaco.languages.CompletionItemKind.Constant;
      case CompletionItemKind.Struct:
        return monaco.languages.CompletionItemKind.Struct;
      case CompletionItemKind.Event:
        return monaco.languages.CompletionItemKind.Event;
      case CompletionItemKind.Operator:
        return monaco.languages.CompletionItemKind.Operator;
      case CompletionItemKind.TypeParameter:
        return monaco.languages.CompletionItemKind.TypeParameter;
      default:
        return monaco.languages.CompletionItemKind.Method;
    }
  };

  private _getMarkerSeverity = (severity?: DiagnosticSeverity): monaco.MarkerSeverity => {
    switch (severity) {
      case DiagnosticSeverity.Error:
        return monaco.MarkerSeverity.Error;
      case DiagnosticSeverity.Hint:
        return monaco.MarkerSeverity.Hint;
      case DiagnosticSeverity.Information:
        return monaco.MarkerSeverity.Info;
      case DiagnosticSeverity.Warning:
        return monaco.MarkerSeverity.Warning;
      default:
        return monaco.MarkerSeverity.Error;
    }
  };
}
