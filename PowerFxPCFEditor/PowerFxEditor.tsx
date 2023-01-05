
import * as React from 'react';
import { IDisposable, MessageProcessor, PowerFxFormulaEditor } from '@microsoft/power-fx-formulabar/lib';
import { sendDataAsync } from './lsp_helper';
import { PowerFxLanguageClient } from './PowerFxLanguageClient';

export interface EditorState {
  formula: string;
  formulaContext: string;
  error?: string;
  evaluateValue?: string
}

export interface PowerFxEditorProps {
  lsp_url: string
  formula: string;
  formulaContext: string;
  editorMaxLine?: number;
  editorMinLine?: number;
  entityName?: string;
  onEditorStateChanged?: (newState: EditorState) => void;
  width?: number;
  height?: number;
}

export class PowerFxEditor extends React.PureComponent<PowerFxEditorProps, EditorState> {
  private _languageClient: PowerFxLanguageClient;
  private _messageProcessor: MessageProcessor;
  private _listener: (data: string) => void = () => null;
  constructor(props: PowerFxEditorProps) {
    super(props);

    this.state = {
      formula: props.formula,
      formulaContext: props.formulaContext,
    };

    const onDataReceived = (data: string) => {
      this._listener(data);
    };

    this._languageClient = new PowerFxLanguageClient(this.props.lsp_url, onDataReceived);
    this._messageProcessor = {
      addListener: (listener: (data: string) => void): IDisposable => {
        this._listener = listener;
        return {
          dispose: () => null
        };
      },
      sendAsync: async (data: string): Promise<void> =>
        this._languageClient.sendAsync(data)
    };
  }

  public async componentDidMount() {
    if (this.props.formula && this.props.formula.length > 0) {
      setTimeout(() => this._evalAsync(this.props.formula), 50);
    }
  }

  public render() {
    const { formula, evaluateValue } = this.state;
    const { editorMaxLine, editorMinLine } = this.props;
    return (
      <>
        <PowerFxFormulaEditor
          getDocumentUriAsync={this._getDocumentUriAsync}
          defaultValue={formula}
          messageProcessor={this._messageProcessor}
          maxLineCount={editorMaxLine || 1}
          minLineCount={editorMinLine || 1}
          monacoEditorOptions={{ fixedOverflowWidgets: false, dimensions: {width : this.props.width, height: this.props.height} }}
          onChange={this._onExpressionChanged}
          lspConfig={{
            enableSignatureHelpRequest: true
          }}
          key={this.props.entityName}
        />
        <div style={{ minHeight: 21, border: '#d2d0ce 1px solid' }}>{evaluateValue ?? ''}</div>
      </>);
  }

  private _onExpressionChanged = (newValue: string): void => {
    const { onEditorStateChanged } = this.props;

    this.setState({ formula: newValue }, () => {
      onEditorStateChanged?.(this.state);
    });

    this._evalAsync(newValue);
  }

  private _evalAsync = async (expression: string): Promise<void> => {
    const { lsp_url, onEditorStateChanged } = this.props;
    const { formulaContext, formula } = this.state;
    const result = await sendDataAsync(lsp_url, 'eval', JSON.stringify({ context: formulaContext, expression }));
    if (!result.ok) {
      return;
    }

    const response = await result.json();
    let newState: EditorState = { formula, formulaContext, evaluateValue: '', error: '' };
    if (response.result) {
      newState.evaluateValue = response.result;
    } else if (response.error) {
      newState.error = response.error;
    }
    this.setState(newState, () => { onEditorStateChanged?.(newState) });
  };

  private _getDocumentUriAsync = async (): Promise<string> => {
    //return `powerfx://demo?context=${this.state.formulaContext}`; 
    //return `powerfx://demo?context=${this.state.formulaContext}`;
    //return `powerfx://formula_columns?entityLogicalName=cr7a3_pfx&getExpressionType=true&localeName=en-US&getTokensFlags=1`;
    console.log(`powerfx://formula_columns?entityLogicalName=` + this.props.entityName + `&getExpressionType=true&localeName=en-US&getTokensFlags=1`);
    return `powerfx://formula_columns?entityLogicalName=` + this.props.entityName + `&getExpressionType=true&localeName=en-US&getTokensFlags=1`;
  };
  /*
{"FormulaBody":"{\"jsonrpc\":\"2.0\",\"id\":\"19bb11d4-5d93-4536-a215-ed463c1f280a\",\"method\":\"textDocument/signatureHelp\",\"params\":{\"textDocument\":{\"uri\":\"powerfx://formula_columns?entityLogicalName=cr7a3_pfx&getExpressionType=true&localeName=en-US&expression=Back%28\"},\"position\":{\"line\":0,\"character\":5}}}"}
  */
}