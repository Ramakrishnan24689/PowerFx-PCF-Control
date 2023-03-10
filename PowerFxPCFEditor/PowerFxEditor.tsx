
import * as React from 'react';
import { IDisposable, MessageProcessor } from '@microsoft/power-fx-formulabar/lib';
import { PowerFxFormulaEditor } from './PowerFXFormulaBar-Forked/PowerFxFormulaEditor';
import { PowerFxLanguageClient } from './PowerFxLanguageClient';
import { sendDataAsync } from './lsp_helper';
// import loader from '@monaco-editor/loader';
// import * as monaco from 'monaco-editor';

export interface EditorState {
  formula: string;
  formulaContext: string;
  parameters?: string;
  error?: string;
  evaluateValue?: string;
  key?: string;
  defaultValueChanged?: boolean;
}

export interface PowerFxEditorProps {
  lsp_url: string
  defaultValueChanged: boolean;
  formula: string;
  formulaContext: string;
  editorMaxLine?: number;
  editorMinLine?: number;
  entityName?: string;
  onEditorStateChanged?: (newState: EditorState) => void;
  width?: number;
  height?: number;
  isReadOnly?: boolean;
  key?: string,
  parameters: string
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
      defaultValueChanged: props.defaultValueChanged,
      parameters: props.parameters
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
        this._languageClient.sendAsync(data, this.state.parameters ?? "")
    };

    // Required to preconfigure monaco loader which fails in custom page during initial load
    // Reference implementation - https://github.com/suren-atoyan/monaco-loader
    // loader.config({ monaco });
    // loader.init().then((monaco) => {
    //   console.info("monaco instance initialized");
    // }).catch((error) => { console.error(error) });

  }

  public async componentDidUpdate(prevProps: Readonly<PowerFxEditorProps>, prevState: Readonly<EditorState>) {
    if (this.props.parameters !== prevProps.parameters) {
      this.setState({ parameters: this.props.parameters });
    }
  }

  public async componentDidMount() {
    if (this.props.formula && this.props.formula.length > 0) {
      setTimeout(() => this._evalAsync(this.props.formula), 50);
    }
  }

  public render() {
    const { formula, evaluateValue } = this.state;
    const { editorMaxLine, editorMinLine, isReadOnly } = this.props;
    return (
      <>
        <PowerFxFormulaEditor
          getDocumentUriAsync={this._getDocumentUriAsync}
          defaultValue={formula}
          messageProcessor={this._messageProcessor}
          maxLineCount={editorMaxLine || 1}
          minLineCount={editorMinLine || 1}
          monacoEditorOptions={{ fixedOverflowWidgets: false }}
          onChange={this._onExpressionChanged}
          lspConfig={{
            enableSignatureHelpRequest: true
          }}
          width={this.props.width}
          isReadOnly={isReadOnly}
          key={this.props.key}
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
    return `powerfx://formula_columns?entityLogicalName=` + this.props.entityName + `&getExpressionType=true&localeName=en-US&getTokensFlags=1`;
  };
}
