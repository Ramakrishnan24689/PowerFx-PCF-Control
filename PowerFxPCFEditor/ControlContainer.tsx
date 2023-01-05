import * as React from "react";
import { EditorState, PowerFxEditor } from "./PowerFxEditor";

export interface ControlContainerProps extends React.ClassAttributes<ControlContainer> {
  onEditorStateChanged?: (newState: EditorState) => void;
  entityName: string;
  formula: string;
  formulaContext: string;
  recId: string;
  editorMaxLine?: number;
  editorMinLine?: number;
  lspServiceURL?: string;
  width?: number;
  height?: number;
}

interface ControlContainerState {
  recId: string | null;
  entityName: string;
  entityRecordJString?: string;

}


export class ControlContainer extends React.Component<ControlContainerProps, ControlContainerState> {
  constructor(props: ControlContainerProps) {
    super(props);


    this.state = {
      recId: props.recId,
      entityName: props.entityName
    };
  }

  public static deriveStateFromProps(props: ControlContainerProps): ControlContainerState {
    return {
      recId: props.recId,
      entityName: props.entityName
    }
  }

  public render() {


    if (!this.props.lspServiceURL) {
      return <div>No LSP endpoint provided.</div>
    }

    if (!this.props.formulaContext) {
      return <div>Loading record context ...</div>
    }

    return (
      // <div className="pa-fx-editor-container">
      <PowerFxEditor
        lsp_url={this.props.lspServiceURL}
        formula={this.props.formula}
        formulaContext={this.props.formulaContext!}
        editorMaxLine={this.props.editorMaxLine!}
        editorMinLine={this.props.editorMinLine}
        onEditorStateChanged={this.props.onEditorStateChanged}
        width={this.props.width}
        height={this.props.height}
        entityName={this.props.entityName}
      />
      //</div>
    );
  }

}
