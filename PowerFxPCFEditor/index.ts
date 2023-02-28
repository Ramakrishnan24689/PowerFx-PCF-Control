import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { ControlContainer, ControlContainerProps } from "./ControlContainer";
import { ContextEx } from "./Component.types";

export class PowerFxPCFEditor implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    _notifyOutputChanged: () => void;
    _editorState: IOutputs
    context: ComponentFramework.Context<IInputs>
    recId: string;
    entityName: string;
    entityRecordJString: string;
    prevDefaultValue: string;
    defaultValueChanged: boolean;
    componentKey: string;
    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        context.mode.trackContainerResize(true);
        this._notifyOutputChanged = notifyOutputChanged;
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        const contextEx = (context as unknown as ContextEx);
        const pageURL = this.parsePageURL();
        const entityName = context?.parameters.entityName.raw ?? pageURL.etn ?? contextEx?.page.entityTypeName;
        const defaultValue = context.parameters.defaultValue.raw ?? '';
        let lspServiceURL = context.parameters.lspServiceURL.raw;
        if (!lspServiceURL && !contextEx.mode.isAuthoringMode) {
            lspServiceURL = contextEx.page.getClientUrl();
            lspServiceURL = lspServiceURL.concat('/api/data/v9.0/RetrieveLanguageServerData');
        }
        // Check if default formula changed
        this.defaultValueChanged = false;
        if (this.prevDefaultValue !== defaultValue) {
            this.prevDefaultValue = defaultValue;
            this.defaultValueChanged = true;
            this.componentKey = this.componentKey === undefined ? entityName : this.componentKey.concat("_1") ;
        }

        const allocatedWidth = parseInt(context?.mode.allocatedWidth as unknown as string);
        const allocatedHeight = parseInt(context?.mode.allocatedHeight as unknown as string);
        let formulaContext: string | undefined;
        let formulaContextProp: string | undefined;
        if (!this.entityRecordJString) {
            this.generateRecordContext(entityName, pageURL.id).then(entityRecordJString => {
                if (entityRecordJString && entityRecordJString !== this.entityRecordJString) {
                    this.entityRecordJString = entityRecordJString;
                }
            });
        }
        if (formulaContextProp && formulaContextProp.length > 0) {
            formulaContext = formulaContextProp;
        } else if (this.entityRecordJString) {
            formulaContext = this.entityRecordJString;
        }
        else {
            this.entityRecordJString = "getExpressionType=true&localeName=en-US&getTokensFlags=1";
            formulaContext = this.entityRecordJString;
        }

        const props: ControlContainerProps = {
            recId: pageURL.id ?? '',
            entityName: context?.parameters.entityName.raw ?? pageURL.etn ?? contextEx?.page.entityTypeName,
            lspServiceURL: lspServiceURL ?? '',
            editorMinLine: context.parameters.editorMinLine.raw ?? 1,
            editorMaxLine: context.parameters.editorMaxLine.raw ?? 4,
            width: allocatedWidth,
            height: allocatedHeight,
            //formula: context.parameters.formula.raw ?? '',
            formula: this.defaultValueChanged ? defaultValue : context.parameters.formula.raw ?? '',
            defaultValueChanged: this.defaultValueChanged,
            formulaContext: formulaContext,
            onEditorStateChanged: (editorState: IOutputs) => { this._editorState = editorState; this._notifyOutputChanged(); },
            isReadOnly: context.parameters.ReadOnly.raw,
            key: this.componentKey
        };

        return React.createElement(ControlContainer, props);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return this._editorState
    }


    private parsePageURL(): { [key: string]: string } {
        const parsedURLObj: { [key: string]: string } = {};

        try {
            const urlParams = window.location.href.split("?")[1];
            if (urlParams) {
                const urlParamSections = urlParams.split("&");
                for (let paramPairStr of urlParamSections) {
                    const paramPair = paramPairStr.split("=");
                    parsedURLObj[paramPair[0]] = paramPair[1];
                }
            }
        } catch {
            console.log("Error parsing URL");
        }

        return parsedURLObj;
    }

    private async generateRecordContext(entityName: string, recId: string) {
        if (recId && entityName) {
            const rawEntityRecord = await this.context.webAPI.retrieveRecord(entityName, recId);
            const entityRecord: ComponentFramework.WebApi.Entity = {};
            for (const key of Object.keys(rawEntityRecord)) {
                if (!key.startsWith("_") && key.indexOf("@") === -1 && key.indexOf(".") === -1) {
                    entityRecord[key] = rawEntityRecord[key];
                }
            }
            return JSON.stringify(entityRecord);
        }
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
