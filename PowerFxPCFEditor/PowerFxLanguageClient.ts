
import { sendDataAsync } from './lsp_helper';

export class PowerFxLanguageClient {
  public constructor(private _lsp_url: string, private _onDataReceived: (data: string) => void) {
  }

  public async sendAsync(data: string) {
    // Hardcoded FormulaType : 1
    //const payloadData = JSON.stringify({ FormulaBody: data, FormulaType: 1 });
    const payloadData = JSON.stringify({ FormulaBody: data});
    console.log('[LSP Client] Send: ' + payloadData);

    try {
      const result = await sendDataAsync(this._lsp_url, 'lsp', payloadData);
      if (!result.ok) {
        return;
      }
      const response = await result.json();
      if (response) {
        const responseArray = response.LanguageServerData;
        responseArray.forEach((item: string) => {
          console.log('[LSP Client] Receive: ' + item);
          this._onDataReceived(item);
        })
      }
    } catch (err) {
      console.log(err);
    }
  }
}
