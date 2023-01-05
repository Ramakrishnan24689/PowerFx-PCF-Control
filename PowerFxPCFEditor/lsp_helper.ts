export async function sendDataAsync(LSP_URL: string, endpoint: string, data: string): Promise<Response> {
  return await fetch(LSP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: data
  });
}