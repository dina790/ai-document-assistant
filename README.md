\# AI Document Assistant



A local prototype for asking questions about PDF, Word, and Excel files.



\## Features



\- Local AI answers using Ollama

\- Table or bullet-point output

\- PDF page, Word paragraph/table-row, and Excel worksheet/row citations

\- Source ID and exact-quotation validation

\- Separate Excel totals calculation mode



\## Running locally



Install the dependencies:



```bat

python -m pip install -r requirements.txt

```



Install Ollama and download the model:



```bat

ollama pull qwen2.5:3b

```



Make sure Ollama is running, then start the backend:



```bat

python -m uvicorn backend:app --reload

```



Open http://127.0.0.1:8000/assistant



\## Limitations



\- Documents are stored temporarily in backend memory.

\- PDF extraction requires selectable text.

\- Word extraction reads body paragraphs and top-level tables.

\- Excel formulas require saved results.

\- AI answers and missing-information detection can be incorrect.

\- Quotation matching does not prove that the evidence supports the answer.

\- This prototype is intended for local use, not unsecured public hosting.

