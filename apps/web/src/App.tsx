import { ChangeEvent, useState } from "react";
import type { Sheet, WorkbookPayload } from "./types";

const apiBase = "/api/excel";

export function App() {
  const [workbook, setWorkbook] = useState<WorkbookPayload | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [isBusy, setIsBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const currentSheet = workbook?.sheets[activeSheet];
  const columnCount = getColumnCount(currentSheet?.rows ?? []);

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    setIsBusy(true);
    setError(null);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch(`${apiBase}/parse`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const message = await response.text();
        throw new Error(message || "No se pudo procesar el archivo.");
      }

      const parsedWorkbook = (await response.json()) as WorkbookPayload;
      setWorkbook(parsedWorkbook);
      setActiveSheet(0);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Error inesperado.");
    } finally {
      setIsBusy(false);
      event.target.value = "";
    }
  }

  function updateCell(rowIndex: number, columnIndex: number, value: string) {
    if (!workbook) return;

    setWorkbook({
      ...workbook,
      sheets: workbook.sheets.map((sheet, sheetIndex) => {
        if (sheetIndex !== activeSheet) return sheet;

        const rows = sheet.rows.map((row) => [...row]);
        rows[rowIndex][columnIndex] = value;
        return { ...sheet, rows };
      }),
    });
  }

  function addRow() {
    if (!workbook) return;

    updateCurrentSheet((sheet) => ({
      ...sheet,
      rows: [...sheet.rows, Array.from({ length: Math.max(getColumnCount(sheet.rows), 1) }, () => "")],
    }));
  }

  function addColumn() {
    if (!workbook) return;

    updateCurrentSheet((sheet) => ({
      ...sheet,
      rows: sheet.rows.length ? sheet.rows.map((row) => [...row, ""]) : [[""]],
    }));
  }

  function updateCurrentSheet(updater: (sheet: Sheet) => Sheet) {
    if (!workbook) return;

    setWorkbook({
      ...workbook,
      sheets: workbook.sheets.map((sheet, sheetIndex) => (sheetIndex === activeSheet ? updater(sheet) : sheet)),
    });
  }

  async function downloadWorkbook() {
    if (!workbook) return;

    setIsBusy(true);
    setError(null);

    try {
      const response = await fetch(`${apiBase}/export`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(workbook),
      });

      if (!response.ok) {
        const message = await response.text();
        throw new Error(message || "No se pudo exportar el archivo.");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${withoutExtension(workbook.fileName)}-editado.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Error inesperado.");
    } finally {
      setIsBusy(false);
    }
  }

  return (
    <main className="app-shell">
      <section className="hero">
        <div className="hero-copy">
          <p className="eyebrow">React Vite + NestJS</p>
          <h1>Previsualizador y editor de Excel</h1>
          <p>
            Sube un archivo Excel o CSV, edita sus celdas desde una tabla interactiva y descarga una copia
            actualizada generada por la API NestJS.
          </p>
        </div>

        <label className="upload-card">
          <span>XLS</span>
          <strong>{isBusy ? "Procesando..." : "Seleccionar archivo"}</strong>
          <small>.xlsx o .csv</small>
          <input
            type="file"
            accept=".xlsx,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
            disabled={isBusy}
            onChange={handleFileChange}
          />
        </label>
      </section>

      <section className="toolbar">
        <div>
          <strong>{workbook?.fileName ?? "Sin archivo cargado"}</strong>
          {currentSheet && <span>{currentSheet.rows.length} filas x {Math.max(columnCount, 1)} columnas</span>}
        </div>
        <div className="actions">
          <button type="button" disabled={!workbook || isBusy} onClick={addRow}>Agregar fila</button>
          <button type="button" disabled={!workbook || isBusy} onClick={addColumn}>Agregar columna</button>
          <button type="button" className="primary" disabled={!workbook || isBusy} onClick={downloadWorkbook}>
            Descargar XLSX
          </button>
        </div>
      </section>

      {error && <p className="error">{error}</p>}

      <section className="workspace">
        {workbook ? (
          <>
            <nav className="sheet-tabs" aria-label="Hojas del libro">
              {workbook.sheets.map((sheet, index) => (
                <button
                  type="button"
                  className={index === activeSheet ? "active" : ""}
                  key={sheet.name}
                  onClick={() => setActiveSheet(index)}
                >
                  {sheet.name}
                </button>
              ))}
            </nav>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th aria-label="Numero de fila" />
                    {Array.from({ length: Math.max(columnCount, 1) }, (_, columnIndex) => (
                      <th key={columnIndex}>{columnLabel(columnIndex)}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {currentSheet?.rows.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      <th>{rowIndex + 1}</th>
                      {Array.from({ length: Math.max(columnCount, 1) }, (_, columnIndex) => (
                        <td key={columnIndex}>
                          <input
                            value={row[columnIndex] ?? ""}
                            onChange={(event) => updateCell(rowIndex, columnIndex, event.target.value)}
                            aria-label={`Celda ${columnLabel(columnIndex)}${rowIndex + 1}`}
                          />
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        ) : (
          <div className="empty-state">
            <h2>Tu hoja aparecerá aquí</h2>
            <p>La API NestJS lee el archivo y devuelve las hojas como datos editables.</p>
          </div>
        )}
      </section>
    </main>
  );
}

function getColumnCount(rows: string[][]) {
  return rows.reduce((max, row) => Math.max(max, row.length), 0);
}

function columnLabel(index: number) {
  let label = "";
  let value = index + 1;

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

function withoutExtension(fileName: string) {
  return fileName.replace(/\.[^.]+$/, "") || "libro";
}
