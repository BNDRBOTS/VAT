import type { ReactNode } from "react";

type Props = {
  columns: string[];
  rows: (string | number | ReactNode)[][];
  maxRows?: number;
};

export function DataTable({ columns, rows, maxRows = 50 }: Props) {
  const visibleRows = rows.slice(0, maxRows);
  return (
    <div className="table-wrap">
      <table className="table">
        <thead>
          <tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr>
        </thead>
        <tbody>
          {visibleRows.length > 0 ? visibleRows.map((row, index) => (
            <tr key={index}>{row.map((cell, cellIndex) => <td key={`${index}-${cellIndex}`}>{cell}</td>)}</tr>
          )) : (
            <tr><td colSpan={columns.length} className="empty-state-cell">No rows.</td></tr>
          )}
        </tbody>
      </table>
      {rows.length > maxRows ? <div className="table-note">Showing first {maxRows} of {rows.length} rows.</div> : null}
    </div>
  );
}