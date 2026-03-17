import React from "react";

export interface ComparisonRow {
  key: string;
  field: string;
  left: string;
  right: string;
  leftTitle?: string;
  rightTitle?: string;
  pass?: boolean;
}

interface ComparisonTableProps {
  leftLabel: string;
  rightLabel: string;
  rows: ComparisonRow[];
}

function resolveParity(row: ComparisonRow): boolean {
  if (typeof row.pass === "boolean") {
    return row.pass;
  }
  return row.left === row.right;
}

export function ComparisonTable({
  leftLabel,
  rightLabel,
  rows,
}: ComparisonTableProps): JSX.Element {
  return (
    <table className="comparison-table">
      <thead>
        <tr>
          <th>Field</th>
          <th>{leftLabel}</th>
          <th>{rightLabel}</th>
          <th>Parity</th>
        </tr>
      </thead>
      <tbody>
        {rows.map((row) => {
          const pass = resolveParity(row);
          return (
            <tr
              key={row.key}
              className={pass ? "comparison-row is-pass" : "comparison-row is-mismatch"}
            >
              <td>{row.field}</td>
              <td title={row.leftTitle ?? row.left}>{row.left}</td>
              <td title={row.rightTitle ?? row.right}>{row.right}</td>
              <td>
                <span className={`parity-pill ${pass ? "is-pass" : "is-mismatch"}`}>
                  {pass ? "PASS" : "MISMATCH"}
                </span>
              </td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}
