import React from "react";

export type TabKey = "records" | "relationships" | "query_paths" | "traceability";

interface TabsProps {
  value: TabKey;
  onChange: (next: TabKey) => void;
}

const TAB_ORDER: Array<{ key: TabKey; label: string }> = [
  { key: "records", label: "Records" },
  { key: "relationships", label: "Relationships" },
  { key: "query_paths", label: "Query Paths" },
  { key: "traceability", label: "Traceability" },
];

export function Tabs({ value, onChange }: TabsProps): JSX.Element {
  return (
    <div className="tabs" role="tablist" aria-label="Semantic views">
      {TAB_ORDER.map((tab) => (
        <button
          key={tab.key}
          type="button"
          role="tab"
          aria-selected={value === tab.key}
          className={value === tab.key ? "tab active" : "tab"}
          onClick={() => onChange(tab.key)}
        >
          {tab.label}
        </button>
      ))}
    </div>
  );
}
