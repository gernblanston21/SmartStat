import React from "react";
import { FilterState } from "../types";

interface FilterBarProps {
  filters: FilterState;
  leagues: string[];
  recordTypes: string[];
  sourceTypes: string[];
  onChange: (next: FilterState) => void;
}

function titleCase(value: string): string {
  if (!value) {
    return value;
  }
  return value.charAt(0).toUpperCase() + value.slice(1);
}

export function FilterBar({
  filters,
  leagues,
  recordTypes,
  sourceTypes,
  onChange,
}: FilterBarProps): JSX.Element {
  return (
    <section className="filter-bar">
      <label className="control">
        <span>Search</span>
        <input
          value={filters.search}
          onChange={(event) => onChange({ ...filters, search: event.target.value })}
          placeholder="Search current view"
        />
      </label>

      <label className="control">
        <span>League</span>
        <select
          value={filters.league}
          onChange={(event) => onChange({ ...filters, league: event.target.value })}
        >
          {leagues.map((league) => (
            <option key={league} value={league}>
              {titleCase(league)}
            </option>
          ))}
        </select>
      </label>

      <label className="control">
        <span>Record Type</span>
        <select
          value={filters.recordType}
          onChange={(event) => onChange({ ...filters, recordType: event.target.value })}
        >
          {recordTypes.map((recordType) => (
            <option key={recordType} value={recordType}>
              {titleCase(recordType)}
            </option>
          ))}
        </select>
      </label>

      <label className="control">
        <span>Source Type</span>
        <select
          value={filters.sourceType}
          onChange={(event) => onChange({ ...filters, sourceType: event.target.value })}
        >
          {sourceTypes.map((sourceType) => (
            <option key={sourceType} value={sourceType}>
              {titleCase(sourceType)}
            </option>
          ))}
        </select>
      </label>
    </section>
  );
}
