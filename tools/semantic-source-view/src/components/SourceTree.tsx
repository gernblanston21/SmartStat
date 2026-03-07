import React, { useMemo, useState } from "react";
import { SourceTreeNode, TreeSelection } from "../types";

interface SourceTreeProps {
  roots: SourceTreeNode[];
  selectedNodeId: string | null;
  onSelect: (selection: TreeSelection | null) => void;
}

function toSelection(node: SourceTreeNode): TreeSelection {
  return {
    nodeId: node.id,
    label: node.label,
    sourcePath: node.source_path,
    sourceType: node.source_type,
    league: node.league,
    recordIds: node.record_ids,
  };
}

interface NodeProps {
  node: SourceTreeNode;
  depth: number;
  selectedNodeId: string | null;
  expanded: Set<string>;
  onToggle: (id: string) => void;
  onSelect: (selection: TreeSelection | null) => void;
}

function TreeNodeRow({
  node,
  depth,
  selectedNodeId,
  expanded,
  onToggle,
  onSelect,
}: NodeProps): JSX.Element {
  const hasChildren = node.children.length > 0;
  const isExpanded = expanded.has(node.id);
  const isSelected = selectedNodeId === node.id;

  return (
    <li>
      <div
        className={isSelected ? "tree-row selected" : "tree-row"}
        style={{ paddingLeft: `${depth * 12 + 8}px` }}
      >
        {hasChildren ? (
          <button
            type="button"
            className="tree-toggle"
            onClick={() => onToggle(node.id)}
            aria-label={isExpanded ? "Collapse node" : "Expand node"}
          >
            {isExpanded ? "▾" : "▸"}
          </button>
        ) : (
          <span className="tree-toggle placeholder">•</span>
        )}
        <button type="button" className="tree-label" onClick={() => onSelect(toSelection(node))}>
          {node.label}
        </button>
      </div>

      {hasChildren && isExpanded ? (
        <ul className="tree-list">
          {node.children.map((child) => (
            <TreeNodeRow
              key={child.id}
              node={child}
              depth={depth + 1}
              selectedNodeId={selectedNodeId}
              expanded={expanded}
              onToggle={onToggle}
              onSelect={onSelect}
            />
          ))}
        </ul>
      ) : null}
    </li>
  );
}

export function SourceTree({ roots, selectedNodeId, onSelect }: SourceTreeProps): JSX.Element {
  const defaultExpanded = useMemo(() => new Set<string>(roots.map((root) => root.id)), [roots]);
  const [expanded, setExpanded] = useState<Set<string>>(defaultExpanded);

  const handleToggle = (id: string): void => {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  return (
    <aside className="source-tree">
      <header className="pane-header">
        <h2>Source Tree</h2>
        <button type="button" className="tiny-button" onClick={() => onSelect(null)}>
          Clear
        </button>
      </header>
      <ul className="tree-list">
        {roots.map((root) => (
          <TreeNodeRow
            key={root.id}
            node={root}
            depth={0}
            selectedNodeId={selectedNodeId}
            expanded={expanded}
            onToggle={handleToggle}
            onSelect={onSelect}
          />
        ))}
      </ul>
    </aside>
  );
}
