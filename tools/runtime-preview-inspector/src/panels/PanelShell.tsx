import React from "react";

interface PanelShellProps {
  title: string;
  subtitle: string;
  children: React.ReactNode;
}

export function PanelShell({ title, subtitle, children }: PanelShellProps): JSX.Element {
  return (
    <section className="panel-shell">
      <header className="panel-header">
        <h2>{title}</h2>
        <p>{subtitle}</p>
      </header>
      <div className="panel-body">{children}</div>
    </section>
  );
}
