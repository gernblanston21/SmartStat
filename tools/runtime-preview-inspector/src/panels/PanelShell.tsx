import React from "react";

interface PanelShellProps {
  title: string;
  subtitle: string;
  badges?: string[];
  children: React.ReactNode;
}

export function PanelShell({
  title,
  subtitle,
  badges = [],
  children,
}: PanelShellProps): JSX.Element {
  return (
    <section className="panel-shell">
      <header className="panel-header">
        <div className="panel-title-row">
          <h2>{title}</h2>
          {badges.length ? (
            <ul className="panel-badges" aria-label={`${title} badges`}>
              {badges.map((badge) => (
                <li key={badge}>{badge}</li>
              ))}
            </ul>
          ) : null}
        </div>
        <p>{subtitle}</p>
      </header>
      <div className="panel-body">{children}</div>
    </section>
  );
}
