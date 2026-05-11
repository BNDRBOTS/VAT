type Props = {
  title: string;
  value: string;
  subtitle: string;
  tone?: "default" | "good" | "warn" | "bad";
};

export function MetricCard({ title, value, subtitle, tone = "default" }: Props) {
  return (
    <div className={`metric-card metric-${tone}`}>
      <div className="metric-title">{title}</div>
      <div className="metric-value">{value}</div>
      <div className="metric-subtitle">{subtitle}</div>
    </div>
  );
}