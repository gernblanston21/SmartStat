export function compactMiddle(value: string, start = 14, end = 12): string {
  if (value.length <= start + end + 3) {
    return value;
  }
  return `${value.slice(0, start)}...${value.slice(-end)}`;
}

export function compactHash(value: string): string {
  return compactMiddle(value, 12, 12);
}

export function compactPath(value: string): string {
  return compactMiddle(value, 26, 18);
}

export function compactRuleId(value: string): string {
  return compactMiddle(value, 18, 16);
}
