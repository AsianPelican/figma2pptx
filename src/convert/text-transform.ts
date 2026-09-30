// Optional host-supplied text transforms. The public converter owns the mechanism; callers own every policy.

export type TextTransformOptions = {
  outline?: (style: any) => boolean;
  substitute?: (style: any) => any | null;
};

export type TextTransformPlan = {outlinedRuns: number, substitutedRuns: number, outline: boolean};
export type SourceRun = {style: any, length: number};

// Contiguous Figma character-style overrides are the source runs. The base style is override 0.
export function sourceRuns(n: any): SourceRun[] {
  const chars = String(n.characters ?? '');
  if (!chars.length) return [];
  const overrides = n.characterStyleOverrides || [], table = n.styleOverrideTable || {};
  const out: SourceRun[] = [];
  let last = -1;
  for (let i = 0; i < chars.length; i++) {
    const id = overrides[i] || 0;
    if (id !== last) out.push({style: {...n.style, ...(table[id] || {})}, length: 1});
    else out[out.length - 1].length++;
    last = id;
  }
  return out;
}

export function transformStyle(style: any, options: TextTransformOptions): {style: any, substituted: boolean} {
  const mapped = options.substitute?.(style);
  return mapped ? {style: mapped, substituted: true} : {style, substituted: false};
}

export function textTransformPlan(n: any, options: TextTransformOptions): TextTransformPlan {
  const runs = sourceRuns(n);
  const outlined = options.outline ? runs.filter(r => options.outline!(r.style)) : [];
  if (outlined.length) {
    if (outlined.length !== runs.length) throw Error(`text transform cannot outline mixed-policy text node ${n.id}; split the outlined text into its own Figma text node`);
    return {outlinedRuns: outlined.length, substitutedRuns: 0, outline: true};
  }
  return {outlinedRuns: 0, substitutedRuns: options.substitute ? runs.filter(r => options.substitute!(r.style)).length : 0, outline: false};
}
