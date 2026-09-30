export type SelectionNode = {
  id: string;
  name: string;
  type: string;
  ancestorFrameIds?: string[];
  x?: number;
  y?: number;
};

export type SelectedFrame = {id: string; name: string};

// Figma does not define selection order. Canvas order makes the resulting deck predictable: rows top-to-bottom,
// then frames left-to-right. A selected frame nested in another selected frame is not a second slide.
export function topLevelSelectedFrames(nodes: readonly SelectionNode[]): SelectedFrame[] {
  const selectedFrameIds = new Set(nodes.filter(node => node.type === 'FRAME').map(node => node.id));
  return nodes
    .filter(node => node.type === 'FRAME' && !(node.ancestorFrameIds ?? []).some(id => selectedFrameIds.has(id)))
    .sort((a, b) => (a.y ?? 0) - (b.y ?? 0) || (a.x ?? 0) - (b.x ?? 0) || a.id.localeCompare(b.id))
    .map(({id, name}) => ({id, name}));
}

export function selectionToExportRequest(fileKey: string | undefined, nodes: readonly SelectionNode[]) {
  const frames = topLevelSelectedFrames(nodes);
  return {
    fileKey: fileKey ?? '',
    frameIds: frames.map(frame => frame.id),
    frameNames: frames.map(frame => frame.name),
  };
}
