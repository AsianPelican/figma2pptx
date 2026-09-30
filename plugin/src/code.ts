import {selectionToExportRequest, type SelectionNode} from './selection';

const SECRET_KEY = 'figma2pptx.bridgeSecret';
const ACTIVE_JOB_KEY = 'figma2pptx.activeJobId';

figma.showUI(__html__, {width: 420, height: 560, themeColors: true});

function selectionNode(node: SceneNode): SelectionNode {
  const bounds = 'absoluteBoundingBox' in node ? node.absoluteBoundingBox : null;
  const ancestorFrameIds: string[] = [];
  let parent = node.parent;
  while (parent && parent.type !== 'PAGE' && parent.type !== 'DOCUMENT') {
    if (parent.type === 'FRAME') ancestorFrameIds.push(parent.id);
    parent = parent.parent;
  }
  return {
    id: node.id,
    name: node.name,
    type: node.type,
    ancestorFrameIds,
    x: bounds?.x,
    y: bounds?.y,
  };
}

async function postState(): Promise<void> {
  const request = selectionToExportRequest(figma.fileKey, figma.currentPage.selection.map(selectionNode));
  const bridgeSecret = await figma.clientStorage.getAsync(SECRET_KEY) as string | undefined;
  const activeJobId = await figma.clientStorage.getAsync(ACTIVE_JOB_KEY) as string | undefined;
  figma.ui.postMessage({type: 'state', request, bridgeSecret: bridgeSecret ?? '', activeJobId: activeJobId ?? ''});
}

figma.ui.onmessage = async (message: {type?: string; secret?: string; jobId?: string}) => {
  if (message.type === 'save-secret') {
    const secret = message.secret?.trim() ?? '';
    if (secret) await figma.clientStorage.setAsync(SECRET_KEY, secret);
    else await figma.clientStorage.deleteAsync(SECRET_KEY);
    await postState();
  }
  if (message.type === 'save-active-job') {
    const jobId = message.jobId?.trim() ?? '';
    if (jobId) await figma.clientStorage.setAsync(ACTIVE_JOB_KEY, jobId);
    else await figma.clientStorage.deleteAsync(ACTIVE_JOB_KEY);
  }
  if (message.type === 'close') figma.closePlugin();
};

figma.on('selectionchange', () => void postState());
void postState();
