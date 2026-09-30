import {expect, test} from 'bun:test';
import {selectionToExportRequest, topLevelSelectedFrames, type SelectionNode} from '../src/selection.ts';

test('maps only top-level selected frames in deterministic canvas order', () => {
  const nodes: SelectionNode[] = [
    {id: '4:1', name: 'Text', type: 'TEXT', x: 0, y: 0},
    {id: '3:1', name: 'Third', type: 'FRAME', x: 500, y: 600},
    {id: '1:1', name: 'First', type: 'FRAME', x: 100, y: 100},
    {id: '2:2', name: 'Nested duplicate', type: 'FRAME', ancestorFrameIds: ['1:1'], x: 120, y: 120},
    {id: '2:1', name: 'Second', type: 'FRAME', x: 400, y: 100},
  ];
  const before = structuredClone(nodes);
  expect(topLevelSelectedFrames(nodes)).toEqual([
    {id: '1:1', name: 'First'},
    {id: '2:1', name: 'Second'},
    {id: '3:1', name: 'Third'},
  ]);
  expect(nodes).toEqual(before);
});

test('builds the bridge request from the file key and selected frame IDs', () => {
  expect(selectionToExportRequest('AbCdEfGhIjKl', [
    {id: '9:2', name: 'Overview', type: 'FRAME'},
    {id: '9:3', name: 'Caption', type: 'TEXT'},
  ])).toEqual({fileKey: 'AbCdEfGhIjKl', frameIds: ['9:2'], frameNames: ['Overview']});
});
