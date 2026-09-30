import {join, resolve} from 'node:path';

export type RuntimePaths = {
  root: string;
  config: string;
  cache: string;
  jobs: string;
  logs: string;
  temp: string;
};

export function runtimePaths(root = process.env.FIGMA2PPTX_RUNTIME_DIR || resolve(import.meta.dir, '..', '..', '..', 'figma2pptx-runtime')): RuntimePaths {
  const absolute = resolve(root);
  return {
    root: absolute,
    config: join(absolute, 'config'),
    cache: join(absolute, 'cache'),
    jobs: join(absolute, 'jobs'),
    logs: join(absolute, 'logs'),
    temp: join(absolute, 'jobs', '.tmp'),
  };
}
