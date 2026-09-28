#!/usr/bin/env bun
// figma2pptx command line entry. See `figma2pptx --help`.
import {main} from './cli/main';

process.exit(await main(process.argv.slice(2)));
