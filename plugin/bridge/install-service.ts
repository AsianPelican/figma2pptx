import {chmodSync, existsSync, mkdirSync, writeFileSync} from 'node:fs';
import {homedir} from 'node:os';
import {join, resolve} from 'node:path';
import {loadOrCreateSecret} from './server.ts';
import {loadBridgeEnvironment} from './launch.ts';
import {BRIDGE_BIND_HOST, BRIDGE_BIND_PORT, BRIDGE_PORT} from './protocol.ts';
import {runtimePaths} from './runtime.ts';
import {assertBridgeDependencies, SERVICE_PATH} from './preflight.ts';

const LABEL = 'local.figma2pptx.bridge';
const LAUNCHER_ID = 'local.figma2pptx.bridge.launcher';
const LAUNCHER_REQUIREMENT = `=designated => identifier "${LAUNCHER_ID}"`;

const xml = (value: string) => value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');

async function run(args: string[]): Promise<{code: number; text: string}> {
  const process = Bun.spawn(args, {stdout: 'pipe', stderr: 'pipe'});
  const [stdout, stderr, code] = await Promise.all([new Response(process.stdout).text(), new Response(process.stderr).text(), process.exited]);
  return {code, text: `${stdout}${stderr}`.trim()};
}

async function stableLauncher(runtimeRoot: string): Promise<string> {
  const app = join(runtimeRoot, 'launcher', 'Figma2Pptx Bridge.app');
  const contents = join(app, 'Contents');
  const executable = join(contents, 'MacOS', 'figma2pptx-bridge-launcher');
  if (!existsSync(executable)) {
    mkdirSync(join(contents, 'MacOS'), {recursive: true, mode: 0o700});
    writeFileSync(join(contents, 'Info.plist'), `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>CFBundleIdentifier</key><string>${LAUNCHER_ID}</string>
  <key>CFBundleName</key><string>Figma2Pptx Bridge</string>
  <key>CFBundleExecutable</key><string>figma2pptx-bridge-launcher</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleVersion</key><string>1</string>
  <key>LSUIElement</key><true/>
</dict></plist>
`, {mode: 0o600});
    const compile = await run(['/usr/bin/clang', '-Os', '-Wall', '-Wextra', '-o', executable, join(import.meta.dir, 'launcher.c')]);
    if (compile.code !== 0) throw Error(`launcher compilation failed: ${compile.text}`);
    chmodSync(executable, 0o700);
  }
  let requirement = await run(['/usr/bin/codesign', '-d', '-r-', app]);
  if (requirement.code !== 0 || !requirement.text.includes(`designated => identifier "${LAUNCHER_ID}"`) || requirement.text.includes('cdhash')) {
    const sign = await run(['/usr/bin/codesign', '--force', '--sign', '-', '--identifier', LAUNCHER_ID, '--requirements', LAUNCHER_REQUIREMENT, app]);
    if (sign.code !== 0) throw Error(`launcher signing failed: ${sign.text}`);
  }
  const verify = await run(['/usr/bin/codesign', '--verify', '--strict', app]);
  if (verify.code !== 0) throw Error(`launcher signature verification failed: ${verify.text}`);
  requirement = await run(['/usr/bin/codesign', '-d', '-r-', app]);
  if (requirement.code !== 0 || !requirement.text.includes(`identifier "${LAUNCHER_ID}"`) || requirement.text.includes('cdhash')) {
    throw Error(`launcher does not have a stable designated requirement: ${requirement.text}`);
  }
  return executable;
}

export async function installService(): Promise<void> {
  const uid = process.getuid?.();
  if (uid === undefined) throw Error('cannot determine the current user id');
  const pluginDir = resolve(import.meta.dir, '..');
  const launchScript = join(import.meta.dir, 'launch.ts');
  const paths = runtimePaths();
  const configDir = paths.config;
  const envPath = join(configDir, 'bridge.env');
  const launchAgents = join(homedir(), 'Library', 'LaunchAgents');
  const plistPath = join(launchAgents, `${LABEL}.plist`);
  for (const path of [paths.root, paths.config, paths.cache, paths.jobs, paths.logs, paths.temp]) mkdirSync(path, {recursive: true, mode: 0o700});
  mkdirSync(launchAgents, {recursive: true});
  loadBridgeEnvironment(envPath);
  loadOrCreateSecret(configDir);
  const launcher = await stableLauncher(paths.root);

  const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>${LABEL}</string>
  <key>ProgramArguments</key><array><string>${xml(launcher)}</string><string>${xml(process.execPath)}</string><string>${xml(launchScript)}</string></array>
  <key>WorkingDirectory</key><string>${xml(pluginDir)}</string>
  <key>EnvironmentVariables</key><dict>
    <key>FIGMA2PPTX_RUNTIME_DIR</key><string>${xml(paths.root)}</string>
    <key>TMPDIR</key><string>${xml(paths.temp)}</string>
    <key>PATH</key><string>${xml(SERVICE_PATH)}</string>
  </dict>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>LimitLoadToSessionType</key><string>Aqua</string>
  <key>ProcessType</key><string>Interactive</string>
  <key>StandardOutPath</key><string>${xml(join(paths.logs, 'bridge.log'))}</string>
  <key>StandardErrorPath</key><string>${xml(join(paths.logs, 'bridge.error.log'))}</string>
</dict></plist>
`;
  writeFileSync(plistPath, plist, {mode: 0o600});
  chmodSync(plistPath, 0o600);
  assertBridgeDependencies(SERVICE_PATH);

  const domain = `gui/${uid}`;
  const existing = await run(['launchctl', 'print', `${domain}/${LABEL}`]);
  if (existing.code === 0) {
    const bootout = await run(['launchctl', 'bootout', `${domain}/${LABEL}`]);
    if (bootout.code !== 0) throw Error(`launchctl bootout failed for ${LABEL}: ${bootout.text}`);
  }
  let bootstrap = await run(['launchctl', 'bootstrap', domain, plistPath]);
  if (bootstrap.code !== 0) {
    await Bun.sleep(250);
    bootstrap = await run(['launchctl', 'bootstrap', domain, plistPath]);
  }
  if (bootstrap.code !== 0) throw Error(`launchctl bootstrap failed: ${bootstrap.text}`);
  const kick = await run(['launchctl', 'kickstart', `${domain}/${LABEL}`]);
  if (kick.code !== 0) throw Error(`launchctl kickstart failed: ${kick.text}`);
  const tailscale = existsSync('/Applications/Tailscale.app/Contents/MacOS/Tailscale')
    ? '/Applications/Tailscale.app/Contents/MacOS/Tailscale'
    : '/opt/homebrew/bin/tailscale';
  if (!existsSync(tailscale)) throw Error('Tailscale CLI not found');
  const serve = await run([tailscale, 'serve', '--bg', '--yes', '--https', String(BRIDGE_PORT), `http://${BRIDGE_BIND_HOST}:${BRIDGE_BIND_PORT}`]);
  if (serve.code !== 0) throw Error(`tailscale serve failed: ${serve.text}`);
  console.log(`Installed and started ${LABEL}`);
  console.log(`Bridge secret: ${join(configDir, 'bridge-secret')} (not printed)`);
}

if (import.meta.main) {
  try { await installService(); }
  catch (error) { console.error(`figma2pptx bridge service: ${error instanceof Error ? error.message : String(error)}`); process.exit(1); }
}
