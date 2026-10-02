import { spawn } from 'node:child_process';
import { copyFile, mkdir, mkdtemp, readdir, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const publish = process.argv.includes('--publish');
const temporaryOutput = await mkdtemp(path.join(os.tmpdir(), 'triade-mail-release-'));
const releaseDirectory = path.resolve('release');
const builderCli = fileURLToPath(new URL('../node_modules/electron-builder/cli.js', import.meta.url));

try {
  const builder = spawn(
    process.execPath,
    [
      builderCli,
      '--win',
      'nsis',
      '--publish',
      publish ? 'always' : 'never',
      `-c.directories.output=${temporaryOutput}`,
    ],
    { stdio: 'inherit' },
  );

  const exitCode = await new Promise((resolve, reject) => {
    builder.once('error', reject);
    builder.once('close', resolve);
  });

  if (exitCode !== 0) {
    throw new Error(`electron-builder failed with exit code ${exitCode}`);
  }

  await mkdir(releaseDirectory, { recursive: true });
  const artifacts = await readdir(temporaryOutput, { withFileTypes: true });

  for (const artifact of artifacts) {
    if (artifact.isFile() && /\.(exe|yml|blockmap)$/i.test(artifact.name)) {
      await copyFile(path.join(temporaryOutput, artifact.name), path.join(releaseDirectory, artifact.name));
    }
  }
} finally {
  await rm(temporaryOutput, { force: true, recursive: true });
}