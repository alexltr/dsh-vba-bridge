/** Prepack cleanup: never ship Python bytecode caches. */
import { readdir, rm } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')

async function walk(dir) {
  let entries
  try {
    entries = await readdir(dir, { withFileTypes: true })
  } catch {
    return
  }
  for (const entry of entries) {
    const full = path.join(dir, entry.name)
    if (entry.isDirectory()) {
      if (entry.name === '__pycache__') {
        await rm(full, { recursive: true, force: true })
      } else {
        await walk(full)
      }
    } else if (entry.name.endsWith('.pyc') || entry.name.endsWith('.pyo')) {
      await rm(full, { force: true })
    }
  }
}

for (const target of ['bridge', 'vendor', 'tests']) {
  await walk(path.join(ROOT, target))
}
