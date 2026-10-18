import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { test } from 'node:test'
import { fileURLToPath } from 'node:url'

import { apply } from '../index.js'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const ROOT = path.resolve(HERE, '..')
const FAKE_BRIDGE = path.join(HERE, 'fake-bridge.mjs')

function makeHarness(baseDir = path.join(ROOT, 'work')) {
  const tools = new Map()
  const ctx = {
    get(name) {
      if (name !== 'workingDirectory') return undefined
      return { ensure: async () => baseDir }
    },
    logger: { warn() {} },
    tools: {
      register(definition) {
        tools.set(definition.name, definition)
        return () => tools.delete(definition.name)
      },
    },
  }
  apply(ctx, { pythonPath: process.execPath, bridgePath: FAKE_BRIDGE })
  return tools
}

const exec = { agent: { id: 'test-session' }, signal: new AbortController().signal }

test('registers the six vba tools', () => {
  const tools = makeHarness()
  assert.deepEqual(
    [...tools.keys()].sort(),
    ['vba_doctor', 'vba_export_modules', 'vba_inspect', 'vba_list_procedures', 'vba_read_module', 'vba_write_module'],
  )
})

test('tool schemas are object schemas with the expected required fields', () => {
  const tools = makeHarness()
  for (const definition of tools.values()) {
    assert.equal(definition.parameters.type, 'object')
    assert.equal(definition.parameters.additionalProperties, false)
    assert.equal(typeof definition.execute, 'function')
    assert.equal(typeof definition.output.render, 'function')
  }
  assert.deepEqual(tools.get('vba_read_module').parameters.required, ['file', 'module'])
  assert.deepEqual(tools.get('vba_export_modules').parameters.required, ['file', 'outDir'])
  assert.deepEqual(tools.get('vba_doctor').parameters.required, [])
})

test('relative file paths resolve against the session working directory', async () => {
  const tools = makeHarness()
  const value = await tools.get('vba_inspect').execute({ file: 'book.xlsm' }, exec)
  assert.equal(value.echo.file, path.join(ROOT, 'work', 'book.xlsm'))
})

test('absolute file paths are left untouched', async () => {
  const tools = makeHarness()
  const absolute = path.join(ROOT, 'fixtures', 'sample.xlsm')
  const value = await tools.get('vba_read_module').execute({ file: absolute, module: '模块1' }, exec)
  assert.equal(value.echo.file, absolute)
  assert.equal(value.echo.module, '模块1')
})

test('doctor works without a file and passes the request through', async () => {
  const tools = makeHarness()
  const value = await tools.get('vba_doctor').execute({}, exec)
  assert.equal(value.echo.action, 'doctor')
  assert.equal(value.echo.file, null)
})

test('locale metadata and bundle patch are present', () => {
  const en = JSON.parse(readFileSync(path.join(ROOT, 'locale', 'en.json'), 'utf8'))
  const zh = JSON.parse(readFileSync(path.join(ROOT, 'locale', 'zh.json'), 'utf8'))
  assert.ok(en.meta.title && en.meta.description)
  assert.ok(zh.meta.title && zh.meta.description)
  const manifest = JSON.parse(readFileSync(path.join(ROOT, 'package.json'), 'utf8'))
  assert.equal(manifest.dsh.bundle.patch, './cordis.patch.yml')
  const patch = readFileSync(path.join(ROOT, 'cordis.patch.yml'), 'utf8')
  assert.match(patch, /dsh-vba-bridge/)
  assert.match(patch, /id: vba-bridge/)
})
