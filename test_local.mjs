/**
 * Local harness: exercise the plugin exactly as the tool registry would, without
 * installing the bundle. Run with `node test_local.mjs`.
 */
import { apply } from './index.js'

const tools = new Map()
const ctx = {
  logger: { warn: (message) => console.error('[warn]', message) },
  tools: {
    register(definition) {
      if (tools.has(definition.name)) throw new Error(`duplicate tool ${definition.name}`)
      tools.set(definition.name, definition)
      return () => tools.delete(definition.name)
    },
  },
}

apply(ctx, {})
console.log('registered tools:', [...tools.keys()].join(', '))

const file = process.argv[2] || String.raw`F:\MyProjects\dsh\vba_bridge\work\edit3.xlsm`
const signal = new AbortController().signal

async function run(name, args) {
  const definition = tools.get(name)
  const value = await definition.execute(args, { signal })
  const content = definition.output.render(args, value)
  console.log(`\n=== ${name} ===`)
  console.log(content.map((block) => block.text).join('\n'))
  return value
}

const info = await run('vba_inspect', { file })
const target = info.components.find((c) => c.lines > 0)?.name || info.components[0]?.name
await run('vba_read_module', { file, module: target })
await run('vba_list_procedures', { file, module: target })
await run('vba_export_modules', {
  file,
  outDir: String.raw`F:\MyProjects\dsh\vba_bridge\work\plugin-export`,
  mode: 'component',
  overwrite: true,
})

const copy = String.raw`F:\MyProjects\dsh\vba_bridge\work\plugin-test.xlsm`
const { copyFileSync } = await import('node:fs')
copyFileSync(file, copy)
await run('vba_write_module', {
  file: copy,
  module: 'DshPluginTest',
  create: true,
  code: 'Sub DshPluginTest()\r\n    Debug.Print "written by dsh-vba-bridge"\r\nEnd Sub\r\n',
})
await run('vba_read_module', { file: copy, module: 'DshPluginTest' })
