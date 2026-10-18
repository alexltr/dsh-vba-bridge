/**
 * dsh-vba-bridge — read and edit Excel VBA projects through Excel COM.
 *
 * Six model-facing tools (`vba_inspect`, `vba_read_module`, `vba_write_module`,
 * `vba_export_modules`, `vba_list_procedures`, `vba_doctor`) delegate to a small
 * Python bridge (`bridge/vba_bridge.py`) that drives a hidden Excel instance.
 * The offline OVBA parser is the read-only fallback when COM is unavailable.
 *
 * @module dsh-vba-bridge
 */
import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const DEFAULT_BRIDGE = path.join(HERE, 'bridge', 'vba_bridge.py')

export const name = 'vba-bridge'
export const inject = ['tools']

const DEFAULT_TIMEOUT_MS = 240_000
const DEFAULT_MAX_OUTPUT_CHARS = 200_000

function fail(message, hint) {
  return new Error(hint ? `${message}\n\nHint: ${hint}` : message)
}

function summarizeError(error, stderr) {
  const tail = String(stderr || '').trim().slice(-1500)
  return tail ? `${error.message}\n\nBridge stderr (tail):\n${tail}` : error.message
}

export function apply(ctx, config = {}) {
  const pythonPath = String(config.pythonPath || process.env.DSH_VBA_PYTHON || 'python')
  const bridgePath = String(config.bridgePath || DEFAULT_BRIDGE)
  const timeoutMs = Number(config.timeoutMs) > 0 ? Number(config.timeoutMs) : DEFAULT_TIMEOUT_MS
  const maxOutputChars =
    Number(config.maxOutputChars) > 0 ? Number(config.maxOutputChars) : DEFAULT_MAX_OUTPUT_CHARS
  const allowWrite = config.allowWrite !== false

  if (!existsSync(bridgePath)) {
    ctx.logger?.warn?.(`vba-bridge: bridge script not found at ${bridgePath}`)
  }

  // Excel COM is single-instance and file writes are not concurrency safe.
  let chain = Promise.resolve()

  function invoke(request, signal) {
    return new Promise((resolve, reject) => {
      let settled = false
      const finish = (fn, value) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        signal?.removeEventListener?.('abort', onAbort)
        fn(value)
      }
      const child = spawn(pythonPath, [bridgePath, '--json-stdin'], {
        windowsHide: true,
        stdio: ['pipe', 'pipe', 'pipe'],
        env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
      })
      let stdout = ''
      let stderr = ''
      child.stdout.setEncoding('utf8')
      child.stderr.setEncoding('utf8')
      child.stdout.on('data', (chunk) => {
        stdout += chunk
      })
      child.stderr.on('data', (chunk) => {
        stderr += chunk
      })
      const onAbort = () => {
        try {
          child.kill()
        } catch {}
        finish(reject, new Error('vba_bridge call aborted'))
      }
      const timer = setTimeout(() => {
        try {
          child.kill()
        } catch {}
        finish(reject, new Error(`vba_bridge timed out after ${timeoutMs} ms`))
      }, timeoutMs)
      signal?.addEventListener?.('abort', onAbort, { once: true })

      child.on('error', (error) => {
        const hint =
          error.code === 'ENOENT'
            ? `Python executable "${pythonPath}" was not found. Set config.pythonPath in cordis.patch.yml.`
            : undefined
        finish(reject, fail(`cannot start the VBA bridge: ${error.message}`, hint))
      })
      child.on('close', (code) => {
        const line = stdout.trim().split(/\r?\n/).filter(Boolean).pop() || ''
        let payload
        try {
          payload = JSON.parse(line)
        } catch {
          finish(
            reject,
            fail(
              `the VBA bridge returned no JSON (exit ${code})`,
              summarizeError(new Error('invalid bridge output'), stderr || stdout),
            ),
          )
          return
        }
        if (!payload.ok) {
          const info = payload.error || {}
          const hint = [info.hint, stderr.trim()].filter(Boolean).join('\n')
          finish(reject, fail(`[${info.code || 'BRIDGE_ERROR'}] ${info.message || 'bridge failed'}`, hint))
          return
        }
        finish(resolve, payload.data)
      })
      child.stdin.on('error', () => {})
      child.stdin.end(JSON.stringify(request))
    })
  }

  function call(request, signal) {
    const run = () => invoke(request, signal)
    const pending = chain.then(run, run)
    chain = pending.catch(() => {})
    return pending
  }

  async function resolveInput(value, exec) {
    if (!value || path.isAbsolute(value) || !exec?.agent) return value
    let wd
    try {
      wd = typeof ctx.get === 'function' ? ctx.get('workingDirectory') : undefined
    } catch {
      return value
    }
    if (!wd || typeof wd.ensure !== 'function') return value
    try {
      const base = await wd.ensure(exec.agent, exec.signal)
      return base ? path.resolve(base, value) : value
    } catch {
      return value
    }
  }

  function textBlock(text) {
    const cap = maxOutputChars
    const clipped =
      text.length > cap ? `${text.slice(0, cap)}\n… [truncated ${text.length - cap} chars]` : text
    return [{ type: 'text', text: clipped }]
  }

  const fileParam = {
    type: 'string',
    description:
      'Path to a macro-enabled workbook (.xlsm/.xlsb/.xls/.xlam). Relative paths resolve against the session working directory. The file is not copied unless backup is requested on write.',
  }

  ctx.tools.register({
    name: 'vba_inspect',
    description:
      'Inspect the VBA project of an Excel workbook: project name, protection state, every component (name, type, line count) and the VBA references. Uses Excel COM by default; falls back to offline OVBA parsing (read-only) when Excel is unavailable.',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        file: fileParam,
        engine: {
          type: 'string',
          enum: ['auto', 'com', 'ovba'],
          description: 'auto (default) prefers Excel COM and falls back to offline parsing.',
        },
      },
      required: ['file'],
    },
    output: {
      schema: { type: 'object' },
      render: (_args, value) => {
        const lines = []
        lines.push(`VBA project: ${value.project}  [engine: ${value.engine}]`)
        if (value.excel_version) lines.push(`Excel version: ${value.excel_version}`)
        if (value.protected === true) lines.push('Protection: PASSWORD-PROTECTED (editing blocked)')
        else if (value.protected === false) lines.push('Protection: none')
        lines.push(`File: ${value.file}`)
        lines.push(`SHA-256: ${value.sha256}`)
        if (Array.isArray(value.warnings) && value.warnings.length) {
          lines.push(`Warnings: ${value.warnings.join('; ')}`)
        }
        lines.push('', `Components (${value.components.length}):`)
        for (const c of value.components) {
          const type = c.type ?? `type${c.type_id}`
          lines.push(`- ${c.name} (${type}, ${c.lines} lines)`)
        }
        if (value.references?.length) {
          lines.push('', 'References:')
          for (const r of value.references) {
            lines.push(`- ${r.name} ${r.version}${r.builtin ? ' (built-in)' : ''}`)
          }
        }
        if (value.note) lines.push('', value.note)
        return textBlock(lines.join('\n'))
      },
    },
    timeoutMs,
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      if (!args?.file) throw fail('file is required')
      const file = await resolveInput(String(args.file), exec)
      return call(
        { action: 'inspect', file, engine: args.engine || 'auto' },
        exec?.signal,
      )
    },
  })

  ctx.tools.register({
    name: 'vba_read_module',
    description:
      'Read the VBA source code of one module (standard module, class, form or document module such as ThisWorkbook/Sheet1). Works for UserForm code-behind, including event handlers like btnCancel_Click. Returns CRLF source as Excel reports it, without the hidden VB_* attribute lines.',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        file: fileParam,
        module: { type: 'string', description: 'Exact VBComponent name, e.g. "Module1" or "ThisWorkbook".' },
        engine: {
          type: 'string',
          enum: ['auto', 'com', 'ovba'],
          description: 'auto (default) prefers Excel COM and falls back to offline parsing.',
        },
      },
      required: ['file', 'module'],
    },
    output: {
      schema: { type: 'object' },
      render: (_args, value) => {
        const head = `Module ${value.module} (${value.type}, ${value.lines} lines) [engine: ${value.engine}]`
        const body = value.code ? '```vb\n' + value.code + '\n```' : '(module is empty)'
        return textBlock(`${head}\n\n${body}`)
      },
    },
    timeoutMs,
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      if (!args?.file) throw fail('file is required')
      if (!args?.module) throw fail('module is required')
      const file = await resolveInput(String(args.file), exec)
      return call(
        {
          action: 'read',
          file,
          module: String(args.module),
          engine: args.engine || 'auto',
        },
        exec?.signal,
      )
    },
  })

  ctx.tools.register({
    name: 'vba_write_module',
    description:
      'Replace or append the VBA source of one module in an .xlsm/.xlsb/.xls workbook, save the workbook in place, then reopen it and verify the stored source. Also rewrites UserForm code-behind (event handlers); it does not touch the form control layout. Creates the module when create=true. A timestamped .bak copy is written next to the file by default. Requires Excel installed and Trust Center "Trust access to the VBA project object model" enabled.',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        file: fileParam,
        module: { type: 'string', description: 'Exact VBComponent name; created as a standard module when create=true.' },
        code: { type: 'string', description: 'Complete VBA source for the module (CRLF or LF).' },
        mode: {
          type: 'string',
          enum: ['replace', 'append'],
          description: 'replace (default) overwrites the module body; append adds the code after existing code.',
        },
        create: {
          type: 'boolean',
          description: 'Create a standard module when it does not exist (default false).',
        },
        backup: {
          type: 'boolean',
          description: 'Write a timestamped .bak copy before saving (default true).',
        },
      },
      required: ['file', 'module', 'code'],
    },
    output: {
      schema: { type: 'object' },
      render: (_args, value) => {
        const lines = [
          `Wrote module ${value.module} (${value.mode}${value.created ? ', created' : ''})`,
          `File: ${value.file}`,
          `Lines: ${value.lines}   changed: ${value.changed}   verified: ${value.verified}`,
          `SHA-256: ${value.sha256_before} -> ${value.sha256_after}`,
        ]
        if (value.backup) lines.push(`Backup: ${value.backup}`)
        if (!value.verified) lines.push('WARNING: the reopened module does not match the requested source.')
        return textBlock(lines.join('\n'))
      },
    },
    timeoutMs,
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      if (!allowWrite) throw fail('vba_write_module is disabled by the plugin config (allowWrite=false)')
      if (!args?.file) throw fail('file is required')
      if (!args?.module) throw fail('module is required')
      if (typeof args?.code !== 'string') throw fail('code must be a string')
      const file = await resolveInput(String(args.file), exec)
      return call(
        {
          action: 'write',
          file,
          module: String(args.module),
          code: args.code,
          mode: args.mode || 'replace',
          create: args.create === true,
          backup: args.backup !== false,
          engine: 'com',
        },
        exec?.signal,
      )
    },
  })

  ctx.tools.register({
    name: 'vba_export_modules',
    description:
      'Export every VBA component of a workbook into a directory. mode=code writes plain source files (.bas/.cls/.frm) for reading and diffing; mode=component uses VBComponent.Export and also writes .frx for UserForms, preserving the form designer. Writes only under outDir; the workbook itself is opened read-only and never modified.',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        file: fileParam,
        outDir: {
          type: 'string',
          description:
            'Output directory. Relative paths resolve next to the workbook. Created when missing.',
        },
        mode: {
          type: 'string',
          enum: ['code', 'component'],
          description:
            'code (default): code module only; component: native VBA export files including .frm/.frx for forms.',
        },
        include: {
          type: 'array',
          items: { type: 'string' },
          description: 'Optional list of component names; omit to export all.',
        },
        overwrite: {
          type: 'boolean',
          description: 'Overwrite existing exported files (default false: skip and report them).',
        },
      },
      required: ['file', 'outDir'],
    },
    output: {
      schema: { type: 'object' },
      render: (_args, value) => {
        const lines = [
          `Exported ${value.counts.exported} component(s) to ${value.outDir} (mode: ${value.mode})`,
        ]
        for (const item of value.exported.slice(0, 60)) {
          const extra = item.files.length > 1 ? ` + ${item.files.slice(1).join(', ')}` : ''
          lines.push(`- ${item.file} (${item.type}, ${item.lines} lines)${extra}`)
        }
        if (value.exported.length > 60) lines.push(`… ${value.exported.length - 60} more`)
        if (value.skipped.length) {
          lines.push(`Skipped ${value.skipped.length} existing file(s); pass overwrite=true to replace.`)
        }
        if (value.errors.length) lines.push(`Errors: ${value.errors.map((e) => e.module).join(', ')}`)
        if (value.mode === 'code' && value.exported.some((item) => item.type === 'form')) {
          lines.push('Note: UserForm .frm files contain code-behind only; mode=component adds the .frx designer data.')
        }
        lines.push(`Manifest: ${value.manifest}`)
        return textBlock(lines.join('\n'))
      },
    },
    timeoutMs,
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      if (!args?.file) throw fail('file is required')
      if (!args?.outDir) throw fail('outDir is required')
      const file = await resolveInput(String(args.file), exec)
      return call(
        {
          action: 'export',
          file,
          outDir: String(args.outDir),
          mode: args.mode || 'code',
          include: Array.isArray(args.include) ? args.include.map(String) : null,
          overwrite: args.overwrite === true,
          engine: 'com',
        },
        exec?.signal,
      )
    },
  })

  ctx.tools.register({
    name: 'vba_list_procedures',
    description:
      'List every Sub, Function and Property declaration with its start/end line and arguments, for all modules or one module. Includes UserForm event handlers such as btnCancel_Click and document-module handlers such as Workbook_Open.',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        file: fileParam,
        module: { type: 'string', description: 'Optional exact component name; omit to scan every module.' },
        engine: {
          type: 'string',
          enum: ['auto', 'com', 'ovba'],
          description: 'auto (default) prefers Excel COM and falls back to offline parsing.',
        },
      },
      required: ['file'],
    },
    output: {
      schema: { type: 'object' },
      render: (_args, value) => {
        const short = (text) => (text.length > 72 ? `${text.slice(0, 69)}...` : text)
        const lines = [
          `Procedures: ${value.total} in ${value.modules.length} module(s) [engine: ${value.engine}]`,
        ]
        for (const mod of value.modules) {
          lines.push(`- ${mod.module} (${mod.type}, ${mod.lines} lines)`)
          for (const proc of mod.procedures) {
            const scope = proc.scope === 'default' ? '' : `${proc.scope} `
            const args = proc.args ? `(${short(proc.args)})` : '()'
            lines.push(`  - ${scope}${proc.kind} ${proc.name}${args}  L${proc.start}-L${proc.end}`)
          }
        }
        if (value.line_numbers_note) lines.push(value.line_numbers_note)
        return textBlock(lines.join('\n'))
      },
    },
    timeoutMs,
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      if (!args?.file) throw fail('file is required')
      const file = await resolveInput(String(args.file), exec)
      return call(
        {
          action: 'procedures',
          file,
          module: args.module ? String(args.module) : null,
          engine: args.engine || 'auto',
        },
        exec?.signal,
      )
    },
  })

  ctx.tools.register({
    name: 'vba_doctor',
    description:
      'Run this first after installing the plugin. Checks the Python runtime, pywin32, Excel COM startup, the Trust Center AccessVBOM setting, the offline parser, the temporary directory, and optionally one workbook. Returns PASS/FAIL per check with a fix hint.',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        file: {
          type: 'string',
          description: 'Optional workbook to test. Relative paths resolve against the session working directory.',
        },
      },
      required: [],
    },
    output: {
      schema: { type: 'object' },
      render: (_args, value) => {
        const lines = [`VBA Bridge doctor — ${value.ready ? 'READY' : 'NOT READY'}`]
        for (const check of value.checks) {
          lines.push(`[${check.ok ? 'PASS' : 'FAIL'}] ${check.name}: ${check.detail}`)
        }
        const hints = value.checks.filter((check) => !check.ok && check.hint)
        if (hints.length) {
          lines.push('', 'How to fix:')
          for (const check of hints) lines.push(`- ${check.name}: ${check.hint}`)
        }
        if (value.file) lines.push('', `File: ${value.file}`)
        return textBlock(lines.join('\n'))
      },
    },
    timeoutMs,
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const file = args?.file ? await resolveInput(String(args.file), exec) : null
      return call({ action: 'doctor', file, engine: 'com' }, exec?.signal)
    },
  })
}
