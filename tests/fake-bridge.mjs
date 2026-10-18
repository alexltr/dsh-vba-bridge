/** Minimal bridge stand-in: echoes the request so tests can inspect it. */
let raw = ''
process.stdin.setEncoding('utf8')
process.stdin.on('data', (chunk) => {
  raw += chunk
})
process.stdin.on('end', () => {
  const request = JSON.parse(raw)
  process.stdout.write(`${JSON.stringify({ ok: true, data: { echo: request } })}\n`)
})
