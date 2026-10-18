// Verifies downloaded add-in release assets with stylistic-hub's own verifier and its key list on main.
// Run from the stylistic-hub checkout: bun ../stylistic-addon/.claude/skills/stylistic-addon-release/verify-release.ts <dir> <version>
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

// The hub's modules, resolved from the working directory (the stylistic-hub checkout), not from this file.
const hub = (path: string): string => pathToFileURL(resolve(path)).href
const { componentSchema, envelopeSchema, keyListSchema } = await import(hub('src/shared/signing/documents.ts'))
const { verifyArtifact, verifyComponent } = await import(hub('src/shared/signing/verify.ts'))

const [dir, version] = process.argv.slice(2)
if (!dir || !version) throw new Error('Usage: verify-release.ts <assets dir> <version>')

const readJson = (file: string): unknown => JSON.parse(readFileSync(join(dir, file), 'utf8'))
const keys = keyListSchema.parse(JSON.parse(execFileSync('git', ['show', 'main:keys/keys.json']).toString()))
const component = componentSchema.parse(readJson('component.json'))
const envelope = envelopeSchema.parse(readJson('component.json.sig'))
const zip = readFileSync(join(dir, `stylistic-word-addin-${version}.zip`))

console.log(
  `component: ${verifyComponent(component, envelope, keys, 'stylistic-addon')}, artifact: ${verifyArtifact(component, {
    sha256: createHash('sha256').update(zip).digest('hex'),
    size: zip.length
  })}`
)
console.log(`keyId ${envelope.keyId}, ${component.id} ${component.version} ${component.channel}, ${zip.length} bytes`)
console.log(`requires ${JSON.stringify(component.requires)}`)
