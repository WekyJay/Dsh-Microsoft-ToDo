/**
 * Regenerate `market/index.json` from the committed tree.
 *
 * The plugin market index lists one exact, immutable version per commit, so the
 * Git source pins HEAD. Run this after committing and before publishing the
 * `market/` directory (for example to GitHub Pages).
 *
 *   node scripts/make-index.mjs
 */
import { execFileSync } from 'node:child_process'
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const manifest = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))

const repoUrl = process.env.DSH_TODO_REPO
  ?? manifest.repository?.url?.replace(/^git\+/, '').replace(/\.git$/, '')
if (typeof repoUrl !== 'string' || !repoUrl.startsWith('https://')) {
  throw new Error('Set DSH_TODO_REPO=https://github.com/<owner>/<repo> or add repository.url to package.json')
}

const commit = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim()
const stamp = Number(execFileSync('git', ['show', '-s', '--format=%ct', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim())

const index = {
  format: 1,
  title: { en: 'Microsoft To Do market', zh: '微软待办插件源' },
  description: {
    en: 'Index for the Dsh-Microsoft-ToDo bundle.',
    zh: 'Dsh-Microsoft-ToDo 插件的自建市场索引。',
  },
  metadata: {
    categories: { productivity: { title: { en: 'Productivity', zh: '效率' } } },
    labels: { community: { title: { en: 'Community', zh: '社区' } } },
  },
  plugins: [
    {
      name: manifest.name,
      title: { en: 'Microsoft To Do', zh: '微软待办' },
      description: {
        en: 'Microsoft To Do panel: lists, tasks, local groups, and My Day.',
        zh: 'Microsoft To Do 面板：清单、任务、本机分组与「我的一天」。',
      },
      icon: 'icon.svg',
      author: manifest.author,
      homepage: repoUrl,
      repository: { url: repoUrl },
      metadata: {
        category: 'productivity',
        labels: ['community'],
        keywords: manifest.keywords ?? [],
      },
      'dist-tags': { latest: manifest.version },
      versions: [
        {
          version: manifest.version,
          publishTimestamp: stamp,
          source: { type: 'git', url: repoUrl, commit },
          engines: { dsh: manifest.engines?.dsh ?? '*' },
        },
      ],
    },
  ],
}

const outDir = join(root, 'market')
await mkdir(outDir, { recursive: true })
await copyFile(join(root, 'icon.svg'), join(outDir, 'icon.svg'))
await writeFile(join(outDir, 'index.json'), `${JSON.stringify(index, null, 2)}\n`, 'utf8')
console.log(`market/index.json written for ${manifest.name}@${manifest.version} at ${commit.slice(0, 12)}`)
