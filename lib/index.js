// 打印文档插件：把 printable-doc 技能部署到 $DSH_HOME/skills/，并给出环境自检。
// 技能本体在 payload/ 里；这里只负责「搬过去 + 告诉用户还缺什么」。
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { spawnSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { ensureDeployed } from './deploy.js'

export const name = 'dsh-go-printable-doc'

const PLUGIN_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const PAYLOAD_DIR = path.join(PLUGIN_DIR, 'payload')
const DSH_HOME = process.env.DSH_HOME || path.join(os.homedir(), '.dsh')
const TARGET_DIR = process.env.DSH_SKILL_DIR || path.join(DSH_HOME, 'skills', 'printable-doc')

const PKG_VERSION = (() => {
  try {
    return JSON.parse(fs.readFileSync(path.join(PLUGIN_DIR, 'package.json'), 'utf8')).version
  } catch {
    return '?'
  }
})()

function runVersion () {
  try {
    return fs.readFileSync(path.join(TARGET_DIR, 'VERSION'), 'utf8').trim()
  } catch {
    return '-'
  }
}

let lastDeploy = null

function deploy () {
  try {
    lastDeploy = ensureDeployed({ payloadDir: PAYLOAD_DIR, targetDir: TARGET_DIR, warn: (m) => console.warn(m) })
  } catch (e) {
    lastDeploy = { action: 'error', message: e && e.message }
    console.warn('[printable-doc] 部署失败：' + (e && e.message))
  }
  return lastDeploy
}

function shell (cmd, args) {
  const r = spawnSync(cmd, args, { encoding: 'utf8', timeout: 4000 })
  return !r.error && r.status === 0
}

// 技能跑起来需要的东西。缺了不阻止安装，但要说清楚缺哪一项、怎么补。
function selfCheck () {
  return [
    {
      id: 'python',
      label: 'python3',
      ok: shell('sh', ['-c', 'command -v python3 >/dev/null']),
      hint: 'apt-get install -y python3',
    },
    {
      id: 'docx',
      label: 'python-docx',
      ok: shell('python3', ['-c', 'import docx']),
      hint: 'pip install python-docx',
    },
    {
      id: 'weasy',
      label: 'weasyprint（预览）',
      ok: shell('python3', ['-c', 'import weasyprint']),
      hint: 'pip install weasyprint　没有也能排版，只是不渲染预览图',
    },
    {
      id: 'pdf',
      label: 'pdftoppm（预览）',
      ok: shell('sh', ['-c', 'command -v pdftoppm >/dev/null']),
      hint: 'apt-get install -y poppler-utils',
    },
    {
      id: 'font',
      label: '中文字体',
      ok: shell('sh', ['-c', 'fc-list :lang=zh | head -1 | grep -q .']),
      hint: 'apt-get install -y fonts-noto-cjk',
    },
  ]
}

export function apply (ctx) {
  deploy()

  ctx.inject(['dshGo'], (ctx) => {
    // 技能版本落在磁盘上，包版本来自 npm 包本身。两个数不一致 = 部署没跑或没重启引擎。
    ctx.dshGo.panel('printdoc', {
      view: (scope, caps) => {
        const en = String((caps && caps.locale) || '').startsWith('en')
        const t = (zh, eng) => (en ? eng : zh)
        return {
          v: 2,
          type: 'stack',
          title: t('打印文档', 'Printable doc'),
          children: [
            {
              type: 'metric',
              label: t('已装技能版本', 'Skill version'),
              value: { $: 'data.skillVersion' },
              tone: { $: 'data.versionTone' },
              caption: { $: 'data.versionCaption' },
            },
            { type: 'text', text: { $: 'data.pathLine' }, tone: 'muted' },
            { type: 'text', text: { $: 'data.checksSummary' }, tone: 'muted' },
            { type: 'keyValue', items: { $: 'data.checks' } },
            { type: 'callout', text: { $: 'data.tip' }, tone: { $: 'data.tipTone' } },
            {
              type: 'actions',
              items: [
                { label: t('重新部署', 'Redeploy'), action: 'redeploy', icon: 'refresh' },
              ],
            },
            {
              type: 'list',
              items: [
                {
                  title: t('怎么用', 'How to use'),
                  subtitle: t('新开一个对话，说「把这篇整理成可直接打印的 Word」。技能在新会话创建时加载，装完要先新开对话。', 'Start a new chat and ask for a print-ready Word. Skills load when a chat starts.'),
                },
                {
                  title: t('装在哪', 'Installed at'),
                  subtitle: TARGET_DIR,
                },
                {
                  title: t('怎么卸', 'Removing'),
                  subtitle: t('卸载插件不会删技能，要清掉得手动删上面这个目录。', 'Removing the plugin keeps the skill; delete the folder above to clear it.'),
                },
              ],
            },
            { type: 'text', text: t(`包版本 ${PKG_VERSION}`, `Package ${PKG_VERSION}`), tone: 'muted' },
          ],
        }
      },
      data: (scope, emit, caps) => {
        const en = String((caps && caps.locale) || '').startsWith('en')
        const t = (zh, eng) => (en ? eng : zh)
        const installed = fs.existsSync(path.join(TARGET_DIR, 'SKILL.md'))
        const diskVersion = runVersion()
        const checks = selfCheck()
        const bad = checks.filter((c) => c.ok === false)
        const onlyOptional = bad.every((c) => c.id === 'weasy' || c.id === 'pdf')

        let tip
        let tipTone = 'info'
        if (!installed) {
          tip = t('技能还没部署：点下面的「重新部署」，或重启一次引擎。', 'Skill not deployed yet - tap Redeploy, or restart the engine.')
          tipTone = 'bad'
        } else if (!bad.length) {
          tip = t('就绪。新开一个对话说「整理成可直接打印的 Word」就行。', 'Ready. Start a new chat and ask for a print-ready Word.')
          tipTone = 'good'
        } else if (onlyOptional) {
          tip = t('能生成 docx，但缺预览工具：排版后看不到效果图，靠 check_docx 校验结构。', 'Can build the docx, but preview is unavailable without weasyprint/pdftoppm.')
          tipTone = 'warn'
        } else {
          tip = t('还差：', 'Missing: ') + bad.map((c) => c.label + '（' + c.hint + '）').join('；')
          tipTone = 'warn'
        }

        emit({
          skillVersion: installed ? diskVersion : t('未安装', 'not installed'),
          versionTone: !installed ? 'bad' : (diskVersion === PKG_VERSION ? 'good' : 'warn'),
          versionCaption: installed
            ? t(`包版本 ${PKG_VERSION}`, `pack ${PKG_VERSION}`)
            : t(`包版本 ${PKG_VERSION}`, `pack ${PKG_VERSION}`),
          pathLine: TARGET_DIR,
          checksSummary: bad.length ? t('环境自检：', 'Checks: ') : t('环境自检：全部正常', 'Checks: all good'),
          checks: checks.map((c) => ({
            label: (c.ok ? '✓ ' : '✗ ') + c.label,
            value: c.ok ? t('已就绪', 'Ready') : c.hint,
            tone: c.ok ? 'good' : 'bad',
          })),
          tip,
          tipTone,
        })
      },
      // 面板按钮的动作写在这里，不是 ctx.dshGo.on()——后者只收清单里声明过的 id。
      actions: {
        redeploy: (scope, event) => {
          const en = String((event && event.locale) || '').startsWith('en')
          const r = deploy()
          if (r && r.action === 'error') return { toast: (en ? 'Deploy failed: ' : '部署失败：') + r.message, refresh: true }
          if (r && r.action === 'current') return { toast: en ? 'Already up to date' : '已经是最新版本', refresh: true }
          return { toast: (en ? 'Deployed ' : '已部署 ') + (r && r.version ? r.version : ''), refresh: true }
        },
      },
    })
  })
}
