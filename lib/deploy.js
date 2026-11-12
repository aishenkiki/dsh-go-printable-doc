// 把 payload 里的技能文件部署到 $DSH_HOME/skills/printable-doc。
// 只在引擎启动、插件加载时跑一次；裁决依据只有目标目录 VERSION 与 payload VERSION 的比较。
// 不用 fs.cpSync：proot / 某些安卓内核下 copy_file_range 会返回 EINVAL。
import fs from 'node:fs'
import path from 'node:path'

// payload 里属于「技能本体」的顶层条目。其余文件不碰。
export const RUNTIME = ['VERSION', 'SKILL.md', 'references', 'scripts']

// 技能目录没有用户数据（不像题库），升级不迁移任何文件；旧目录整体改名成备份。
const MAX_BACKUPS = 2

export function readVersion (dir) {
  try {
    const v = fs.readFileSync(path.join(dir, 'VERSION'), 'utf8').trim()
    return v === '' ? null : v
  } catch {
    return null
  }
}

export function compareVersion (a, b) {
  const pa = String(a).split('.').map((n) => Number.parseInt(n, 10) || 0)
  const pb = String(b).split('.').map((n) => Number.parseInt(n, 10) || 0)
  for (let i = 0; i < Math.max(pa.length, pb.length); i += 1) {
    const d = (pa[i] ?? 0) - (pb[i] ?? 0)
    if (d !== 0) return d > 0 ? 1 : -1
  }
  return 0
}

function stamp () {
  const d = new Date()
  const p = (n) => String(n).padStart(2, '0')
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`
}

function copyTree (src, dst) {
  const st = fs.statSync(src)
  if (st.isDirectory()) {
    fs.mkdirSync(dst, { recursive: true })
    for (const name of fs.readdirSync(src)) copyTree(path.join(src, name), path.join(dst, name))
    return
  }
  fs.copyFileSync(src, dst)
}

function copyRuntime (payloadDir, tmpDir) {
  fs.rmSync(tmpDir, { recursive: true, force: true })
  fs.mkdirSync(tmpDir, { recursive: true })
  for (const name of RUNTIME) {
    const src = path.join(payloadDir, name)
    if (!fs.existsSync(src)) continue
    copyTree(src, path.join(tmpDir, name))
  }
}

// 备份会越攒越多，只留最近几次。只认「<技能名>.bak-<时间戳>」这个形状的目录。
function pruneBackups (targetDir, warn = () => {}) {
  const parent = path.dirname(targetDir)
  const base = path.basename(targetDir) + '.bak-'
  let names = []
  try {
    names = fs.readdirSync(parent).filter((n) => n.startsWith(base)).sort()
  } catch {
    return []
  }
  const stale = names.slice(0, Math.max(0, names.length - MAX_BACKUPS))
  for (const n of stale) {
    try {
      fs.rmSync(path.join(parent, n), { recursive: true, force: true })
    } catch (e) {
      warn(`[printable-doc] 清理旧备份 ${n} 失败：${e && e.message}`)
    }
  }
  return names.slice(-MAX_BACKUPS)
}

export function ensureDeployed ({ payloadDir, targetDir, warn = () => {} }) {
  const payloadVersion = readVersion(payloadDir)
  if (payloadVersion === null) {
    warn('[printable-doc] payload 缺少 VERSION，跳过部署')
    return { action: 'no-payload' }
  }

  const targetVersion = readVersion(targetDir)
  if (targetVersion === payloadVersion) {
    pruneBackups(targetDir, warn)
    return { action: 'current', version: payloadVersion }
  }
  if (targetVersion !== null && compareVersion(payloadVersion, targetVersion) <= 0) {
    warn(`[printable-doc] 目标版本 ${targetVersion} 不低于 payload ${payloadVersion}，不覆盖`)
    return { action: 'keep-target', version: targetVersion }
  }

  const tmpDir = `${targetDir}.tmp`
  copyRuntime(payloadDir, tmpDir)

  const existed = fs.existsSync(targetDir)
  if (existed) {
    // 同一秒内连续升级会让时间戳撞车，撞上就加序号，否则 rename 抛 ENOTEMPTY。
    let bak = `${targetDir}.bak-${stamp()}`
    for (let n = 1; fs.existsSync(bak); n += 1) bak = `${targetDir}.bak-${stamp()}-${n}`
    fs.renameSync(targetDir, bak)
    warn(`[printable-doc] 旧版本（${targetVersion ?? '无 VERSION'}）已备份到 ${bak}`)
  }
  fs.mkdirSync(path.dirname(targetDir), { recursive: true })
  fs.renameSync(tmpDir, targetDir)
  pruneBackups(targetDir, warn)
  return { action: existed ? 'upgrade' : 'deploy', version: payloadVersion }
}
