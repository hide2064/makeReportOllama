import React, { useEffect, useState } from 'react'
import HistoryPanel from './components/HistoryPanel'
import LoadingOverlay from './components/LoadingOverlay'
import ReferenceManager from './components/ReferenceManager'
import UploadForm, { type GenerateParams } from './components/UploadForm'
import './App.css'

type Status = 'idle' | 'loading' | 'success' | 'error'

const POLL_INTERVAL = 2000
const POLL_TIMEOUT  = 1_200_000

// ── L-1: ダークモード ────────────────────────────────────────
function getInitialTheme(): 'light' | 'dark' {
  const saved = localStorage.getItem('theme')
  if (saved === 'dark' || saved === 'light') return saved
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

// ── L-3: ブラウザ通知 ────────────────────────────────────────
async function requestNotificationPermission(): Promise<boolean> {
  if (!('Notification' in window)) return false
  if (Notification.permission === 'granted') return true
  if (Notification.permission === 'denied') return false
  const result = await Notification.requestPermission()
  return result === 'granted'
}

function sendNotification(title: string, body: string) {
  if (!('Notification' in window) || Notification.permission !== 'granted') return
  new Notification(title, { body, icon: '/favicon.svg' })
}

const App: React.FC = () => {
  const [status,      setStatus]      = useState<Status>('idle')
  const [step,        setStep]        = useState<string>('')
  const [errorMsg,    setErrorMsg]    = useState<string>('')
  const [downloadUrl, setDownloadUrl] = useState<string>('')
  const [historyKey,  setHistoryKey]  = useState(0)

  // L-1: テーマ状態
  const [theme, setTheme] = useState<'light' | 'dark'>(getInitialTheme)

  useEffect(() => {
    document.documentElement.classList.toggle('dark', theme === 'dark')
    localStorage.setItem('theme', theme)
  }, [theme])

  const toggleTheme = () => setTheme(t => t === 'light' ? 'dark' : 'light')

  const handleGenerate = async (params: GenerateParams) => {
    setStatus('loading')
    setStep('アップロード中...')
    setErrorMsg('')
    setDownloadUrl('')

    // L-3: 処理開始時に通知権限をリクエスト
    await requestNotificationPermission()

    const form = new FormData()
    form.append('excel_file',        params.excelFile)
    if (params.templateFile) form.append('template_file', params.templateFile)
    form.append('template_name',        params.templateName)
    form.append('slide_product_table',  String(params.slideProductTable))
    form.append('slide_region_table',   String(params.slideRegionTable))
    form.append('slide_rep_table',      String(params.slideRepTable))
    form.append('slide_chart',          String(params.slideChart))
    form.append('chart_product_type',   params.chartProductType)
    form.append('analyst_model',        params.analystModel)
    form.append('writer_model',         params.writerModel)
    form.append('date_from',            params.dateFrom)
    form.append('date_to',              params.dateTo)
    form.append('extra_context',        params.extraContext)

    try {
      const res = await fetch('/api/generate', {
        method: 'POST',
        body: form,
        signal: AbortSignal.timeout(30_000),
      })
      if (!res.ok) {
        const err = await res.json().catch(() => ({ detail: 'Unknown error' }))
        throw new Error(err.detail ?? `HTTP ${res.status}`)
      }

      const deadline = Date.now() + POLL_TIMEOUT
      while (Date.now() < deadline) {
        await new Promise(resolve => setTimeout(resolve, POLL_INTERVAL))

        let data: { step: string; done: boolean; error: string } | null = null
        try {
          const progressRes = await fetch('/api/progress', { signal: AbortSignal.timeout(10_000) })
          if (progressRes.ok) data = await progressRes.json()
        } catch { continue }
        if (!data) continue

        if (data.error) throw new Error(data.error)
        if (data.step)  setStep(data.step)

        if (data.done) {
          setStep('完了しました。ファイルをダウンロードしています...')
          const dlRes = await fetch('/api/download', { signal: AbortSignal.timeout(30_000) })
          if (!dlRes.ok) throw new Error('ダウンロードに失敗しました')
          const blob = await dlRes.blob()
          setDownloadUrl(URL.createObjectURL(blob))
          setStatus('success')
          setHistoryKey(k => k + 1)
          // L-3: 完了通知
          sendNotification('レポート生成完了', '売上報告書の生成が完了しました。ダウンロードできます。')
          return
        }
      }
      throw new Error('処理がタイムアウトしました（20分）。')

    } catch (e: unknown) {
      setErrorMsg(e instanceof Error ? e.message : String(e))
      setStatus('error')
      // L-3: エラー通知
      sendNotification('レポート生成エラー', e instanceof Error ? e.message : '生成中にエラーが発生しました。')
    }
  }

  const handleDownload = () => {
    const a = document.createElement('a')
    a.href = downloadUrl
    a.download = 'report.pptx'
    a.click()
  }

  return (
    <div className="app-wrapper">
      {status === 'loading' && <LoadingOverlay step={step} />}

      <header className="app-header">
        <h1>売上報告書 自動生成</h1>
        <p className="app-subtitle">
          Excel / CSV 売上データと PPTX テンプレートをアップロードすると、
          ローカル LLM (Ollama) が分析して PowerPoint 報告書を生成します。
        </p>
        {/* L-1: ダークモードトグル */}
        <button className="theme-toggle" onClick={toggleTheme} aria-label="テーマ切り替え">
          {theme === 'dark' ? '☀️ ライト' : '🌙 ダーク'}
        </button>
      </header>

      <main className="app-main">
        <section className="card">
          <h2>ファイルをアップロード</h2>
          <UploadForm onGenerate={handleGenerate} disabled={status === 'loading'} />
        </section>

        {status === 'success' && (
          <section className="card result-card">
            <p className="result-ok">✓ 報告書の生成が完了しました！</p>
            <button className="btn-download" onClick={handleDownload}>
              report.pptx をダウンロード
            </button>
          </section>
        )}

        {status === 'error' && (
          <section className="card error-card">
            <p className="result-error">エラーが発生しました</p>
            <pre className="error-detail">{errorMsg}</pre>
          </section>
        )}

        <section className="card">
          <HistoryPanel refreshKey={historyKey} />
        </section>

        <section className="card">
          <ReferenceManager />
        </section>
      </main>
    </div>
  )
}

export default App
