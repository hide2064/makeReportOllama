import React, { useEffect, useState } from 'react'
import HistoryPanel from './components/HistoryPanel'
import LoadingOverlay from './components/LoadingOverlay'
import ReferenceManager from './components/ReferenceManager'
import SlidePreview, { type SlideInfo } from './components/SlidePreview'
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
  const [status,          setStatus]          = useState<Status>('idle')
  const [step,            setStep]            = useState<string>('')
  const [progress,        setProgress]        = useState<number>(0)
  const [queuePosition,   setQueuePosition]   = useState<number>(0)
  const [errorMsg,        setErrorMsg]        = useState<string>('')
  const [jobId,           setJobId]           = useState<string>('')
  const [originalFilename, setOriginalFilename] = useState<string>('')
  const [historyKey,      setHistoryKey]      = useState(0)
  const [slidePreviews,   setSlidePreviews]   = useState<SlideInfo[] | null>(null)
  const [showSlides,      setShowSlides]      = useState(false)

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
    setProgress(0)
    setQueuePosition(0)
    setErrorMsg('')
    setJobId('')
    setSlidePreviews(null)
    setShowSlides(false)
    setOriginalFilename(params.excelFile.name)

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

      const { job_id } = await res.json()
      setJobId(job_id)

      // B-4: job_id でポーリング
      const deadline = Date.now() + POLL_TIMEOUT
      while (Date.now() < deadline) {
        await new Promise(resolve => setTimeout(resolve, POLL_INTERVAL))

        type ProgressData = {
          step: string; done: boolean; error: string
          progress: number; queue_position: number
        }
        let data: ProgressData | null = null
        try {
          const progressRes = await fetch(`/api/progress/${job_id}`, { signal: AbortSignal.timeout(10_000) })
          if (progressRes.ok) data = await progressRes.json()
        } catch { continue }
        if (!data) continue

        if (data.error) throw new Error(data.error)

        // B-4: キュー待機中 or 実行中
        const pos = data.queue_position ?? 0
        setQueuePosition(pos)
        if (pos > 0) {
          setStep(`キューで待機中... (${pos} 番目)`)
          setProgress(0)
        } else {
          if (data.step)     setStep(data.step)
          if (data.progress !== undefined) setProgress(data.progress)
        }

        if (data.done) {
          setProgress(100)
          setQueuePosition(0)
          setStep('完了しました！')
          setStatus('success')
          setHistoryKey(k => k + 1)
          sendNotification('レポート生成完了', '売上報告書の生成が完了しました。ダウンロードできます。')
          return
        }
      }
      throw new Error('処理がタイムアウトしました（20分）。')

    } catch (e: unknown) {
      setErrorMsg(e instanceof Error ? e.message : String(e))
      setStatus('error')
      sendNotification('レポート生成エラー', e instanceof Error ? e.message : '生成中にエラーが発生しました。')
    }
  }

  // B-3: 日付付きファイル名でダウンロード
  const handleDownload = () => {
    const a = document.createElement('a')
    a.href = `/api/download/${jobId}`
    const base = originalFilename.replace(/\.[^.]+$/, '') || 'report'
    const today = new Date().toISOString().slice(0, 10).replace(/-/g, '')
    a.download = `report_${base}_${today}.pptx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
  }

  // B-5: スライドプレビュー
  const handleToggleSlides = async () => {
    if (showSlides) {
      setShowSlides(false)
      return
    }
    setShowSlides(true)
    if (slidePreviews !== null) return  // キャッシュ済み
    try {
      const res = await fetch(`/api/slides/${jobId}`, { signal: AbortSignal.timeout(15_000) })
      if (res.ok) {
        const data = await res.json()
        setSlidePreviews(data.slides ?? [])
      } else {
        setSlidePreviews([])
      }
    } catch {
      setSlidePreviews([])
    }
  }

  return (
    <div className="app-wrapper">
      {status === 'loading' && (
        <LoadingOverlay step={step} progress={progress} queuePosition={queuePosition} />
      )}

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
            <div className="result-buttons">
              <button className="btn-download" onClick={handleDownload}>
                ⬇ report_{originalFilename.replace(/\.[^.]+$/, '')}_{new Date().toISOString().slice(0,10).replace(/-/g,'')}.pptx
              </button>
              <button className="btn-preview-toggle" onClick={handleToggleSlides}>
                {showSlides ? '▲ プレビューを閉じる' : '▼ スライドプレビュー'}
              </button>
            </div>
            {showSlides && slidePreviews === null && (
              <p className="slide-loading">スライドを読み込み中...</p>
            )}
            {showSlides && slidePreviews !== null && slidePreviews.length === 0 && (
              <p className="slide-loading">スライドの読み込みに失敗しました。</p>
            )}
            {showSlides && slidePreviews && slidePreviews.length > 0 && (
              <SlidePreview slides={slidePreviews} onClose={() => setShowSlides(false)} />
            )}
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
