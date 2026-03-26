import React, { useState } from 'react'
import LoadingOverlay from './components/LoadingOverlay'
import ReferenceManager from './components/ReferenceManager'
import UploadForm from './components/UploadForm'
import './App.css'

type Status = 'idle' | 'loading' | 'success' | 'error'

const POLL_INTERVAL = 2000   // 2 秒ごとにポーリング
const POLL_TIMEOUT  = 1_200_000  // 最大 20 分

const App: React.FC = () => {
  const [status, setStatus]           = useState<Status>('idle')
  const [step, setStep]               = useState<string>('')
  const [errorMsg, setErrorMsg]       = useState<string>('')
  const [downloadUrl, setDownloadUrl] = useState<string>('')

  const handleGenerate = async (excel: File, template: File) => {
    setStatus('loading')
    setStep('アップロード中...')
    setErrorMsg('')
    setDownloadUrl('')

    const form = new FormData()
    form.append('excel_file', excel)
    form.append('template_file', template)

    try {
      // ── 1. 処理開始リクエスト ──────────────────────────
      const res = await fetch('/api/generate', {
        method: 'POST',
        body: form,
        signal: AbortSignal.timeout(30_000),
      })
      if (!res.ok) {
        const err = await res.json().catch(() => ({ detail: 'Unknown error' }))
        throw new Error(err.detail ?? `HTTP ${res.status}`)
      }

      // ── 2. 進捗ポーリング ──────────────────────────────
      const deadline = Date.now() + POLL_TIMEOUT
      while (Date.now() < deadline) {
        await new Promise(resolve => setTimeout(resolve, POLL_INTERVAL))

        // ネットワーク瞬断（サーバー再起動など）は無視してリトライ
        let data: { step: string; done: boolean; error: string } | null = null
        try {
          const progressRes = await fetch('/api/progress', {
            signal: AbortSignal.timeout(10_000),
          })
          if (progressRes.ok) {
            data = await progressRes.json()
          }
        } catch {
          // 瞬断 → 次の poll まで待つ
          continue
        }
        if (!data) continue

        if (data.error) throw new Error(data.error)
        if (data.step)  setStep(data.step)

        if (data.done) {
          // ── 3. ダウンロード ─────────────────────────────
          setStep('完了しました。ファイルをダウンロードしています...')
          const dlRes = await fetch('/api/download', {
            signal: AbortSignal.timeout(30_000),
          })
          if (!dlRes.ok) throw new Error('ダウンロードに失敗しました')
          const blob = await dlRes.blob()
          setDownloadUrl(URL.createObjectURL(blob))
          setStatus('success')
          return
        }
      }

      throw new Error('処理がタイムアウトしました（20分）。')

    } catch (e: unknown) {
      setErrorMsg(e instanceof Error ? e.message : String(e))
      setStatus('error')
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
      </header>

      <main className="app-main">
        <section className="card">
          <h2>ファイルをアップロード</h2>
          <UploadForm onGenerate={handleGenerate} disabled={status === 'loading'} />
        </section>

        <section className="card">
          <ReferenceManager />
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
      </main>
    </div>
  )
}

export default App
