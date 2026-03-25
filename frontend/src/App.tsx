import React, { useState } from 'react'
import LoadingOverlay from './components/LoadingOverlay'
import UploadForm from './components/UploadForm'
import './App.css'

type Status = 'idle' | 'loading' | 'success' | 'error'

const App: React.FC = () => {
  const [status, setStatus]           = useState<Status>('idle')
  const [errorMsg, setErrorMsg]       = useState<string>('')
  const [downloadUrl, setDownloadUrl] = useState<string>('')

  const handleGenerate = async (excel: File, template: File) => {
    setStatus('loading')
    setErrorMsg('')
    setDownloadUrl('')

    const form = new FormData()
    form.append('excel_file', excel)
    form.append('template_file', template)

    try {
      const res = await fetch('/api/generate', {
        method: 'POST',
        body: form,
        signal: AbortSignal.timeout(360_000), // 6 分
      })

      if (!res.ok) {
        const err = await res.json().catch(() => ({ detail: 'Unknown error' }))
        throw new Error(err.detail ?? `HTTP ${res.status}`)
      }

      const blob = await res.blob()
      const url  = URL.createObjectURL(blob)
      setDownloadUrl(url)
      setStatus('success')
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e)
      setErrorMsg(msg)
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
      {status === 'loading' && <LoadingOverlay />}

      <header className="app-header">
        <h1>売上報告書 自動生成</h1>
        <p className="app-subtitle">
          Excel 売上データと PPTX テンプレートをアップロードすると、
          ローカル LLM (Ollama) が分析して PowerPoint 報告書を生成します。
        </p>
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
      </main>
    </div>
  )
}

export default App
