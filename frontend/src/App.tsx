import React, { useState } from 'react'
import LoadingOverlay from './components/LoadingOverlay'
import UploadForm from './components/UploadForm'
import './App.css'

type Status = 'idle' | 'loading' | 'success' | 'error'

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
      const res = await fetch('/api/generate', {
        method: 'POST',
        body: form,
        signal: AbortSignal.timeout(1_200_000), // 20 分
      })

      if (!res.ok || !res.body) {
        const err = await res.json().catch(() => ({ detail: 'Unknown error' }))
        throw new Error(err.detail ?? `HTTP ${res.status}`)
      }

      // SSE ストリームを読み込む
      const reader  = res.body.getReader()
      const decoder = new TextDecoder()
      let buffer    = ''

      while (true) {
        const { done, value } = await reader.read()
        if (done) break

        buffer += decoder.decode(value, { stream: true })
        const lines = buffer.split('\n')
        buffer = lines.pop() ?? ''

        for (const line of lines) {
          if (!line.startsWith('data: ')) continue
          const data = JSON.parse(line.slice(6)) as {
            step?: string
            error?: string
            done?: boolean
          }

          if (data.error) throw new Error(data.error)

          if (data.step) setStep(data.step)

          if (data.done) {
            setStep('完了しました。ファイルをダウンロードしています...')
            const dlRes = await fetch('/api/download', {
              signal: AbortSignal.timeout(30_000),
            })
            if (!dlRes.ok) throw new Error('ダウンロードに失敗しました')
            const blob = await dlRes.blob()
            const url  = URL.createObjectURL(blob)
            setDownloadUrl(url)
            setStatus('success')
          }
        }
      }
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
