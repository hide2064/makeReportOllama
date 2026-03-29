/**
 * ReferenceManager.tsx
 * 過去レポート PPTX を RAG 用に登録・削除・一覧表示するコンポーネント。
 * ファイルをクリックすると抽出チャンク（参考情報）を右パネルに表示する。
 */
import React, { useCallback, useEffect, useRef, useState } from 'react'
import './ReferenceManager.css'

interface Reference {
  filename: string
  file_id:  string
  chunks:   number
}

interface Chunk {
  id:        string
  text:      string
  chunk_idx: number
}

const ReferenceManager: React.FC = () => {
  const [refs, setRefs]               = useState<Reference[]>([])
  const [uploading, setUploading]     = useState(false)
  const [msg, setMsg]                 = useState<{ text: string; ok: boolean } | null>(null)
  const [selectedRef, setSelectedRef] = useState<Reference | null>(null)
  const [chunks, setChunks]           = useState<Chunk[]>([])
  const [loadingChunks, setLoadingChunks] = useState(false)
  const fileInputRef                  = useRef<HTMLInputElement>(null)

  const loadRefs = useCallback(async () => {
    try {
      const res = await fetch('/api/references', { signal: AbortSignal.timeout(10_000) })
      if (res.ok) {
        const data = await res.json()
        setRefs(data.references ?? [])
      }
    } catch { /* ignore */ }
  }, [])

  useEffect(() => { loadRefs() }, [loadRefs])

  const handleUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setUploading(true)
    setMsg(null)
    try {
      const form = new FormData()
      form.append('file', file)
      const res = await fetch('/api/references/upload', {
        method: 'POST',
        body: form,
        signal: AbortSignal.timeout(120_000),
      })
      const data = await res.json()
      if (res.ok) {
        setMsg({ text: `「${data.filename}」を登録しました（${data.chunks} チャンク）`, ok: true })
        await loadRefs()
      } else {
        setMsg({ text: data.detail ?? 'アップロードに失敗しました', ok: false })
      }
    } catch (err) {
      setMsg({ text: err instanceof Error ? err.message : '通信エラー', ok: false })
    } finally {
      setUploading(false)
      if (fileInputRef.current) fileInputRef.current.value = ''
    }
  }

  const handleDelete = async (ref: Reference) => {
    if (!confirm(`「${ref.filename}」を削除しますか？`)) return
    try {
      const res = await fetch(`/api/references/${ref.file_id}`, {
        method: 'DELETE',
        signal: AbortSignal.timeout(10_000),
      })
      if (res.ok) {
        setMsg({ text: `「${ref.filename}」を削除しました`, ok: true })
        if (selectedRef?.file_id === ref.file_id) {
          setSelectedRef(null)
          setChunks([])
        }
        await loadRefs()
      } else {
        const data = await res.json()
        setMsg({ text: data.detail ?? '削除に失敗しました', ok: false })
      }
    } catch (err) {
      setMsg({ text: err instanceof Error ? err.message : '通信エラー', ok: false })
    }
  }

  const handleSelectRef = async (ref: Reference) => {
    if (selectedRef?.file_id === ref.file_id) {
      setSelectedRef(null)
      setChunks([])
      return
    }
    setSelectedRef(ref)
    setChunks([])
    setLoadingChunks(true)
    try {
      const res = await fetch(`/api/references/${ref.file_id}/chunks`, {
        signal: AbortSignal.timeout(10_000),
      })
      if (res.ok) {
        const data = await res.json()
        setChunks(data.chunks ?? [])
      }
    } catch { /* ignore */ }
    finally { setLoadingChunks(false) }
  }

  return (
    <div className="ref-manager">
      <div className="ref-header">
        <div className="ref-title-wrap">
          <span className="ref-badge">RAG</span>
          <h2>過去レポート管理</h2>
        </div>
        <p className="ref-desc">
          過去の報告書 PPTX を登録しておくと、新規レポート生成時に類似の文脈を自動参照して文章品質が向上します。
          ファイル名をクリックすると抽出済み参考情報を確認できます。
        </p>
      </div>

      <label className={`btn-upload-label ${uploading ? 'disabled' : ''}`}>
        <input
          ref={fileInputRef}
          type="file"
          accept=".pptx"
          onChange={handleUpload}
          disabled={uploading}
          style={{ display: 'none' }}
        />
        {uploading ? '登録中...' : '＋ PPTX を登録'}
      </label>

      {msg && (
        <p className={`ref-msg ${msg.ok ? 'ref-msg-ok' : 'ref-msg-err'}`}>
          {msg.text}
        </p>
      )}

      <div className="ref-body">
        {/* 左: ファイル一覧 */}
        <div className="ref-list-col">
          {refs.length === 0 ? (
            <p className="ref-empty">登録済みの過去レポートはありません。</p>
          ) : (
            <ul className="ref-list">
              {refs.map(r => (
                <li
                  key={r.file_id}
                  className={`ref-item ${selectedRef?.file_id === r.file_id ? 'ref-item-active' : ''}`}
                >
                  <button
                    className="ref-item-btn"
                    onClick={() => handleSelectRef(r)}
                    title="クリックで参考情報を表示"
                  >
                    <div className="ref-item-info">
                      <span className="ref-item-name">{r.filename}</span>
                      <span className="ref-item-chunks">{r.chunks} チャンク</span>
                    </div>
                    <span className="ref-item-arrow">
                      {selectedRef?.file_id === r.file_id ? '▲' : '▼'}
                    </span>
                  </button>
                  <button className="btn-delete" onClick={() => handleDelete(r)}>削除</button>
                </li>
              ))}
            </ul>
          )}
        </div>

        {/* 右: チャンクパネル */}
        {selectedRef && (
          <div className="ref-chunks-panel">
            <div className="ref-chunks-header">
              <span className="ref-chunks-title">
                参考情報 — {selectedRef.filename}
              </span>
              <button
                className="ref-chunks-close"
                onClick={() => { setSelectedRef(null); setChunks([]) }}
              >
                ✕
              </button>
            </div>
            <p className="ref-chunks-hint">
              以下のテキストが RAG コンテキストとして新規レポート生成時に参照されます。
            </p>
            {loadingChunks ? (
              <p className="ref-chunks-loading">読み込み中...</p>
            ) : chunks.length === 0 ? (
              <p className="ref-chunks-empty">チャンクが見つかりません。</p>
            ) : (
              <ol className="ref-chunks-list">
                {chunks.map((c, i) => (
                  <li key={c.id} className="ref-chunk-item">
                    <span className="ref-chunk-idx">スライド {i + 1}</span>
                    <pre className="ref-chunk-text">{c.text}</pre>
                  </li>
                ))}
              </ol>
            )}
          </div>
        )}
      </div>
    </div>
  )
}

export default ReferenceManager
