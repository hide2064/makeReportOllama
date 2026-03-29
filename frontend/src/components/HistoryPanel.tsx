/**
 * HistoryPanel.tsx
 * 生成済みレポートの履歴一覧を表示し、各ジョブを再ダウンロードできるコンポーネント。
 */
import React, { useCallback, useEffect, useState } from 'react'
import './HistoryPanel.css'

interface HistoryItem {
  job_id:            string
  created_at:        string
  original_filename: string
  analyst_model:     string
  writer_model:      string
}

interface Props {
  refreshKey?: number   // 外部から increment するとリストを再取得する
}

const HistoryPanel: React.FC<Props> = ({ refreshKey }) => {
  const [items, setItems]     = useState<HistoryItem[]>([])
  const [loading, setLoading] = useState(false)

  const loadHistory = useCallback(async () => {
    setLoading(true)
    try {
      const res = await fetch('/api/history?n=20', { signal: AbortSignal.timeout(10_000) })
      if (res.ok) {
        const data = await res.json()
        setItems(data.history ?? [])
      }
    } catch { /* ignore */ } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => { loadHistory() }, [loadHistory, refreshKey])

  const handleDownload = (item: HistoryItem) => {
    const a = document.createElement('a')
    a.href = `/api/history/${item.job_id}/download`
    a.download = `report_${item.job_id}.pptx`
    a.click()
  }

  const formatDate = (iso: string) => {
    try {
      return new Date(iso).toLocaleString('ja-JP', {
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit',
      })
    } catch {
      return iso
    }
  }

  return (
    <div className="history-panel">
      <div className="history-header">
        <h2>生成履歴</h2>
        <button className="btn-refresh" onClick={loadHistory} title="更新">↻</button>
      </div>

      {loading && <p className="history-loading">読み込み中...</p>}

      {!loading && items.length === 0 && (
        <p className="history-empty">生成済みのレポートはありません。</p>
      )}

      {items.length > 0 && (
        <ul className="history-list">
          {items.map(item => (
            <li key={item.job_id} className="history-item">
              <div className="history-item-info">
                <span className="history-filename">{item.original_filename}</span>
                <span className="history-date">{formatDate(item.created_at)}</span>
                {item.analyst_model && (
                  <span className="history-models">
                    {item.analyst_model} / {item.writer_model}
                  </span>
                )}
              </div>
              <button
                className="btn-redownload"
                onClick={() => handleDownload(item)}
                title="再ダウンロード"
              >
                ↓ DL
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}

export default HistoryPanel
