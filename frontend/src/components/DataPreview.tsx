/**
 * DataPreview.tsx
 * アップロードされたファイルの先頭行プレビューと必須列チェック結果を表示する。
 */
import React from 'react'
import './DataPreview.css'

interface Props {
  columns:     string[]
  rows:        unknown[][]
  missingCols: string[]
}

const REQUIRED_COLS = ['日付', '商品名', '担当者', '地域', '数量', '売上金額']

const DataPreview: React.FC<Props> = ({ columns, rows, missingCols }) => {
  return (
    <div className="data-preview">
      {/* 必須列チェック */}
      <div className="preview-badges">
        {REQUIRED_COLS.map(col => (
          <span
            key={col}
            className={`preview-badge ${missingCols.includes(col) ? 'badge-missing' : 'badge-ok'}`}
          >
            {missingCols.includes(col) ? '✗' : '✓'} {col}
          </span>
        ))}
      </div>

      {missingCols.length > 0 && (
        <p className="preview-warn">
          不足列: <strong>{missingCols.join('、')}</strong> — 必須列が揃うまでレポートを生成できません。
        </p>
      )}

      {/* データテーブル */}
      <div className="preview-table-wrap">
        <table className="preview-table">
          <thead>
            <tr>
              {columns.map(col => (
                <th key={col} className={missingCols.includes(col) ? '' : REQUIRED_COLS.includes(col) ? 'col-required' : ''}>
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => (
              <tr key={i}>
                {row.map((val, j) => (
                  <td key={j}>{String(val)}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="preview-note">先頭 {rows.length} 行を表示</p>
    </div>
  )
}

export default DataPreview
