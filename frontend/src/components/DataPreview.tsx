/**
 * DataPreview.tsx
 * アップロードされたファイルの先頭行プレビュー、必須列チェック、
 * およびブラウザ内ミニチャートを表示する。
 */
import React, { useState } from 'react'
import './DataPreview.css'

interface ChartData {
  product_totals: Record<string, number>
  monthly_totals: Record<string, number>
  total_amount:   number
  total_qty:      number
  period:         string
}

interface Props {
  columns:     string[]
  rows:        unknown[][]
  missingCols: string[]
  chartData?:  ChartData | null
}

const REQUIRED_COLS = ['日付', '商品名', '担当者', '地域', '数量', '売上金額']

function fmtAmount(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(1)}M円`
  if (n >= 1_000)     return `${(n / 1_000).toFixed(0)}K円`
  return `${n}円`
}

const MiniBarChart: React.FC<{ data: Record<string, number>; label: string }> = ({ data, label }) => {
  const entries = Object.entries(data)
  const max     = Math.max(...entries.map(([, v]) => v), 1)
  return (
    <div className="mini-chart">
      <p className="mini-chart-label">{label}</p>
      <div className="mini-chart-bars">
        {entries.map(([name, val]) => (
          <div key={name} className="mini-bar-row">
            <span className="mini-bar-name" title={name}>{name}</span>
            <div className="mini-bar-track">
              <div
                className="mini-bar-fill"
                style={{ width: `${(val / max) * 100}%` }}
              />
            </div>
            <span className="mini-bar-value">{fmtAmount(val)}</span>
          </div>
        ))}
      </div>
    </div>
  )
}

const MonthlyChart: React.FC<{ data: Record<string, number> }> = ({ data }) => {
  const entries = Object.entries(data)
  const max     = Math.max(...entries.map(([, v]) => v), 1)
  return (
    <div className="mini-chart">
      <p className="mini-chart-label">月次売上（直近 6 ヶ月）</p>
      <div className="monthly-bars">
        {entries.map(([month, val]) => (
          <div key={month} className="monthly-col">
            <span className="monthly-val">{fmtAmount(val)}</span>
            <div className="monthly-bar-track">
              <div
                className="monthly-bar-fill"
                style={{ height: `${Math.max((val / max) * 100, 4)}%` }}
              />
            </div>
            <span className="monthly-label">{month.slice(5)}</span>
          </div>
        ))}
      </div>
    </div>
  )
}

const DataPreview: React.FC<Props> = ({ columns, rows, missingCols, chartData }) => {
  const [showCharts, setShowCharts] = useState(false)

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

      {/* チャート集計サマリー */}
      {chartData && (
        <div className="chart-summary">
          <div className="chart-summary-stats">
            <div className="summary-stat">
              <span className="summary-stat-label">集計期間</span>
              <span className="summary-stat-value">{chartData.period}</span>
            </div>
            <div className="summary-stat">
              <span className="summary-stat-label">総売上</span>
              <span className="summary-stat-value primary">{fmtAmount(chartData.total_amount)}</span>
            </div>
            <div className="summary-stat">
              <span className="summary-stat-label">総販売数</span>
              <span className="summary-stat-value">{chartData.total_qty.toLocaleString()} 個</span>
            </div>
          </div>
          <button
            type="button"
            className="btn-toggle-charts"
            onClick={() => setShowCharts(v => !v)}
          >
            {showCharts ? '▲ チャートを閉じる' : '▼ データチャートを表示'}
          </button>
          {showCharts && (
            <div className="charts-panel">
              <MiniBarChart data={chartData.product_totals} label="商品別売上 TOP 8" />
              <MonthlyChart data={chartData.monthly_totals} />
            </div>
          )}
        </div>
      )}

      {/* データテーブル */}
      <div className="preview-table-wrap">
        <table className="preview-table">
          <thead>
            <tr>
              {columns.map(col => (
                <th key={col} className={!missingCols.includes(col) && REQUIRED_COLS.includes(col) ? 'col-required' : ''}>
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => (
              <tr key={i}>
                {(row as unknown[]).map((val, j) => (
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
