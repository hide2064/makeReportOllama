import React, { useCallback, useEffect, useRef, useState } from 'react'
import DataPreview from './DataPreview'
import './UploadForm.css'

export interface GenerateParams {
  excelFile:         File
  templateFile:      File | null
  templateName:      string
  slideProductTable: boolean
  slideRegionTable:  boolean
  slideRepTable:     boolean
  slideChart:        boolean
  chartProductType:  'bar' | 'pie'
  analystModel:      string
  writerModel:       string
  dateFrom:          string
  dateTo:            string
  extraContext:      string
}

interface ChartData {
  product_totals: Record<string, number>
  monthly_totals: Record<string, number>
  total_amount:   number
  total_qty:      number
  period:         string
}

interface PreviewData {
  columns:      string[]
  rows:         unknown[][]
  missing_cols: string[]
  chart_data:   ChartData | null
}

interface ContextTemplate {
  id:   string
  name: string
  text: string
}

const TEMPLATES_KEY = 'extra_context_templates'

interface Props {
  onGenerate: (params: GenerateParams) => void
  disabled:   boolean
}

const EXCEL_EXTS = ['xlsx', 'xls', 'csv']

const UploadForm: React.FC<Props> = ({ onGenerate, disabled }) => {
  const [excelFile,        setExcelFile]        = useState<File | null>(null)
  const [templateFile,     setTemplateFile]     = useState<File | null>(null)
  const [templateMode,     setTemplateMode]     = useState<'upload' | 'server'>('upload')
  const [serverTemplates,  setServerTemplates]  = useState<{ name: string }[]>([])
  const [selectedTemplate, setSelectedTemplate] = useState<string>('')
  const [slideProductTable, setSlideProductTable] = useState(true)
  const [slideRegionTable,  setSlideRegionTable]  = useState(false)
  const [slideRepTable,     setSlideRepTable]     = useState(false)
  const [slideChart,        setSlideChart]        = useState(true)
  const [chartProductType,  setChartProductType]  = useState<'bar' | 'pie'>('bar')

  // H-3: モデル選択
  const [availableModels, setAvailableModels] = useState<string[]>([])
  const [analystModel,    setAnalystModel]    = useState<string>('')
  const [writerModel,     setWriterModel]     = useState<string>('')

  // H-5: 分析期間
  const [dateFrom,  setDateFrom]  = useState<string>('')
  const [dateTo,    setDateTo]    = useState<string>('')
  const [dateError, setDateError] = useState<string>('')

  // 追加プロンプト
  const [extraContext, setExtraContext] = useState<string>('')

  // B-6: テンプレート管理
  const [savedTemplates, setSavedTemplates] = useState<ContextTemplate[]>(() => {
    try {
      const raw = localStorage.getItem(TEMPLATES_KEY)
      return raw ? (JSON.parse(raw) as ContextTemplate[]) : []
    } catch {
      return []
    }
  })

  const saveTemplate = () => {
    const name = window.prompt('テンプレート名を入力してください')
    if (!name || !extraContext.trim()) return
    const newTpl: ContextTemplate = { id: Date.now().toString(), name, text: extraContext }
    const updated = [...savedTemplates, newTpl]
    setSavedTemplates(updated)
    localStorage.setItem(TEMPLATES_KEY, JSON.stringify(updated))
  }

  const loadTemplate = (id: string) => {
    const tpl = savedTemplates.find(t => t.id === id)
    if (tpl) setExtraContext(tpl.text)
  }

  const deleteTemplate = (id: string) => {
    if (!window.confirm('このテンプレートを削除しますか？')) return
    const updated = savedTemplates.filter(t => t.id !== id)
    setSavedTemplates(updated)
    localStorage.setItem(TEMPLATES_KEY, JSON.stringify(updated))
  }

  // H-2: プレビュー
  const [previewData,    setPreviewData]    = useState<PreviewData | null>(null)
  const [previewLoading, setPreviewLoading] = useState(false)
  const [previewError,   setPreviewError]   = useState<string>('')

  // H-1: ドラッグ＆ドロップ
  const [isDragOver, setIsDragOver] = useState(false)

  const excelRef    = useRef<HTMLInputElement>(null)
  const templateRef = useRef<HTMLInputElement>(null)

  // テンプレート一覧取得
  useEffect(() => {
    fetch('/api/templates')
      .then(r => r.json())
      .then(d => {
        const list: { name: string }[] = d.templates ?? []
        setServerTemplates(list)
        if (list.length > 0) setSelectedTemplate(list[0].name)
      })
      .catch(() => {})
  }, [])

  // H-3: モデル一覧取得
  useEffect(() => {
    fetch('/api/models')
      .then(r => r.json())
      .then(d => {
        setAvailableModels(d.available ?? [])
        setAnalystModel(d.current_analyst ?? '')
        setWriterModel(d.current_writer ?? '')
      })
      .catch(() => {})
  }, [])

  // H-2: プレビュー取得
  const fetchPreview = useCallback(async (file: File) => {
    setPreviewLoading(true)
    setPreviewError('')
    setPreviewData(null)
    try {
      const form = new FormData()
      form.append('file', file)
      const res = await fetch('/api/preview', {
        method: 'POST',
        body: form,
        signal: AbortSignal.timeout(30_000),
      })
      if (res.ok) {
        setPreviewData(await res.json())
      } else {
        const err = await res.json().catch(() => ({ detail: 'プレビュー取得に失敗しました' }))
        setPreviewError(err.detail ?? 'プレビュー取得に失敗しました')
      }
    } catch (e) {
      setPreviewError(e instanceof Error ? e.message : 'プレビュー取得に失敗しました')
    } finally {
      setPreviewLoading(false)
    }
  }, [])

  const handleExcelFile = useCallback((file: File | undefined) => {
    if (!file) return
    const ext = file.name.split('.').pop()?.toLowerCase() ?? ''
    if (!EXCEL_EXTS.includes(ext)) {
      setPreviewError('.xlsx / .xls / .csv ファイルのみ対応しています')
      return
    }
    setExcelFile(file)
    fetchPreview(file)
  }, [fetchPreview])

  const clearExcel = (e: React.MouseEvent) => {
    e.stopPropagation()
    setExcelFile(null)
    setPreviewData(null)
    setPreviewError('')
    if (excelRef.current) excelRef.current.value = ''
  }

  // H-1: ドラッグ＆ドロップ
  const handleDragEnter = (e: React.DragEvent) => { e.preventDefault(); if (!disabled) setIsDragOver(true) }
  const handleDragOver  = (e: React.DragEvent) => { e.preventDefault() }
  const handleDragLeave = (e: React.DragEvent) => { e.preventDefault(); setIsDragOver(false) }
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragOver(false)
    if (disabled) return
    handleExcelFile(e.dataTransfer.files[0])
  }

  // H-5: 日付バリデーション
  const handleDateFrom = (v: string) => {
    setDateFrom(v)
    setDateError(dateTo && v && v > dateTo ? '開始日は終了日より前の日付を指定してください' : '')
  }
  const handleDateTo = (v: string) => {
    setDateTo(v)
    setDateError(dateFrom && v && dateFrom > v ? '終了日は開始日より後の日付を指定してください' : '')
  }

  const hasMissingCols = previewData ? previewData.missing_cols.length > 0 : false
  const canSubmit = !!excelFile && !hasMissingCols && !dateError && (
    templateMode === 'upload' ? !!templateFile : !!selectedTemplate
  )

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    if (!excelFile) return
    onGenerate({
      excelFile,
      templateFile:      templateMode === 'upload' ? templateFile : null,
      templateName:      templateMode === 'server' ? selectedTemplate : '',
      slideProductTable,
      slideRegionTable,
      slideRepTable,
      slideChart,
      chartProductType,
      analystModel,
      writerModel,
      dateFrom,
      dateTo,
      extraContext,
    })
  }

  return (
    <form className="upload-form" onSubmit={handleSubmit}>

      {/* ── 売上データ (H-1: ドロップゾーン) ── */}
      <div className="form-group">
        <label>売上データ (.xlsx / .csv)</label>
        <div
          className={`drop-zone${isDragOver ? ' drag-over' : ''}${excelFile ? ' has-file' : ''}`}
          onDragEnter={handleDragEnter}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => !disabled && excelRef.current?.click()}
          style={disabled ? { pointerEvents: 'none', opacity: 0.5 } : {}}
        >
          <input
            ref={excelRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            style={{ display: 'none' }}
            onChange={e => handleExcelFile(e.target.files?.[0])}
            disabled={disabled}
          />
          {excelFile ? (
            <div className="drop-zone-file">
              <span className="drop-zone-filename">{excelFile.name}</span>
              <button type="button" className="drop-zone-clear" onClick={clearExcel}>×</button>
            </div>
          ) : (
            <div className="drop-zone-hint">
              <span className="drop-zone-icon">📂</span>
              <span>ファイルをドロップ、またはクリックして選択</span>
              <span className="drop-zone-sub">.xlsx / .xls / .csv</span>
            </div>
          )}
        </div>

        {/* H-2: プレビュー */}
        {previewLoading && <p className="preview-loading">プレビュー読み込み中...</p>}
        {previewError   && <p className="preview-error">{previewError}</p>}
        {previewData    && (
          <DataPreview
            columns={previewData.columns}
            rows={previewData.rows}
            missingCols={previewData.missing_cols}
            chartData={previewData.chart_data}
          />
        )}
      </div>

      {/* ── PPTX テンプレート ── */}
      <div className="form-group">
        <label>PPTX テンプレート (.pptx)</label>
        <div className="template-tabs">
          <button type="button" className={`tab-btn${templateMode === 'upload' ? ' active' : ''}`}
            onClick={() => setTemplateMode('upload')} disabled={disabled}>
            ファイルアップロード
          </button>
          <button type="button" className={`tab-btn${templateMode === 'server' ? ' active' : ''}`}
            onClick={() => setTemplateMode('server')} disabled={disabled}>
            サーバーのテンプレート
          </button>
        </div>

        <div style={{ display: templateMode === 'upload' ? 'flex' : 'none' }} className="file-row">
          <input
            ref={templateRef}
            type="file"
            accept=".pptx"
            style={{ display: 'none' }}
            onChange={e => setTemplateFile(e.target.files?.[0] ?? null)}
          />
          <button type="button" className="btn-secondary"
            onClick={() => templateRef.current?.click()} disabled={disabled}>
            ファイルを選択
          </button>
          <span className="file-name">{templateFile ? templateFile.name : '未選択'}</span>
        </div>

        {templateMode === 'server' && (
          <select className="template-select" value={selectedTemplate}
            onChange={e => setSelectedTemplate(e.target.value)} disabled={disabled}>
            {serverTemplates.length === 0
              ? <option value="">テンプレートなし</option>
              : serverTemplates.map(t => <option key={t.name} value={t.name}>{t.name}</option>)
            }
          </select>
        )}
      </div>

      {/* ── H-5: 分析期間 ── */}
      <div className="form-group">
        <label>分析期間（任意）</label>
        <div className="date-row">
          <input
            type="date"
            className="date-input"
            value={dateFrom}
            onChange={e => handleDateFrom(e.target.value)}
            disabled={disabled}
          />
          <span className="date-sep">〜</span>
          <input
            type="date"
            className="date-input"
            value={dateTo}
            onChange={e => handleDateTo(e.target.value)}
            disabled={disabled}
          />
        </div>
        {dateError && <p className="date-error">{dateError}</p>}
      </div>

      {/* ── H-3: AI モデル設定 ── */}
      <div className="form-group">
        <label>AI モデル設定</label>
        <div className="model-row">
          <div className="model-item">
            <span className="model-label">Analyst</span>
            {availableModels.length > 0 ? (
              <select className="model-select" value={analystModel}
                onChange={e => setAnalystModel(e.target.value)} disabled={disabled}>
                {availableModels.map(m => <option key={m} value={m}>{m}</option>)}
              </select>
            ) : (
              <span className="model-fixed">{analystModel || 'qwen2.5:3b'}</span>
            )}
          </div>
          <div className="model-item">
            <span className="model-label">Writer</span>
            {availableModels.length > 0 ? (
              <select className="model-select" value={writerModel}
                onChange={e => setWriterModel(e.target.value)} disabled={disabled}>
                {availableModels.map(m => <option key={m} value={m}>{m}</option>)}
              </select>
            ) : (
              <span className="model-fixed">{writerModel || 'qwen3:8b'}</span>
            )}
          </div>
        </div>
      </div>

      {/* ── スライド構成 ── */}
      <div className="form-group slide-options">
        <label>スライド構成</label>
        <div className="options-row">
          {[
            ['商品別売上表', slideProductTable, setSlideProductTable],
            ['地域別売上表', slideRegionTable,  setSlideRegionTable],
            ['担当者別売上表', slideRepTable,   setSlideRepTable],
            ['グラフスライド', slideChart,      setSlideChart],
          ].map(([label, checked, setter]) => (
            <label key={label as string} className="checkbox-label">
              <input type="checkbox" checked={checked as boolean}
                onChange={e => (setter as React.Dispatch<React.SetStateAction<boolean>>)(e.target.checked)}
                disabled={disabled} />
              {label as string}
            </label>
          ))}
        </div>
        {slideChart && (
          <div className="chart-type-row">
            <span>商品グラフの種類：</span>
            {(['bar', 'pie'] as const).map(v => (
              <label key={v} className="radio-label">
                <input type="radio" name="chartType" value={v}
                  checked={chartProductType === v}
                  onChange={() => setChartProductType(v)}
                  disabled={disabled} />
                {v === 'bar' ? '棒グラフ' : '円グラフ（ドーナツ）'}
              </label>
            ))}
          </div>
        )}
      </div>

      {/* ── 追加プロンプト ── */}
      <div className="form-group">
        <div className="extra-label-row">
          <label className="label-with-badge">
            追加プロンプト
            <span className="badge-optional">任意</span>
          </label>
          {savedTemplates.length > 0 && (
            <select
              className="template-load-select"
              defaultValue=""
              onChange={e => { if (e.target.value) loadTemplate(e.target.value); e.target.value = '' }}
              disabled={disabled}
            >
              <option value="" disabled>テンプレートを読込...</option>
              {savedTemplates.map(t => (
                <option key={t.id} value={t.id}>{t.name}</option>
              ))}
            </select>
          )}
        </div>
        <textarea
          className="extra-context-input"
          rows={3}
          placeholder="例: 重点的に分析してほしい点や特記事項を入力してください。&#10;例: 「地域別の課題を重点的に記載すること」「コスト削減施策を提案すること」"
          value={extraContext}
          onChange={e => setExtraContext(e.target.value)}
          disabled={disabled}
        />
        <div className="extra-context-footer">
          <p className="field-hint">ここに入力した内容はAIへの追加指示として反映されます。</p>
          <button
            type="button"
            className="btn-save-template"
            onClick={saveTemplate}
            disabled={disabled || !extraContext.trim()}
          >
            ＋ テンプレートに保存
          </button>
        </div>
        {savedTemplates.length > 0 && (
          <div className="template-list">
            {savedTemplates.map(t => (
              <div key={t.id} className="template-item">
                <button
                  type="button"
                  className="template-item-load"
                  onClick={() => loadTemplate(t.id)}
                  disabled={disabled}
                  title={t.text}
                >
                  {t.name}
                </button>
                <button
                  type="button"
                  className="template-item-del"
                  onClick={() => deleteTemplate(t.id)}
                  disabled={disabled}
                  title="削除"
                >
                  ✕
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      <button type="submit" className="btn-primary" disabled={disabled || !canSubmit}>
        レポートを生成する
      </button>
    </form>
  )
}

export default UploadForm
