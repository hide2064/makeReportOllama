import React, { useEffect, useRef, useState } from 'react'
import './UploadForm.css'

export interface GenerateParams {
  excelFile: File
  templateFile: File | null
  templateName: string
  slideProductTable: boolean
  slideRegionTable: boolean
  slideRepTable: boolean
  slideChart: boolean
  chartProductType: 'bar' | 'pie'
}

interface Props {
  onGenerate: (params: GenerateParams) => void
  disabled: boolean
}

const UploadForm: React.FC<Props> = ({ onGenerate, disabled }) => {
  const [excelFile, setExcelFile]           = useState<File | null>(null)
  const [templateFile, setTemplateFile]     = useState<File | null>(null)
  const [templateMode, setTemplateMode]     = useState<'upload' | 'server'>('upload')
  const [serverTemplates, setServerTemplates] = useState<{ name: string }[]>([])
  const [selectedTemplate, setSelectedTemplate] = useState<string>('')
  const [slideProductTable, setSlideProductTable] = useState(true)
  const [slideRegionTable, setSlideRegionTable]   = useState(false)
  const [slideRepTable, setSlideRepTable]         = useState(false)
  const [slideChart, setSlideChart]               = useState(true)
  const [chartProductType, setChartProductType]   = useState<'bar' | 'pie'>('bar')

  const excelRef    = useRef<HTMLInputElement>(null)
  const templateRef = useRef<HTMLInputElement>(null)

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

  const canSubmit = !!excelFile && (
    templateMode === 'upload' ? !!templateFile : !!selectedTemplate
  )

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    if (!excelFile) return
    onGenerate({
      excelFile,
      templateFile: templateMode === 'upload' ? templateFile : null,
      templateName: templateMode === 'server' ? selectedTemplate : '',
      slideProductTable,
      slideRegionTable,
      slideRepTable,
      slideChart,
      chartProductType,
    })
  }

  return (
    <form className="upload-form" onSubmit={handleSubmit}>
      {/* 売上データ */}
      <div className="form-group">
        <label>売上データ (.xlsx / .csv)</label>
        <div className="file-row">
          <input
            ref={excelRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            style={{ display: 'none' }}
            onChange={e => setExcelFile(e.target.files?.[0] ?? null)}
          />
          <button
            type="button"
            className="btn-secondary"
            onClick={() => excelRef.current?.click()}
            disabled={disabled}
          >
            ファイルを選択
          </button>
          <span className="file-name">{excelFile ? excelFile.name : '未選択'}</span>
        </div>
      </div>

      {/* PPTX テンプレート */}
      <div className="form-group">
        <label>PPTX テンプレート (.pptx)</label>
        <div className="template-tabs">
          <button
            type="button"
            className={`tab-btn${templateMode === 'upload' ? ' active' : ''}`}
            onClick={() => setTemplateMode('upload')}
            disabled={disabled}
          >
            ファイルアップロード
          </button>
          <button
            type="button"
            className={`tab-btn${templateMode === 'server' ? ' active' : ''}`}
            onClick={() => setTemplateMode('server')}
            disabled={disabled}
          >
            サーバーのテンプレート
          </button>
        </div>

        {/* ファイルアップロードタブ（常にレンダリング、visibility切替） */}
        <div style={{ display: templateMode === 'upload' ? 'flex' : 'none' }} className="file-row">
          <input
            ref={templateRef}
            type="file"
            accept=".pptx"
            style={{ display: 'none' }}
            onChange={e => setTemplateFile(e.target.files?.[0] ?? null)}
          />
          <button
            type="button"
            className="btn-secondary"
            onClick={() => templateRef.current?.click()}
            disabled={disabled}
          >
            ファイルを選択
          </button>
          <span className="file-name">{templateFile ? templateFile.name : '未選択'}</span>
        </div>

        {/* サーバーテンプレートタブ */}
        {templateMode === 'server' && (
          <select
            className="template-select"
            value={selectedTemplate}
            onChange={e => setSelectedTemplate(e.target.value)}
            disabled={disabled}
          >
            {serverTemplates.length === 0
              ? <option value="">テンプレートなし</option>
              : serverTemplates.map(t => (
                  <option key={t.name} value={t.name}>{t.name}</option>
                ))
            }
          </select>
        )}
      </div>

      {/* スライド構成オプション */}
      <div className="form-group slide-options">
        <label>スライド構成</label>
        <div className="options-row">
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={slideProductTable}
              onChange={e => setSlideProductTable(e.target.checked)}
              disabled={disabled}
            />
            商品別売上表
          </label>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={slideRegionTable}
              onChange={e => setSlideRegionTable(e.target.checked)}
              disabled={disabled}
            />
            地域別売上表
          </label>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={slideRepTable}
              onChange={e => setSlideRepTable(e.target.checked)}
              disabled={disabled}
            />
            担当者別売上表
          </label>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={slideChart}
              onChange={e => setSlideChart(e.target.checked)}
              disabled={disabled}
            />
            グラフスライド
          </label>
        </div>
        {slideChart && (
          <div className="chart-type-row">
            <span>商品グラフの種類：</span>
            <label className="radio-label">
              <input
                type="radio"
                name="chartType"
                value="bar"
                checked={chartProductType === 'bar'}
                onChange={() => setChartProductType('bar')}
                disabled={disabled}
              />
              棒グラフ
            </label>
            <label className="radio-label">
              <input
                type="radio"
                name="chartType"
                value="pie"
                checked={chartProductType === 'pie'}
                onChange={() => setChartProductType('pie')}
                disabled={disabled}
              />
              円グラフ（ドーナツ）
            </label>
          </div>
        )}
      </div>

      <button
        type="submit"
        className="btn-primary"
        disabled={disabled || !canSubmit}
      >
        レポートを生成する
      </button>
    </form>
  )
}

export default UploadForm
