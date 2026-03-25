import React, { useRef, useState } from 'react'
import './UploadForm.css'

interface Props {
  onGenerate: (excel: File, template: File) => void
  disabled: boolean
}

const UploadForm: React.FC<Props> = ({ onGenerate, disabled }) => {
  const [excelFile, setExcelFile]       = useState<File | null>(null)
  const [templateFile, setTemplateFile] = useState<File | null>(null)
  const excelRef    = useRef<HTMLInputElement>(null)
  const templateRef = useRef<HTMLInputElement>(null)

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    if (!excelFile || !templateFile) return
    onGenerate(excelFile, templateFile)
  }

  return (
    <form className="upload-form" onSubmit={handleSubmit}>
      <div className="form-group">
        <label>売上データ Excel (.xlsx)</label>
        <div className="file-row">
          <input
            ref={excelRef}
            type="file"
            accept=".xlsx"
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
          <span className="file-name">
            {excelFile ? excelFile.name : '未選択'}
          </span>
        </div>
      </div>

      <div className="form-group">
        <label>PPTX テンプレート (.pptx)</label>
        <div className="file-row">
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
          <span className="file-name">
            {templateFile ? templateFile.name : '未選択'}
          </span>
        </div>
      </div>

      <button
        type="submit"
        className="btn-primary"
        disabled={disabled || !excelFile || !templateFile}
      >
        レポートを生成する
      </button>
    </form>
  )
}

export default UploadForm
