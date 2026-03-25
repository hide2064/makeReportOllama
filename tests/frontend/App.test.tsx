/**
 * App.test.tsx — App コンポーネントの UI テスト (Vitest + Testing Library)
 * fetch は vi.fn() でモック化する。
 */
import { render, screen, fireEvent, waitFor } from '@testing-library/react'
import { describe, it, expect, vi, beforeEach } from 'vitest'
import App from '../../frontend/src/App'

// fetch をグローバルにモック
const mockFetch = vi.fn()
global.fetch = mockFetch

// URL.createObjectURL をモック
global.URL.createObjectURL = vi.fn(() => 'blob:mock-url')

beforeEach(() => {
  mockFetch.mockReset()
})

describe('App', () => {
  it('初期状態でタイトルが表示される', () => {
    render(<App />)
    expect(screen.getByText('売上報告書 自動生成')).toBeInTheDocument()
  })

  it('ファイル未選択時は生成ボタンが無効', () => {
    render(<App />)
    const btn = screen.getByRole('button', { name: /レポートを生成する/ })
    expect(btn).toBeDisabled()
  })

  it('API 成功時にダウンロードボタンが表示される', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      blob: async () => new Blob(['dummy'], { type: 'application/octet-stream' }),
    })

    render(<App />)

    // Excel ファイルを選択
    const excelInput = document.querySelectorAll('input[type="file"]')[0] as HTMLInputElement
    fireEvent.change(excelInput, {
      target: { files: [new File(['dummy'], 'sales.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })] },
    })

    // PPTX ファイルを選択
    const pptxInput = document.querySelectorAll('input[type="file"]')[1] as HTMLInputElement
    fireEvent.change(pptxInput, {
      target: { files: [new File(['dummy'], 'template.pptx', { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })] },
    })

    // 生成ボタンをクリック
    const btn = screen.getByRole('button', { name: /レポートを生成する/ })
    fireEvent.click(btn)

    await waitFor(() => {
      expect(screen.getByText(/report.pptx をダウンロード/)).toBeInTheDocument()
    })
  })

  it('API エラー時にエラーメッセージが表示される', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: false,
      json: async () => ({ detail: 'Ollama に接続できません' }),
    })

    render(<App />)

    const excelInput = document.querySelectorAll('input[type="file"]')[0] as HTMLInputElement
    fireEvent.change(excelInput, {
      target: { files: [new File(['d'], 'sales.xlsx')] },
    })
    const pptxInput = document.querySelectorAll('input[type="file"]')[1] as HTMLInputElement
    fireEvent.change(pptxInput, {
      target: { files: [new File(['d'], 'template.pptx')] },
    })

    fireEvent.click(screen.getByRole('button', { name: /レポートを生成する/ }))

    await waitFor(() => {
      expect(screen.getByText(/エラーが発生しました/)).toBeInTheDocument()
      expect(screen.getByText(/Ollama に接続できません/)).toBeInTheDocument()
    })
  })
})
