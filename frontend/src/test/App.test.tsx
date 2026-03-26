/**
 * App.test.tsx — App コンポーネントの UI テスト (Vitest + Testing Library)
 * fetch は vi.fn() でモック化する。SSE レスポンスに対応。
 */
import { render, screen, fireEvent, waitFor } from '@testing-library/react'
import { describe, it, expect, vi, beforeEach } from 'vitest'
import App from '../App'

const mockFetch = vi.fn()
global.fetch = mockFetch
global.URL.createObjectURL = vi.fn(() => 'blob:mock-url')

/** SSE テキストから ReadableStream を生成するヘルパー */
function makeSSEStream(events: object[]): ReadableStream {
  const text = events.map(e => `data: ${JSON.stringify(e)}\n\n`).join('')
  return new ReadableStream({
    start(controller) {
      controller.enqueue(new TextEncoder().encode(text))
      controller.close()
    },
  })
}

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
    expect(screen.getByRole('button', { name: /レポートを生成する/ })).toBeDisabled()
  })

  it('SSE 成功時にダウンロードボタンが表示される', async () => {
    // 1回目: SSE ストリーム (generate)
    mockFetch.mockResolvedValueOnce({
      ok: true,
      body: makeSSEStream([
        { step: '[1/3]  Excel を読み込んでいます...' },
        { step: '[2/3]  Ollama で分析中...' },
        { step: '[3/3]  PPTX を生成中...' },
        { done: true },
      ]),
    })
    // 2回目: ダウンロード (GET /api/download)
    mockFetch.mockResolvedValueOnce({
      ok: true,
      blob: async () => new Blob(['dummy'], { type: 'application/octet-stream' }),
    })

    render(<App />)

    fireEvent.change(document.querySelectorAll('input[type="file"]')[0] as HTMLInputElement, {
      target: { files: [new File(['d'], 'sales.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })] },
    })
    fireEvent.change(document.querySelectorAll('input[type="file"]')[1] as HTMLInputElement, {
      target: { files: [new File(['d'], 'template.pptx', { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })] },
    })
    fireEvent.click(screen.getByRole('button', { name: /レポートを生成する/ }))

    await waitFor(() => {
      expect(screen.getByText(/report.pptx をダウンロード/)).toBeInTheDocument()
    })
  })

  it('SSE error イベント時にエラーメッセージが表示される', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      body: makeSSEStream([
        { step: '[1/3]  Excel を読み込んでいます...' },
        { error: 'Ollama に接続できません' },
      ]),
    })

    render(<App />)

    fireEvent.change(document.querySelectorAll('input[type="file"]')[0] as HTMLInputElement, {
      target: { files: [new File(['d'], 'sales.xlsx')] },
    })
    fireEvent.change(document.querySelectorAll('input[type="file"]')[1] as HTMLInputElement, {
      target: { files: [new File(['d'], 'template.pptx')] },
    })
    fireEvent.click(screen.getByRole('button', { name: /レポートを生成する/ }))

    await waitFor(() => {
      expect(screen.getByText(/エラーが発生しました/)).toBeInTheDocument()
      expect(screen.getByText(/Ollama に接続できません/)).toBeInTheDocument()
    })
  })

  it('HTTP エラー時にエラーメッセージが表示される', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: false,
      json: async () => ({ detail: 'Internal Server Error' }),
    })

    render(<App />)

    fireEvent.change(document.querySelectorAll('input[type="file"]')[0] as HTMLInputElement, {
      target: { files: [new File(['d'], 'sales.xlsx')] },
    })
    fireEvent.change(document.querySelectorAll('input[type="file"]')[1] as HTMLInputElement, {
      target: { files: [new File(['d'], 'template.pptx')] },
    })
    fireEvent.click(screen.getByRole('button', { name: /レポートを生成する/ }))

    await waitFor(() => {
      expect(screen.getByText(/エラーが発生しました/)).toBeInTheDocument()
    })
  })
})
