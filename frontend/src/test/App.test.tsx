/**
 * App.test.tsx — App コンポーネントの UI テスト (Vitest + Testing Library)
 * fetch は vi.fn() でモック化する。ポーリング方式に対応。
 * ReferenceManager が mount 時に /api/references を呼ぶため、
 * URL でルーティングする mockFetch.mockImplementation を使用する。
 */
import { render, screen, fireEvent, act } from '@testing-library/react'
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import App from '../App'

const mockFetch = vi.fn()
global.fetch = mockFetch
global.URL.createObjectURL = vi.fn(() => 'blob:mock-url')

/** /api/references は常に空リストを返し、それ以外は queue から順番に返す */
function setupFetch(...responses: object[]) {
  const queue = [...responses]
  mockFetch.mockImplementation(async (url: unknown) => {
    if (String(url).startsWith('/api/references')) {
      return { ok: true, json: async () => ({ references: [] }) }
    }
    const next = queue.shift()
    if (!next) throw new Error(`Unexpected fetch: ${url}`)
    return next
  })
}

beforeEach(() => {
  mockFetch.mockReset()
  vi.useFakeTimers()
})

afterEach(() => {
  vi.useRealTimers()
})

function selectFiles() {
  fireEvent.change(
    document.querySelectorAll('input[type="file"]')[0] as HTMLInputElement,
    { target: { files: [new File(['d'], 'sales.xlsx')] } }
  )
  fireEvent.change(
    document.querySelectorAll('input[type="file"]')[1] as HTMLInputElement,
    { target: { files: [new File(['d'], 'template.pptx')] } }
  )
}

describe('App', () => {
  it('初期状態でタイトルが表示される', () => {
    setupFetch()
    render(<App />)
    expect(screen.getByText('売上報告書 自動生成')).toBeInTheDocument()
  })

  it('ファイル未選択時は生成ボタンが無効', () => {
    setupFetch()
    render(<App />)
    expect(screen.getByRole('button', { name: /レポートを生成する/ })).toBeDisabled()
  })

  it('ポーリング成功時にダウンロードボタンが表示される', async () => {
    setupFetch(
      { ok: true, json: async () => ({ status: 'started' }) },                              // POST /api/generate
      { ok: true, json: async () => ({ step: '完了しました！', done: true, error: '' }) }, // GET /api/progress
      { ok: true, blob: async () => new Blob(['dummy']) },                                  // GET /api/download
    )

    render(<App />)
    selectFiles()

    await act(async () => {
      fireEvent.click(screen.getByRole('button', { name: /レポートを生成する/ }))
    })

    await act(async () => {
      await vi.runAllTimersAsync()
    })

    expect(screen.getByText(/report.pptx をダウンロード/)).toBeInTheDocument()
  })

  it('progress に error が入ったときエラー表示される', async () => {
    setupFetch(
      { ok: true, json: async () => ({ status: 'started' }) },
      { ok: true, json: async () => ({ step: '', done: false, error: 'Ollama に接続できません' }) },
    )

    render(<App />)
    selectFiles()

    await act(async () => {
      fireEvent.click(screen.getByRole('button', { name: /レポートを生成する/ }))
    })

    await act(async () => {
      await vi.runAllTimersAsync()
    })

    expect(screen.getByText(/エラーが発生しました/)).toBeInTheDocument()
    expect(screen.getByText(/Ollama に接続できません/)).toBeInTheDocument()
  })

  it('POST 失敗時にエラーメッセージが表示される', async () => {
    setupFetch(
      { ok: false, json: async () => ({ detail: 'Internal Server Error' }) },
    )

    render(<App />)
    selectFiles()

    await act(async () => {
      fireEvent.click(screen.getByRole('button', { name: /レポートを生成する/ }))
    })

    await act(async () => {
      await vi.runAllTimersAsync()
    })

    expect(screen.getByText(/エラーが発生しました/)).toBeInTheDocument()
  })
})
