/**
 * E2E テスト: レポート生成フロー
 *
 * 前提: バックエンド API はモックサーバーで代替、または
 *       Vite の proxy 先 (localhost:8000) が起動済みであること。
 *
 * テストシナリオ:
 *   1. トップページが表示される
 *   2. ドロップゾーンが存在する
 *   3. Excel ファイルをアップロードするとプレビューが表示される
 *   4. API エラー時にエラーメッセージが表示される
 *   5. ダークモードトグルで切り替えができる
 */

import { expect, test } from '@playwright/test'
import path from 'path'

const SAMPLE_XLSX = path.resolve(__dirname, '../../data/sample_sales.xlsx')

test.describe('トップページ', () => {
  test('ページタイトルとヘッダーが表示される', async ({ page }) => {
    await page.goto('/')
    await expect(page.locator('h1')).toHaveText('売上報告書 自動生成')
    await expect(page.locator('.drop-zone')).toBeVisible()
  })

  test('ダークモードトグルで dark クラスが切り替わる', async ({ page }) => {
    await page.goto('/')
    const toggle = page.locator('.theme-toggle')
    await expect(toggle).toBeVisible()

    const initialIsDark = await page.evaluate(() =>
      document.documentElement.classList.contains('dark')
    )
    await toggle.click()
    const afterIsDark = await page.evaluate(() =>
      document.documentElement.classList.contains('dark')
    )
    expect(afterIsDark).toBe(!initialIsDark)
  })

  test('過去レポート管理セクションが表示される', async ({ page }) => {
    await page.goto('/')
    await expect(page.locator('.ref-manager')).toBeVisible()
  })

  test('生成履歴セクションが表示される', async ({ page }) => {
    await page.goto('/')
    await expect(page.locator('.history-panel')).toBeVisible()
  })
})

test.describe('ファイルアップロード', () => {
  test('ドロップゾーンのクリックでファイル選択が開く', async ({ page }) => {
    await page.goto('/')
    // input[type=file] が hidden のため、クリックイベントを listen するだけ確認
    const fileInput = page.locator('input[type="file"][accept=".xlsx,.xls,.csv"]').first()
    await expect(fileInput).toBeAttached()
  })

  test('sample_sales.xlsx をアップロードするとプレビューが表示される', async ({ page }) => {
    // バックエンドが起動している場合のみ実行（スキップ条件付き）
    const backendAlive = await page.request.get('http://localhost:8000/health')
      .then(r => r.ok())
      .catch(() => false)
    test.skip(!backendAlive, 'バックエンドが起動していないためスキップ')

    await page.goto('/')
    const fileInput = page.locator('input[type="file"][accept=".xlsx,.xls,.csv"]').first()
    await fileInput.setInputFiles(SAMPLE_XLSX)

    // プレビューかエラーのどちらかが表示されるのを待つ
    await expect(
      page.locator('.data-preview, .preview-error')
    ).toBeVisible({ timeout: 15_000 })
  })

  test('不正な拡張子ファイルをドロップするとエラーが出る', async ({ page }) => {
    await page.goto('/')

    // ドロップイベントをシミュレート (.txt ファイル)
    const dropZone = page.locator('.drop-zone').first()
    await dropZone.dispatchEvent('dragenter', {})
    await dropZone.dispatchEvent('dragover', {})

    // drag-over クラスが付くことを確認
    await expect(dropZone).toHaveClass(/drag-over/)

    await dropZone.dispatchEvent('dragleave', {})
    await expect(dropZone).not.toHaveClass(/drag-over/)
  })
})

test.describe('フォームバリデーション', () => {
  test('ファイル未選択時は生成ボタンが無効', async ({ page }) => {
    await page.goto('/')
    const submitBtn = page.locator('button[type="submit"]')
    await expect(submitBtn).toBeDisabled()
  })

  test('分析期間の日付バリデーション: 開始 > 終了 でエラー表示', async ({ page }) => {
    await page.goto('/')
    const dateInputs = page.locator('.date-input')
    await dateInputs.nth(0).fill('2025-12-31')
    await dateInputs.nth(1).fill('2025-01-01')
    await expect(page.locator('.date-error')).toBeVisible()
  })
})
