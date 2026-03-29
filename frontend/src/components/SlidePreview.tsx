/**
 * SlidePreview.tsx
 * 生成済み PPTX のスライド内容をカード形式でブラウザ表示するコンポーネント。
 */
import React from 'react'
import './SlidePreview.css'

export interface SlideInfo {
  slide_num: number
  title:     string
  texts:     string[]
  type:      'text' | 'table' | 'chart'
}

interface Props {
  slides:  SlideInfo[]
  onClose: () => void
}

const TYPE_LABEL: Record<SlideInfo['type'], string> = {
  text:  '',
  table: '📊 表データ',
  chart: '📈 グラフ',
}

const SlidePreview: React.FC<Props> = ({ slides, onClose }) => {
  return (
    <div className="slide-preview">
      <div className="slide-preview-header">
        <h3 className="slide-preview-title">スライドプレビュー（{slides.length} スライド）</h3>
        <button type="button" className="slide-preview-close" onClick={onClose}>✕ 閉じる</button>
      </div>

      <div className="slide-preview-list">
        {slides.map(slide => (
          <div
            key={slide.slide_num}
            className={`slide-card slide-card-${slide.type}`}
          >
            <div className="slide-card-num">{slide.slide_num}</div>
            <div className="slide-card-body">
              {slide.title && (
                <div className="slide-card-title">{slide.title}</div>
              )}
              {slide.type !== 'text' && TYPE_LABEL[slide.type] && (
                <div className="slide-card-indicator">{TYPE_LABEL[slide.type]}</div>
              )}
              {slide.type === 'text' && slide.texts.length > 0 && (
                <div className="slide-card-texts">
                  {slide.texts.map((t, i) => (
                    <p key={i} className="slide-card-text">{t}</p>
                  ))}
                </div>
              )}
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}

export default SlidePreview
