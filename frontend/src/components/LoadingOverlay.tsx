import React from 'react'
import './LoadingOverlay.css'

interface Props {
  step?:          string
  progress?:      number
  queuePosition?: number
}

const LoadingOverlay: React.FC<Props> = ({ step = '', progress = 0, queuePosition = 0 }) => {
  return (
    <div className="overlay">
      <div className="overlay-box">
        <div className="spinner" />
        <p className="overlay-message">
          {queuePosition > 0 ? 'キューで待機中...' : 'レポートを生成中です...'}
        </p>

        {queuePosition > 0 ? (
          <div className="overlay-queue-badge">
            {queuePosition} 番目
          </div>
        ) : (
          <>
            {step && <p className="overlay-step">{step}</p>}
            <div className="progress-wrap">
              <div className="progress-bar">
                <div
                  className="progress-fill"
                  style={{ width: `${Math.max(progress, 3)}%` }}
                />
              </div>
              <span className="progress-pct">{progress}%</span>
            </div>
          </>
        )}

        <p className="overlay-sub">
          処理が完了するまでこのままお待ちください。
          <br />
          CPU 推論のため最大 20 分かかる場合があります。
        </p>
      </div>
    </div>
  )
}

export default LoadingOverlay
