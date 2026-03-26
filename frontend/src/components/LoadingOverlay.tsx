import React from 'react'
import './LoadingOverlay.css'

interface Props {
  step?: string
}

const LoadingOverlay: React.FC<Props> = ({ step = '' }) => {
  return (
    <div className="overlay">
      <div className="overlay-box">
        <div className="spinner" />
        <p className="overlay-message">レポートを生成中です...</p>
        {step && (
          <p className="overlay-step">{step}</p>
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
