import React from 'react'
import './LoadingOverlay.css'

interface Props {
  message?: string
}

const LoadingOverlay: React.FC<Props> = ({ message = 'レポートを生成中です...' }) => {
  return (
    <div className="overlay">
      <div className="overlay-box">
        <div className="spinner" />
        <p className="overlay-message">{message}</p>
        <p className="overlay-sub">
          ローカル LLM (Ollama) で分析中です。
          <br />
          CPU 推論のため数分かかる場合があります。
        </p>
      </div>
    </div>
  )
}

export default LoadingOverlay
