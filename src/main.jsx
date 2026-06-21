import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

function showError(message, stack) {
  const root = document.getElementById('root')
  root.innerHTML = `<div style="background:#1a0505;color:#fecaca;padding:20px;font-family:monospace;min-height:100vh;direction:ltr;text-align:left;"><h2 style="color:#f87171;margin-top:0;">Error</h2><p style="white-space:pre-wrap;word-break:break-word;font-size:14px;background:#2a0a0a;padding:12px;border-radius:8px;">${message}</p>${stack ? `<pre style="white-space:pre-wrap;word-break:break-word;font-size:11px;color:#fda4af;margin-top:12px;">${stack}</pre>` : ''}</div>`
}

window.addEventListener('error', (e) => { showError(e.message, e.error && e.error.stack) })
window.addEventListener('unhandledrejection', (e) => { showError('Promise Rejection: ' + (e.reason && e.reason.message ? e.reason.message : String(e.reason)), e.reason && e.reason.stack) })

class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null } }
  static getDerivedStateFromError(error) { return { error } }
  render() {
    if (this.state.error) {
      return (
        <div style={{background:'#1a0505',color:'#fecaca',padding:20,fontFamily:'monospace',minHeight:'100vh',direction:'ltr',textAlign:'left'}}>
          <h2 style={{color:'#f87171',marginTop:0}}>React Error</h2>
          <pre style={{whiteSpace:'pre-wrap',wordBreak:'break-word',fontSize:14,background:'#2a0a0a',padding:12,borderRadius:8}}>{this.state.error.message}</pre>
          <pre style={{whiteSpace:'pre-wrap',wordBreak:'break-word',fontSize:11,color:'#fda4af',marginTop:12}}>{this.state.error.stack}</pre>
        </div>
      )
    }
    return this.props.children
  }
}

try {
  ReactDOM.createRoot(document.getElementById('root')).render(
    <ErrorBoundary><App /></ErrorBoundary>
  )
} catch (err) { showError(err.message, err.stack) }
