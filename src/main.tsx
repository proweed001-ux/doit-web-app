import React from 'react'
import ReactDOM from 'react-dom/client'

function App() {
  return (
    <div style={{padding:'40px',fontFamily:'sans-serif'}}>
      <h1>DOIT Web App</h1>
      <p>Deployment connected successfully.</p>
    </div>
  )
}

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
