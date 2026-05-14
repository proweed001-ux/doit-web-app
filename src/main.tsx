import React, { useState } from 'react'
import ReactDOM from 'react-dom/client'

function App() {
  const [fileName, setFileName] = useState('')
  const [rows, setRows] = useState(0)
  const [status, setStatus] = useState('พร้อมใช้งาน')

  async function handleFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return

    setFileName(file.name)
    setStatus('กำลังอ่านไฟล์...')

    try {
      const text = await file.text()
      const lineCount = text.split(/\r?\n/).length
      setRows(lineCount)
      setStatus('โหลดไฟล์สำเร็จ')
    } catch (err) {
      setStatus('ไม่สามารถอ่านไฟล์ได้')
    }
  }

  return (
    <div style={{
      minHeight: '100vh',
      background: '#0f172a',
      color: 'white',
      fontFamily: 'sans-serif',
      padding: '40px'
    }}>
      <h1 style={{fontSize:'32px',marginBottom:'8px'}}>AYA DOIT Web App</h1>
      <p style={{opacity:0.7}}>ระบบอัปโหลดและตรวจข้อมูลเบื้องต้น</p>

      <div style={{
        marginTop:'24px',
        padding:'24px',
        border:'1px solid #334155',
        borderRadius:'16px',
        background:'#111827'
      }}>
        <input type="file" onChange={handleFile} />

        <div style={{marginTop:'20px'}}>
          <div><strong>สถานะ:</strong> {status}</div>
          <div><strong>ไฟล์:</strong> {fileName || '-'} </div>
          <div><strong>จำนวนบรรทัด:</strong> {rows}</div>
        </div>
      </div>
    </div>
  )
}

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
