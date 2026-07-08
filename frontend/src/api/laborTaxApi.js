const API_BASE = import.meta.env.VITE_API_BASE || ''

async function handleJsonResponse(response) {
  const payload = await response.json().catch(() => ({}))
  if (!response.ok) {
    throw new Error(payload.detail || '请求失败')
  }
  return payload
}

export async function downloadTemplate() {
  const response = await fetch(`${API_BASE}/api/template`)
  if (!response.ok) throw new Error('模板下载失败')
  const blob = await response.blob()
  downloadBlob(blob, '劳务费基础数据导入模板.xlsx')
}

export async function calculateUpload(file) {
  const formData = new FormData()
  formData.append('file', file)
  const response = await fetch(`${API_BASE}/api/calculate/upload`, {
    method: 'POST',
    body: formData
  })
  return handleJsonResponse(response)
}

export async function calculateManual(rows) {
  const response = await fetch(`${API_BASE}/api/calculate/manual`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ rows })
  })
  return handleJsonResponse(response)
}

export async function exportLedger(rows) {
  const response = await fetch(`${API_BASE}/api/export`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ rows })
  })
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}))
    throw new Error(payload.detail || '导出失败')
  }
  const blob = await response.blob()
  downloadBlob(blob, '劳务费税费换算台账.xlsx')
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
}
