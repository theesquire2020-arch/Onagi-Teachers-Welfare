// Same-origin bridge for the Vercel UI. Apps Script's ContentService redirect is
// followed server-side, so browsers never need cross-origin access to Sheets.
const DEFAULT_WEB_APP = 'https://script.google.com/macros/s/AKfycbyd4zRrxgvSHNmTBF3FMhMDPRE5RsUt6VVpZ3gzQBv-VOwVSj7iuggqsqFjkbK99jfHIw/exec'
const ACTIONS = new Set(['register','login','logout','getDashboard','getMembers','getContributions','getLoans','getBonuses','getBeneficiaries','getGallery','getDocuments','getSettings','getPerformance','getExpenditures','addExpenditure','testConnection','registerDatabase','requestLoan','addBeneficiary','initiatePayment','addMemberContribution','approveLoan','updateMember','awardBonus','saveSettings','switchDatabase','exportData','rebuildMatrices','addGalleryItem','addDocument','getAuditLog'])

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, private')
  res.setHeader('Content-Type', 'application/json; charset=utf-8')
  if (req.method !== 'POST') return res.status(405).json({success:false,message:'POST required.',data:null})
  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body || {}
    if (!ACTIONS.has(body.action)) return res.status(400).json({success:false,message:'Unknown action.',data:null})
    const payload = JSON.stringify({action:body.action,payload:body.payload || {},token:body.token || ''})
    if (payload.length > 65536) return res.status(413).json({success:false,message:'Request too large.',data:null})
    const endpoint = process.env.GAS_WEBAPP_URL || DEFAULT_WEB_APP
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(endpoint)) throw Error('Apps Script deployment URL is not configured correctly.')
    const upstream = await fetch(endpoint, {method:'POST',headers:{'Content-Type':'application/json'},body:payload,redirect:'follow',signal:AbortSignal.timeout(25000)})
    const text = await upstream.text()
    let result
    try { result = JSON.parse(text) } catch { throw Error('The Apps Script deployment did not return API data. Redeploy the latest Code.gs and allow web-app access.') }
    if (!result || typeof result.success !== 'boolean') throw Error('Invalid response from the Apps Script deployment.')
    return res.status(upstream.ok ? 200 : 502).json(result)
  } catch (error) {
    return res.status(502).json({success:false,message:error.name === 'TimeoutError' ? 'Live database timed out. Try again.' : String(error.message || 'Live database unavailable.'),data:null})
  }
}
