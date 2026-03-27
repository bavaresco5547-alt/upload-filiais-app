import { supabase } from './supabase'

function sanitizeFileName(name) {
  return String(name)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '_')
    .replace(/[^a-zA-Z0-9._/-]/g, '_')
    .replace(/_+/g, '_')
}

export async function uploadFileToStorage(file, filial) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-')
  const safeName = sanitizeFileName(file.name)
  const safeFilial = sanitizeFileName(filial)
  const filePath = `${safeFilial}/${timestamp}-${safeName}`

  const { data, error } = await supabase.storage
    .from('uploads-planilhas')
    .upload(filePath, file, {
      upsert: false,
      contentType: file.type || 'application/octet-stream'
    })

  if (error) {
    throw new Error(`Erro no upload para o Storage: ${error.message}`)
  }

  return data
}