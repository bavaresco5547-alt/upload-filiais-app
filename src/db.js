import { supabase } from './supabase'

function deduplicateByKey(rows) {
  const map = new Map()

  for (const row of rows) {
    const key = row.chave_unica
    if (!key) continue
    map.set(key, row)
  }

  return Array.from(map.values())
}

export async function createUploadRecord({ filial, nomeArquivo, caminhoStorage }) {
  const { data, error } = await supabase
    .from('uploads')
    .insert([
      {
        filial,
        nome_arquivo: nomeArquivo,
        caminho_storage: caminhoStorage,
        status: 'PROCESSANDO'
      }
    ])
    .select()
    .single()

  if (error) {
    throw new Error(`Erro ao gravar upload: ${error.message}`)
  }

  return data
}

export async function updateUploadStatus(id, data) {
  if (!id) return

  const { error } = await supabase
    .from('uploads')
    .update(data)
    .eq('id', id)

  if (error) {
    console.error('Erro ao atualizar status do upload:', error)
  }
}

export async function insertFretes(rows, uploadId) {
  if (!rows.length) return 0

  const payload = deduplicateByKey(
    rows.map((row) => ({
      ...row,
      upload_id: uploadId
    }))
  )

  const { error } = await supabase
    .from('fretes_unificado')
    .upsert(payload, { onConflict: 'chave_unica' })

  if (error) {
    throw new Error(`Erro ao gravar FRETES: ${error.message}`)
  }

  return payload.length
}

export async function insertCustoFrota(rows, uploadId) {
  if (!rows.length) return 0

  const payload = deduplicateByKey(
    rows.map((row) => ({
      ...row,
      upload_id: uploadId
    }))
  )

  const { error } = await supabase
    .from('custo_frota_unificado')
    .upsert(payload, { onConflict: 'chave_unica' })

  if (error) {
    throw new Error(`Erro ao gravar CUSTO FROTA: ${error.message}`)
  }

  return payload.length
}

export async function insertOciosidade(rows, uploadId) {
  if (!rows.length) return 0

  const payload = deduplicateByKey(
    rows.map((row) => ({
      ...row,
      upload_id: uploadId
    }))
  )

  const { error } = await supabase
    .from('ociosidade_unificada')
    .upsert(payload, { onConflict: 'chave_unica' })

  if (error) {
    throw new Error(`Erro ao gravar OCIOSIDADE: ${error.message}`)
  }

  return payload.length
}