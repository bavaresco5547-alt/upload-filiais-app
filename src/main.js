import './style.css'
import { readExcelFile } from './excel-reader'
import { uploadFileToStorage } from './storage'
import {
  createUploadRecord,
  insertFretes,
  insertCustoFrota,
  insertOciosidade,
  updateUploadStatus
} from './db'

document.getElementById('app').innerHTML = `
  <div class="container">
    <h1>Upload Unificado das Filiais</h1>
    <p class="subtitle">Envie a planilha para extrair os dados e alimentar o Supabase.</p>

    <div class="card">
      <label for="filial">Filial</label>
      <select id="filial">
        <option value="">Selecione a filial</option>
        <option value="SAO_JOSE">SÃO JOSÉ</option>
        <option value="SAO_LOURENCO">SÃO LOURENÇO</option>
        <option value="MARINGA">MARINGA</option>
        <option value="BRASILIA">BRASILIA</option>
        <option value="HORIZONTE">HORIZONTE</option>
        <option value="CURITIBA">CURITIBA</option>
        <option value="CAMPINAS">CAMPINAS</option>
        <option value="CAMPINAS">CAJURU</option>
      </select>

      <label for="fileInput">Planilha Excel</label>
      <input type="file" id="fileInput" accept=".xlsx,.xls,.xlsm,.xlsb" />

      <button id="btnUpload">Processar planilha</button>

      <div id="status" class="status"></div>
      <div id="summary" class="summary"></div>
      <pre id="preview" class="preview"></pre>
    </div>
  </div>
`

const btnUpload = document.getElementById('btnUpload')
const fileInput = document.getElementById('fileInput')
const filialSelect = document.getElementById('filial')
const status = document.getElementById('status')
const summary = document.getElementById('summary')
const preview = document.getElementById('preview')

function renderSummary(data) {
  summary.innerHTML = `
    <div class="summary-grid">
      <div class="summary-card">
        <div class="summary-title">Filial</div>
        <div class="summary-value">${data.filial}</div>
      </div>

      <div class="summary-card">
        <div class="summary-title">Arquivo</div>
        <div class="summary-value">${data.arquivo}</div>
      </div>

      <div class="summary-card">
        <div class="summary-title">Upload ID</div>
        <div class="summary-value small">${data.uploadId}</div>
      </div>

      <div class="summary-card">
        <div class="summary-title">Abas reconhecidas</div>
        <div class="summary-value">${data.totalAbas}</div>
      </div>

      <div class="summary-card">
        <div class="summary-title">FRETES gravados</div>
        <div class="summary-value">${data.fretes}</div>
      </div>

      <div class="summary-card">
        <div class="summary-title">CUSTO FROTA gravados</div>
        <div class="summary-value">${data.custoFrota}</div>
      </div>

      <div class="summary-card">
        <div class="summary-title">OCIOSIDADE gravados</div>
        <div class="summary-value">${data.ociosidade}</div>
      </div>
    </div>
  `
}

btnUpload.addEventListener('click', async () => {
  let uploadRecord = null

  try {
    const file = fileInput.files[0]
    const filial = filialSelect.value

    if (!filial) {
      status.innerText = 'Selecione a filial.'
      return
    }

    if (!file) {
      status.innerText = 'Selecione um arquivo Excel.'
      return
    }

    summary.innerHTML = ''
    preview.innerText = ''

    status.innerText = '1/5 Lendo planilha...'
    const analysis = await readExcelFile(file, filial)

    if (
      analysis.fretes.length === 0 &&
      analysis.custoFrota.length === 0 &&
      analysis.ociosidade.length === 0
    ) {
      throw new Error('Planilha inválida: nenhuma aba útil foi identificada.')
    }

    status.innerText = '2/5 Enviando arquivo original para o Storage...'
    const uploadStorage = await uploadFileToStorage(file, filial)

    status.innerText = '3/5 Registrando upload no banco...'
    uploadRecord = await createUploadRecord({
      filial,
      nomeArquivo: file.name,
      caminhoStorage: uploadStorage.path
    })

    status.innerText = '4/5 Gravando FRETES...'
    const totalFretes = await insertFretes(analysis.fretes, uploadRecord.id)

    status.innerText = '4/5 Gravando CUSTO FROTA...'
    const totalCusto = await insertCustoFrota(analysis.custoFrota, uploadRecord.id)

    status.innerText = '4/5 Gravando OCIOSIDADE...'
    const totalOciosidade = await insertOciosidade(analysis.ociosidade, uploadRecord.id)

    await updateUploadStatus(uploadRecord.id, {
      total_fretes: totalFretes,
      total_custo: totalCusto,
      total_ociosidade: totalOciosidade,
      status: 'SUCESSO',
      mensagem_erro: null
    })

    status.innerText = '5/5 Processamento concluído com sucesso.'

    renderSummary({
      filial,
      arquivo: file.name,
      uploadId: uploadRecord.id,
      totalAbas: analysis.sheets.length,
      fretes: totalFretes,
      custoFrota: totalCusto,
      ociosidade: totalOciosidade
    })

    preview.innerText = JSON.stringify(
      {
        filial,
        arquivo: file.name,
        storagePath: uploadStorage.path,
        uploadId: uploadRecord.id,
        abasLidas: analysis.sheets,
        totaisInseridos: {
          fretes: totalFretes,
          custoFrota: totalCusto,
          ociosidade: totalOciosidade
        }
      },
      null,
      2
    )
  } catch (e) {
    console.error(e)

    if (uploadRecord?.id) {
      await updateUploadStatus(uploadRecord.id, {
        status: 'ERRO',
        mensagem_erro: e.message
      })
    }

    status.innerText = `Erro: ${e.message}`
  }
})