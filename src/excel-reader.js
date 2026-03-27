import * as XLSX from 'xlsx'

const CAP_RULES = {
  FIORINO: 650,
  VAN: 1500,
  VUC: 2500,
  '3/4': 4000,
  TOCO: 7500,
  TRUCK: 12000,
  BITRUCK: 16000,
  CARRETA: 28000
}

function normalizeText(value) {
  return String(value ?? '').trim()
}

function normalizeHeader(value) {
  return normalizeText(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\?/g, '')
    .replace(/\s+/g, ' ')
    .toUpperCase()
}

function classifySheet(name) {
  const upper = normalizeHeader(name)

  if (upper.includes('FRETES')) return 'FRETES'
  if (upper.includes('CUSTO')) return 'CUSTO_FROTA'
  if (upper.includes('OCIOSIDADE')) return 'OCIOSIDADE'

  return 'IGNORAR'
}

function detectHeaderRow(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ''
  })

  let bestIndex = 0
  let bestScore = 0

  const expected = [
    'DATA',
    'EMP',
    'ORIGEM',
    'DESTINO',
    'MOTORISTA',
    'PLACA',
    'TIPO',
    'CAP',
    'TOTAL',
    'CD',
    'OCIOSIDADE'
  ]

  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i].map(normalizeHeader)
    const score = expected.filter((item) => row.includes(item)).length

    if (score > bestScore) {
      bestScore = score
      bestIndex = i
    }
  }

  return bestIndex
}

function sheetToObjects(sheet) {
  const headerRow = detectHeaderRow(sheet)

  return XLSX.utils.sheet_to_json(sheet, {
    defval: null,
    range: headerRow,
    raw: true
  })
}

function getValue(row, aliases = []) {
  for (const alias of aliases) {
    if (row[alias] !== undefined && row[alias] !== null && row[alias] !== '') {
      return row[alias]
    }
  }
  return null
}

function isRowEmpty(obj) {
  return Object.values(obj).every((value) => {
    return value === null || value === undefined || String(value).trim() === ''
  })
}

function excelDateToISO(serial) {
  const n = Number(serial)
  if (Number.isNaN(n)) return null

  const utcDays = Math.floor(n - 25569)
  const utcValue = utcDays * 86400
  const dateInfo = new Date(utcValue * 1000)

  if (Number.isNaN(dateInfo.getTime())) return null

  return dateInfo.toISOString().split('T')[0]
}

function parseDateString(str) {
  const value = String(str).trim()
  if (!value) return null

  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(value)) {
    const [day, month, year] = value.split('/')
    return `${year}-${month}-${day}`
  }

  if (/^\d{2}-\d{2}-\d{4}$/.test(value)) {
    const [day, month, year] = value.split('-')
    return `${year}-${month}-${day}`
  }

  const parsed = new Date(value)
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString().split('T')[0]
  }

  return null
}

function toDateString(value) {
  if (value === null || value === undefined || value === '') return null

  if (typeof value === 'number') {
    return excelDateToISO(value)
  }

  const parsed = parseDateString(value)
  return parsed || String(value).trim()
}

function toText(value) {
  if (value === null || value === undefined || value === '') return null
  return String(value).trim()
}

function toUpperText(value) {
  const text = toText(value)
  return text ? text.toUpperCase() : null
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return null

  if (typeof value === 'number') {
    return Number.isNaN(value) ? null : value
  }

  let str = String(value).trim()
  if (!str) return null

  str = str.replace(/R\$/gi, '').trim()

  const hasComma = str.includes(',')
  const hasDot = str.includes('.')

  if (hasComma && hasDot) {
    str = str.replace(/\./g, '').replace(',', '.')
  } else if (hasComma && !hasDot) {
    str = str.replace(',', '.')
  }

  str = str.replace(/[^\d.-]/g, '')

  const num = Number(str)
  return Number.isNaN(num) ? null : num
}

function buildKey(parts) {
  return parts
    .map((part) => String(part ?? '').trim().toUpperCase())
    .join('|')
}

function cleanFretesRow(row) {
  return !isRowEmpty(row) && (
    getValue(row, ['Data', 'DATA']) ||
    getValue(row, ['Motorista', 'MOTORISTA']) ||
    getValue(row, ['Placa', 'PLACA']) ||
    getValue(row, ['Destino', 'DESTINO'])
  )
}

function cleanCustoRow(row) {
  return !isRowEmpty(row) && (
    getValue(row, ['Mês', 'MES', 'MÊS']) ||
    getValue(row, ['CD']) ||
    getValue(row, ['Motorista', 'MOTORISTA']) ||
    getValue(row, ['Placa', 'PLACA'])
  )
}

function cleanOciosidadeRow(row) {
  return !isRowEmpty(row) && (
    getValue(row, ['DATA', 'Data']) ||
    getValue(row, ['CD']) ||
    getValue(row, ['PLACA', 'Placa']) ||
    getValue(row, ['MOTORISTA', 'Motorista'])
  )
}

function mapFretesRow(row, filial, origemAba) {
  const data = toDateString(getValue(row, ['Data', 'DATA']))
  const motorista = toText(getValue(row, ['Motorista', 'MOTORISTA']))
  const placa = toUpperText(getValue(row, ['Placa', 'PLACA']))
  const destino = toText(getValue(row, ['Destino', 'DESTINO']))
  const peso = toNumber(getValue(row, ['Peso', 'PESO']))

  const tipoRaw = toUpperText(getValue(row, ['Tipo', 'TIPO']))
  const capRaw = toNumber(getValue(row, ['Cap', 'CAP']))

  let tipo = tipoRaw
  let cap = capRaw
  let erroValidacao = null

  if (tipo && CAP_RULES[tipo]) {
    const capEsperada = CAP_RULES[tipo]

    if (cap !== capEsperada) {
      erroValidacao = `Capacidade ajustada automaticamente (${cap ?? 'vazio'} -> ${capEsperada})`
      cap = capEsperada
    }
  }

  if (tipo && !CAP_RULES[tipo]) {
    erroValidacao = `Tipo não mapeado: ${tipo}`
  }

  return {
    filial,
    origem_aba: origemAba,
    data,
    emp: toText(getValue(row, ['Emp', 'EMP'])),
    origem: toText(getValue(row, ['Origem', 'ORIGEM'])),
    destino,
    rota: toText(getValue(row, ['Rota', 'ROTA'])),
    canhoto_entregue: toText(getValue(row, ['Canhoto Entregue ?', 'CANHOTO ENTREGUE', 'Canhoto Entregue'])),
    uf: toUpperText(getValue(row, ['UF'])),
    transportador: toText(getValue(row, ['TRANSPORTADOR', 'Transportador', 'TRANSPORTADORA'])),
    motorista,
    placa,
    tipo,
    cap,
    mod: toText(getValue(row, ['Mod', 'MOD'])),
    peso,
    qt_entregas: toNumber(getValue(row, ['Qt Entregas', 'QT ENTREGAS'])),
    km_saida: toNumber(getValue(row, ['Km Saida', 'KM SAIDA'])),
    km_fim: toNumber(getValue(row, ['Km FIM', 'KM FIM', 'KM Fim'])),
    km_rod: toNumber(getValue(row, ['Km Rod', 'KM ROD'])),
    vr_saida: toNumber(getValue(row, ['VR Saida', 'VR SAIDA'])),
    vr_km: toNumber(getValue(row, ['VR km', 'VR KM'])),
    vr_km_rod: toNumber(getValue(row, ['VR KM Rod', 'VR KM ROD'])),
    qt_pernoite: toNumber(getValue(row, ['Qt Pernoite', 'QT PERNOITE'])),
    pernoite: toNumber(getValue(row, ['Pernoite', 'PERNOITE'])),
    pedagio: toNumber(getValue(row, ['Pedágio', 'PEDAGIO'])),
    ajudante: toNumber(getValue(row, ['Ajudante', 'AJUDANTE'])),
    descarga: toNumber(getValue(row, ['Descarga', 'DESCARGA'])),
    total: toNumber(getValue(row, ['Total', 'TOTAL'])),
    ocupacao: toNumber(getValue(row, ['Ocupação', 'OCUPACAO'])),
    custo_kg: toNumber(getValue(row, ['Custo-KG', 'CUSTO-KG', 'CUSTO KG'])),
    status_entrega: toText(getValue(row, ['Status Entrega', 'STATUS ENTREGA'])),
    data_status: toDateString(getValue(row, ['Data Status', 'DATA STATUS'])),
    tempo_status: toText(getValue(row, ['Tempo Status', 'TEMPO STATUS'])),
    canhoto: toText(getValue(row, ['Canhoto', 'CANHOTO'])),
    erro_validacao: erroValidacao,
    chave_unica: buildKey([filial, data, placa, motorista, destino, peso])
  }
}

function mapCustoRow(row, filial, origemAba) {
  const mes = toText(getValue(row, ['Mês', 'MES', 'MÊS']))
  const placa = toUpperText(getValue(row, ['Placa', 'PLACA']))
  const motorista = toText(getValue(row, ['Motorista', 'MOTORISTA']))
  const tipo = toText(getValue(row, ['Tipo', 'TIPO']))

  return {
    filial,
    origem_aba: origemAba,
    mes,
    emp: toText(getValue(row, ['Emp', 'EMP'])),
    cd: toText(getValue(row, ['CD'])),
    motorista,
    placa,
    tipo,
    peso_tr: toNumber(getValue(row, ['Peso Tr', 'PESO TR'])),
    qt_entregas: toNumber(getValue(row, ['Qt Entregas', 'QT ENTREGAS'])),
    km_saida: toNumber(getValue(row, ['Km Saida', 'KM SAIDA'])),
    km_fim: toNumber(getValue(row, ['KM Fim', 'KM FIM'])),
    km_rod: toNumber(getValue(row, ['Km Rod', 'KM ROD'])),
    salario_mot: toNumber(getValue(row, ['Salário Mot', 'SALARIO MOT'])),
    encargos: toNumber(getValue(row, ['Encargos', 'ENCARGOS'])),
    despesas_viagens: toNumber(getValue(row, ['Despesas Viagens', 'DESPESAS VIAGENS'])),
    pedagio: toNumber(getValue(row, ['Pedágio', 'PEDAGIO'])),
    combustivel: toNumber(getValue(row, ['Combustivel', 'COMBUSTIVEL'])),
    depreciacao: toNumber(getValue(row, ['Depreciação', 'DEPRECIACAO'])),
    outros_gastos: toNumber(getValue(row, ['Outros Gastos', 'OUTROS GASTOS'])),
    manut_geral: toNumber(getValue(row, ['Manut Geral', 'MANUT GERAL'])),
    total: toNumber(getValue(row, ['Total', 'TOTAL'])),
    financ_parcela: toNumber(getValue(row, ['Financ (parcela)', 'FINANC (PARCELA)'])),
    seguro: toNumber(getValue(row, ['Seguro', 'SEGURO'])),
    chave_unica: buildKey([filial, mes, placa, motorista, tipo])
  }
}

function mapOciosidadeRow(row, filial, origemAba) {
  const data = toDateString(getValue(row, ['DATA', 'Data']))
  const placa = toUpperText(getValue(row, ['PLACA', 'Placa']))
  const motorista = toText(getValue(row, ['MOTORISTA', 'Motorista']))

  return {
    filial,
    origem_aba: origemAba,
    mes: toText(getValue(row, ['MÊS', 'MES', 'Mês'])),
    data,
    cd: toText(getValue(row, ['CD'])),
    placa,
    tipo: toText(getValue(row, ['TIPO', 'Tipo'])),
    motorista,
    ociosidade: toNumber(getValue(row, ['OCIOSIDADE', 'Ociosidade'])),
    percentual: toNumber(getValue(row, ['%', 'Percentual'])),
    chave_unica: buildKey([filial, data, placa, motorista])
  }
}

export async function readExcelFile(file, filial) {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellDates: false,
    raw: true
  })

  const sheets = []
  const fretes = []
  const custoFrota = []
  const ociosidade = []

  workbook.SheetNames.forEach((sheetName) => {
    const type = classifySheet(sheetName)

    if (type === 'IGNORAR') return

    const sheet = workbook.Sheets[sheetName]
    const rows = sheetToObjects(sheet)

    sheets.push({
      name: sheetName,
      type,
      totalRows: rows.length
    })

    if (type === 'FRETES') {
      rows
        .filter(cleanFretesRow)
        .map((row) => mapFretesRow(row, filial, sheetName))
        .forEach((item) => fretes.push(item))
    }

    if (type === 'CUSTO_FROTA') {
      rows
        .filter(cleanCustoRow)
        .map((row) => mapCustoRow(row, filial, sheetName))
        .forEach((item) => custoFrota.push(item))
    }

    if (type === 'OCIOSIDADE') {
      rows
        .filter(cleanOciosidadeRow)
        .map((row) => mapOciosidadeRow(row, filial, sheetName))
        .forEach((item) => ociosidade.push(item))
    }
  })

  return {
    sheets,
    fretes,
    custoFrota,
    ociosidade
  }
}