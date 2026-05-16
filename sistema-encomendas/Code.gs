const SPREADSHEET_ID = '1V8tBYJ26kEnsvljkCFI91BcQl-tM_D2e8sxFSD03HIQ';
const ABA_PEDIDOS = 'Pedidos';
const ABA_PRODUTOS = 'Produtos';

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Queijo Verde Pedidos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getProdutos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(ABA_PRODUTOS);
  const last = Math.max(sh.getLastRow() - 1, 1);
  const values = sh.getRange(2, 1, last, 4).getValues();
  return values
    .filter(row => row[0] && String(row[3]).toLowerCase() !== 'não')
    .map(row => ({ produto: row[0], unidade: row[1], preco: Number(row[2]) || 0, ativo: row[3] }));
}

function salvarPedido(dados) {
  if (!dados || !dados.cliente || !dados.telefone || !dados.produto || !dados.quantidade) {
    throw new Error('Preencha nome, telefone, produto e quantidade.');
  }
  if (dados.entrega === 'Entrega' && !dados.endereco) {
    throw new Error('Informe o endereço para entrega.');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(ABA_PEDIDOS);
  const produtos = getProdutos();
  const item = produtos.find(p => p.produto === dados.produto);
  const preco = item ? item.preco : 0;
  const quantidade = Number(dados.quantidade) || 0;
  const total = preco * quantidade;
  const hoje = new Date();
  const sexta = calcularProximaSexta(hoje);

  sh.appendRow([
    hoje,
    sexta,
    dados.cliente,
    dados.telefone,
    dados.entrega || 'Retirada',
    dados.produto,
    quantidade,
    preco,
    total,
    dados.pagamento || 'Pix',
    dados.status || 'Pendente',
    dados.observacoes || '',
    dados.registradoPor || 'Sistema'
  ]);

  return { ok: true, mensagem: 'Pedido cadastrado com sucesso!', total, sexta: formatarData(sexta) };
}

function listarPedidos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(ABA_PEDIDOS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 13).getValues();
  const sextaKey = dataKey(calcularProximaSexta(new Date()));
  return values
    .map((row, i) => ({ row, linha: i + 2 }))
    .filter(obj => obj.row[2])
    .filter(obj => dataKey(obj.row[1]) === sextaKey)
    .map(obj => ({
      linha: obj.linha,
      dataPedido: formatarData(obj.row[0]),
      sextaEntrega: formatarData(obj.row[1]),
      cliente: obj.row[2],
      telefone: obj.row[3],
      entrega: obj.row[4],
      produto: obj.row[5],
      quantidade: obj.row[6],
      preco: Number(obj.row[7]) || 0,
      total: Number(obj.row[8]) || 0,
      pagamento: obj.row[9],
      status: obj.row[10],
      observacoes: obj.row[11],
      registradoPor: obj.row[12]
    }))
    .sort((a, b) => String(a.cliente).localeCompare(String(b.cliente)));
}

function resumoSexta() {
  const pedidos = listarPedidos();
  const resumo = { totalPedidos: pedidos.length, receita: 0, queijo1kg: 0, queijo500g: 0, leite: 0, pendentes: 0, entregues: 0 };
  pedidos.forEach(p => {
    if (p.status !== 'Cancelado') resumo.receita += Number(p.total) || 0;
    if (p.produto === 'Queijo 1kg') resumo.queijo1kg += Number(p.quantidade) || 0;
    if (p.produto === 'Queijo 500g') resumo.queijo500g += Number(p.quantidade) || 0;
    if (p.produto === 'Leite') resumo.leite += Number(p.quantidade) || 0;
    if (p.status === 'Pendente') resumo.pendentes += 1;
    if (p.status === 'Entregue') resumo.entregues += 1;
  });
  return resumo;
}

function atualizarStatus(linha, status) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(ABA_PEDIDOS);
  sh.getRange(Number(linha), 11).setValue(status);
  return { ok: true, mensagem: 'Status atualizado.' };
}

function calcularProximaSexta(data) {
  const d = new Date(data);
  const dia = d.getDay();
  const diasAteSexta = (5 - dia + 7) % 7;
  d.setDate(d.getDate() + diasAteSexta);
  d.setHours(0, 0, 0, 0);
  return d;
}

function dataKey(data) {
  return Utilities.formatDate(new Date(data), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatarData(data) {
  if (!data) return '';
  return Utilities.formatDate(new Date(data), Session.getScriptTimeZone(), 'dd/MM/yyyy');
}
