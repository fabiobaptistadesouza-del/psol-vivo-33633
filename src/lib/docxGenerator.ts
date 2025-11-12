import { Document, Packer, Paragraph, TextRun, Table, TableCell, TableRow, WidthType, AlignmentType, BorderStyle } from 'docx';
import { QuoteProduct, ClientContacts } from '@/types';
import { formatCurrency } from './formatters';

export const generateProposalDOCX = async (
  quoteNumber: string,
  clientName: string,
  clientCNPJ: string,
  products: QuoteProduct[],
  contacts?: ClientContacts
) => {
  // Calcular total
  const totalValue = products.reduce((sum, product) => 
    sum + (product.precoVenda * product.quantidade), 0
  );

  // Criar documento
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // Página 1: Título e informações do cliente
          new Paragraph({
            text: 'Proposta Comercial',
            heading: 'Heading1',
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: clientName,
                size: 30,
              }),
            ],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `CNPJ: ${clientCNPJ}`,
                size: 30,
              }),
            ],
            spacing: { after: 400 },
          }),
          // Page break
          new Paragraph({
            text: '',
            pageBreakBefore: true,
          }),
          
          // Página 2: Tabela de produtos
          new Paragraph({
            text: 'Lista de Produtos',
            heading: 'Heading2',
            spacing: { after: 200 },
          }),
          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            rows: [
              // Cabeçalho
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: 'Fabricante', bold: true })],
                    shading: { fill: 'CCCCCC' },
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: 'Descrição', bold: true })],
                    shading: { fill: 'CCCCCC' },
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: 'Valor Unit. Venda', bold: true })],
                    shading: { fill: 'CCCCCC' },
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: 'Qtd', bold: true })],
                    shading: { fill: 'CCCCCC' },
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: 'Valor de Venda', bold: true })],
                    shading: { fill: 'CCCCCC' },
                  }),
                ],
              }),
              // Linhas de produtos
              ...products.map(product => 
                new TableRow({
                  children: [
                    new TableCell({
                      children: [new Paragraph(product.fabricante)],
                    }),
                    new TableCell({
                      children: [new Paragraph(product.descricao)],
                    }),
                    new TableCell({
                      children: [new Paragraph(formatCurrency(product.precoVenda))],
                    }),
                    new TableCell({
                      children: [new Paragraph(product.quantidade.toString())],
                    }),
                    new TableCell({
                      children: [new Paragraph(formatCurrency(product.precoVenda * product.quantidade))],
                    }),
                  ],
                })
              ),
            ],
          }),
          new Paragraph({
            text: `Valor Total: ${formatCurrency(totalValue)}`,
            spacing: { before: 200, after: 400 },
            bold: true,
          }),
        ],
      },
      // Página 3: Contatos (se houver)
      ...(contacts ? [{
        properties: {},
        children: [
          new Paragraph({
            text: 'Informações de Contato',
            heading: 'Heading2',
            spacing: { after: 200 },
          }),
          new Paragraph({
            text: 'Cliente',
            heading: 'Heading3',
            spacing: { after: 100 },
          }),
          new Paragraph({
            text: `Nome: ${contacts.cliente.nome}`,
            spacing: { after: 50 },
          }),
          new Paragraph({
            text: `Email: ${contacts.cliente.email}`,
            spacing: { after: 50 },
          }),
          new Paragraph({
            text: `Telefone: ${contacts.cliente.telefone}`,
            spacing: { after: 200 },
          }),
          
          new Paragraph({
            text: 'Pré Vendas',
            heading: 'Heading3',
            spacing: { after: 100 },
          }),
          new Paragraph({
            text: `Nome: ${contacts.preVendas.nome}`,
            spacing: { after: 50 },
          }),
          new Paragraph({
            text: `Email: ${contacts.preVendas.email}`,
            spacing: { after: 50 },
          }),
          new Paragraph({
            text: `Telefone: ${contacts.preVendas.telefone}`,
            spacing: { after: 200 },
          }),
          
          new Paragraph({
            text: 'Gerente de Negócios',
            heading: 'Heading3',
            spacing: { after: 100 },
          }),
          new Paragraph({
            text: `Nome: ${contacts.gerenteNegocios.nome}`,
            spacing: { after: 50 },
          }),
          new Paragraph({
            text: `Email: ${contacts.gerenteNegocios.email}`,
            spacing: { after: 50 },
          }),
          new Paragraph({
            text: `Telefone: ${contacts.gerenteNegocios.telefone}`,
          }),
        ],
      }] : []),
    ],
  });

  // Gerar e fazer download do arquivo
  const blob = await Packer.toBlob(doc);
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `Proposta comercial-${quoteNumber}.docx`;
  link.click();
  window.URL.revokeObjectURL(url);
};
