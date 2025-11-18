import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Card } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Button } from '@/components/ui/button';
import { useApp } from '@/contexts/AppContext';
import { formatCurrency } from '@/lib/formatters';
import { generateProposalPDF } from '@/lib/pdfGenerator';
import { generateProposalDOCX } from '@/lib/docxGenerator';
import { toast } from 'sonner';
import * as XLSX from 'xlsx';
import { Download, FileText, ChevronDown } from 'lucide-react';
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from '@/components/ui/collapsible';

// Tabela de coeficientes baseada no PRV
const COEFICIENTES: Record<number, number> = {
  30: 1.01698633663687,
  45: 1.03426120890608,
  60: 1.03426120890608,
  75: 1.05182951797102,
  90: 1.05182951797102,
};

export default function ResumoFinanceiro() {
  const navigate = useNavigate();
  const { selectedQuote, quoteProducts, rateioServices, quoteConfigs, adminSettings, clientContacts, clients } = useApp();

  const [receitaBruta, setReceitaBruta] = useState(0);
  const [impostos, setImpostos] = useState(0);
  const [receitaLiquida, setReceitaLiquida] = useState(0);
  const [custoProduto, setCustoProduto] = useState(0);
  const [rateio, setRateio] = useState(0);
  const [custoTotal, setCustoTotal] = useState(0);
  const [margemDireta, setMargemDireta] = useState(0);
  const [inadimplencia, setInadimplencia] = useState(0);
  const [ebitda, setEbitda] = useState(0);
  const [margemEbitda, setMargemEbitda] = useState(0);
  const [irCsll, setIrCsll] = useState(0);
  const [lucroLiquido, setLucroLiquido] = useState(0);
  const [margemLiquida, setMargemLiquida] = useState(0);
  const [vpl, setVpl] = useState(0);
  const [alcada, setAlcada] = useState('');
  const [lucroLiquidoMesPrv, setLucroLiquidoMesPrv] = useState(0);
  const [vplPrv, setVplPrv] = useState(0);
  const [vplDoProjeto, setVplDoProjeto] = useState(0);

  // Estados para cálculos de mês zero
  const [margemDiretaMesZero, setMargemDiretaMesZero] = useState(0);
  const [ebitdaMesZero, setEbitdaMesZero] = useState(0);
  const [margemEbitMesZero, setMargemEbitMesZero] = useState(0);
  const [irCsllMesZero, setIrCsllMesZero] = useState(0);
  const [lucroLiquidoMesZero, setLucroLiquidoMesZero] = useState(0);
  const [vplMesUm, setVplMesUm] = useState(0);

  // Estados para controle de expansão das seções
  const [isCustoOpen, setIsCustoOpen] = useState(false);
  const [isReceitaOpen, setIsReceitaOpen] = useState(false);

  useEffect(() => {
    if (!selectedQuote) {
      toast.error('Selecione uma cotação primeiro');
      navigate('/');
    }
  }, [selectedQuote, navigate]);

  useEffect(() => {
    if (!selectedQuote) return;

    const products = quoteProducts[selectedQuote.id] || [];
    const services = rateioServices[selectedQuote.id] || [];
    const config = quoteConfigs[selectedQuote.id];

    // Receita Bruta = valor total da Máscara do Fornecedor
    const receitaBrutaCalc = products.reduce((sum, p) => sum + (p.precoVenda * p.quantidade), 0);
    setReceitaBruta(receitaBrutaCalc);

    // Impostos = Receita Bruta × 0,25
    const impostosCalc = receitaBrutaCalc * 0.25;
    setImpostos(impostosCalc);

    // Receita Líquida = Receita Bruta
    const receitaLiquidaCalc = receitaBrutaCalc;
    setReceitaLiquida(receitaLiquidaCalc);

    // Custo do Produto = soma da coluna "custo" (custoUnitario * quantidade)
    const custoProdutoCalc = products.reduce((sum, p) => sum + (p.custoUnitario * p.quantidade), 0);
    setCustoProduto(custoProdutoCalc);

    // Rateio = valor total da página Rateio
    const rateioCalc = services.reduce((sum, s) => sum + s.valorComImpostos, 0);
    setRateio(rateioCalc);

    // Custo Total = Custo do Produto + Rateio
    const custoTotalCalc = custoProdutoCalc + rateioCalc;
    setCustoTotal(custoTotalCalc);

    // Margem Direta = Receita Bruta
    const margemDiretaCalc = receitaBrutaCalc;
    setMargemDireta(margemDiretaCalc);

    // Inadimplência = 0 (fixo)
    const inadimplenciaCalc = 0;
    setInadimplencia(inadimplenciaCalc);

    // EBITDA = Receita Bruta
    const ebitdaCalc = receitaBrutaCalc;
    setEbitda(ebitdaCalc);

    // Margem EBITDA = (EBITDA ÷ Receita Líquida) × 100
    const margemEbitdaCalc = receitaLiquidaCalc > 0 ? (ebitdaCalc / receitaLiquidaCalc) * 100 : 0;
    setMargemEbitda(margemEbitdaCalc);

    // IR/CSLL (34%) = 0,34 × EBITDA
    const irCsllCalc = ebitdaCalc * 0.34;
    setIrCsll(irCsllCalc);

    // Lucro líquido - mês PRV = EBITDA − IR/CSLL (34%)
    const lucroLiquidoMesPrvCalc = ebitdaCalc - irCsllCalc;
    setLucroLiquidoMesPrv(lucroLiquidoMesPrvCalc);

    // Cálculos de mês zero
    // Margem Direta - mês zero = Impostos + Custo Total
    const margemDiretaMesZeroCalc = impostosCalc + custoTotalCalc;
    setMargemDiretaMesZero(margemDiretaMesZeroCalc);

    // EBITDA - mês zero = Impostos + Custo Total
    const ebitdaMesZeroCalc = impostosCalc + custoTotalCalc;
    setEbitdaMesZero(ebitdaMesZeroCalc);

    // Margem EBIT - mês zero = Impostos + Custo Total
    const margemEbitMesZeroCalc = impostosCalc + custoTotalCalc;
    setMargemEbitMesZero(margemEbitMesZeroCalc);

    // IR/CSLL - mês zero = 0,34 × Margem Ebit - mês zero
    const irCsllMesZeroCalc = margemEbitMesZeroCalc * 0.34;
    setIrCsllMesZero(irCsllMesZeroCalc);

    // Lucro líquido - mês zero = IR/CSLL - mês zero − Margem Ebit - mês zero
    const lucroLiquidoMesZeroCalc = irCsllMesZeroCalc - margemEbitMesZeroCalc;
    setLucroLiquidoMesZero(lucroLiquidoMesZeroCalc);

    // VPL - mês um = Lucro líquido - mês zero ÷ 1.01698633663687
    const vplMesUmCalc = lucroLiquidoMesZeroCalc / 1.01698633663687;
    setVplMesUm(vplMesUmCalc);

    // VPL (PRV) = Lucro líquido - mês PRV ÷ coeficiente
    const prv = config?.prv || 30;
    const coeficiente = COEFICIENTES[prv] || COEFICIENTES[30];
    const vplPrvCalc = lucroLiquidoMesPrvCalc / coeficiente;
    setVplPrv(vplPrvCalc);

    // Lucro Líquido = Lucro líquido - mês PRV + Lucro líquido - mês zero
    const lucroLiquidoCalc = lucroLiquidoMesPrvCalc + lucroLiquidoMesZeroCalc;
    setLucroLiquido(lucroLiquidoCalc);

    // VPL do projeto = VPL (PRV) + VPL - mês um
    const vplDoProjetoCalc = vplPrvCalc + vplMesUmCalc;
    setVplDoProjeto(vplDoProjetoCalc);

    // Margem Líquida = Lucro Líquido ÷ (Receita Líquida - Impostos) × 100
    const denominador = receitaLiquidaCalc - impostosCalc;
    const margemLiquidaCalc = denominador > 0 ? (lucroLiquidoCalc / denominador) * 100 : 0;
    setMargemLiquida(margemLiquidaCalc);

    // Alçada de aprovação baseada na margem líquida
    const alcadas = adminSettings.alcadas;
    let alcadaCalc = '';
    
    if (margemLiquidaCalc >= alcadas.preVendas) {
      alcadaCalc = 'Pré Vendas';
    } else if (margemLiquidaCalc >= alcadas.diretor) {
      alcadaCalc = 'Diretor';
    } else {
      alcadaCalc = 'CDG';
    }
    setAlcada(alcadaCalc);

  }, [selectedQuote, quoteProducts, rateioServices, quoteConfigs, adminSettings]);

  if (!selectedQuote) {
    return null;
  }

  const client = clients.find(c => c.id === selectedQuote.clientId);

  const handleDownloadExcel = async () => {
    try {
      // Carregar o template ou arquivo da administração
      let arrayBuffer: ArrayBuffer;
      if (adminSettings.cashFlowFile?.fileData) {
        arrayBuffer = adminSettings.cashFlowFile.fileData;
      } else {
        const response = await fetch('/Template_FluxoCaixa.xlsx');
        arrayBuffer = await response.arrayBuffer();
      }

      const wb = XLSX.read(arrayBuffer, { type: 'array' });

      // Atualizar células na aba "FC"
      if (wb.SheetNames.includes('FC')) {
        const wsFC = wb.Sheets['FC'];
        
        const config = quoteConfigs[selectedQuote.id];
        const prv = config?.prv || 30;

        const fcUpdates = {
          'B13': receitaBruta,
          'B14': prv,
          'B15': custoProduto,
          'B16': rateio,
          'B18': impostos,
        };

        Object.entries(fcUpdates).forEach(([cellRef, value]) => {
          if (!wsFC[cellRef]) {
            wsFC[cellRef] = { t: 'n', v: value };
          } else {
            wsFC[cellRef].v = value;
            wsFC[cellRef].t = 'n';
          }
        });
      }

      // Download
      XLSX.writeFile(wb, `Fluxo de caixa-${selectedQuote.codigo}.xlsx`);
      toast.success('Excel gerado com sucesso');
    } catch (error) {
      console.error('Erro ao gerar Excel:', error);
      toast.error('Erro ao gerar Excel');
    }
  };

  const handleViewProposal = () => {
    const products = quoteProducts[selectedQuote.id] || [];
    const contacts = clientContacts[selectedQuote.id];
    
    if (products.length === 0) {
      toast.error('Adicione produtos antes de gerar a proposta');
      return;
    }

    if (!client) {
      toast.error('Cliente não encontrado');
      return;
    }

    generateProposalPDF(
      selectedQuote.codigo,
      client.razaoSocial,
      client.cnpj,
      products,
      contacts
    );
    
    toast.success('Proposta gerada com sucesso');
  };

  const handleViewProposalWord = async () => {
    const products = quoteProducts[selectedQuote.id] || [];
    const contacts = clientContacts[selectedQuote.id];
    
    if (products.length === 0) {
      toast.error('Adicione produtos antes de gerar a proposta');
      return;
    }

    if (!client) {
      toast.error('Cliente não encontrado');
      return;
    }

    try {
      await generateProposalDOCX(
        selectedQuote.codigo,
        client.razaoSocial,
        client.cnpj,
        products,
        contacts
      );
      toast.success('Proposta Word gerada com sucesso');
    } catch (error) {
      console.error('Erro ao gerar proposta Word:', error);
      toast.error('Erro ao gerar proposta Word');
    }
  };

  return (
    <div className="p-6 space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-foreground">Resumo Financeiro</h1>
        <p className="text-muted-foreground">Visão consolidada dos cálculos financeiros</p>
      </div>

      {/* Seção 1: Custo + impostos */}
      <Card className="p-6">
        <Collapsible open={isCustoOpen} onOpenChange={setIsCustoOpen}>
          <CollapsibleTrigger className="flex items-center justify-between w-full">
            <h2 className="text-lg font-semibold">Custo + impostos</h2>
            <ChevronDown 
              className={`h-5 w-5 transition-transform duration-200 ${isCustoOpen ? 'rotate-0' : '-rotate-90'}`}
            />
          </CollapsibleTrigger>
          <CollapsibleContent className="mt-4">
            <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">
              <div>
                <Label htmlFor="impostos">Impostos</Label>
                <Input
                  id="impostos"
                  type="text"
                  value={formatCurrency(impostos)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="custoProduto">Custo do Produto</Label>
                <Input
                  id="custoProduto"
                  type="text"
                  value={formatCurrency(custoProduto)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="rateio">Rateio</Label>
                <Input
                  id="rateio"
                  type="text"
                  value={formatCurrency(rateio)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="custoTotal">Custo Total</Label>
                <Input
                  id="custoTotal"
                  type="text"
                  value={formatCurrency(custoTotal)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="margemDiretaMesZero">Margem Direta - mês um</Label>
                <Input
                  id="margemDiretaMesZero"
                  type="text"
                  value={formatCurrency(margemDiretaMesZero)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="ebitdaMesZero">Ebitda - mês um</Label>
                <Input
                  id="ebitdaMesZero"
                  type="text"
                  value={formatCurrency(ebitdaMesZero)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="margemEbitMesZero">Margem Ebit - mês um</Label>
                <Input
                  id="margemEbitMesZero"
                  type="text"
                  value={formatCurrency(margemEbitMesZero)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="irCsllMesZero">IR/CSLL (34%) - mês um</Label>
                <Input
                  id="irCsllMesZero"
                  type="text"
                  value={formatCurrency(irCsllMesZero)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="lucroLiquidoMesZero">Lucro líquido - mês um</Label>
                <Input
                  id="lucroLiquidoMesZero"
                  type="text"
                  value={formatCurrency(lucroLiquidoMesZero)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="vplMesUm">VPL - mês um</Label>
                <Input
                  id="vplMesUm"
                  type="text"
                  value={formatCurrency(vplMesUm)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>
            </div>
          </CollapsibleContent>
        </Collapsible>
      </Card>

      {/* Seção 2: Receita */}
      <Card className="p-6">
        <Collapsible open={isReceitaOpen} onOpenChange={setIsReceitaOpen}>
          <CollapsibleTrigger className="flex items-center justify-between w-full">
            <h2 className="text-lg font-semibold">Receita</h2>
            <ChevronDown 
              className={`h-5 w-5 transition-transform duration-200 ${isReceitaOpen ? 'rotate-0' : '-rotate-90'}`}
            />
          </CollapsibleTrigger>
          <CollapsibleContent className="mt-4">
            <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">
              <div>
                <Label htmlFor="receitaBruta">Receita Bruta</Label>
                <Input
                  id="receitaBruta"
                  type="text"
                  value={formatCurrency(receitaBruta)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="receitaLiquida">Receita Líquida</Label>
                <Input
                  id="receitaLiquida"
                  type="text"
                  value={formatCurrency(receitaLiquida)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="margemDireta">Margem Direta</Label>
                <Input
                  id="margemDireta"
                  type="text"
                  value={formatCurrency(margemDireta)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="inadimplencia">Inadimplência</Label>
                <Input
                  id="inadimplencia"
                  type="text"
                  value={formatCurrency(inadimplencia)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="ebitda">EBITDA</Label>
                <Input
                  id="ebitda"
                  type="text"
                  value={formatCurrency(ebitda)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="irCsll">IR/CSLL (34%)</Label>
                <Input
                  id="irCsll"
                  type="text"
                  value={formatCurrency(irCsll)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="lucroLiquidoMesPrv">Lucro líquido - mês PRV</Label>
                <Input
                  id="lucroLiquidoMesPrv"
                  type="text"
                  value={formatCurrency(lucroLiquidoMesPrv)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>

              <div>
                <Label htmlFor="vplPrv">VPL (PRV)</Label>
                <Input
                  id="vplPrv"
                  type="text"
                  value={formatCurrency(vplPrv)}
                  placeholder="R$ 0,00"
                  className="mt-2 bg-muted"
                  readOnly
                />
              </div>
            </div>
          </CollapsibleContent>
        </Collapsible>
      </Card>

      {/* Campos de Destaque */}
      <div className="grid gap-6 md:grid-cols-3">
        <Card className="p-6 border-2 border-primary">
          <Label className="text-sm text-muted-foreground">Lucro Líquido</Label>
          <div className="mt-2 text-3xl font-bold text-primary">
            {formatCurrency(lucroLiquido)}
          </div>
        </Card>

        <Card className="p-6 border-2 border-primary">
          <Label className="text-sm text-muted-foreground">Margem Líquida</Label>
          <div className="mt-2 text-3xl font-bold text-primary">
            {margemLiquida.toFixed(1)}%
          </div>
        </Card>

        <Card className="p-6 border-2 border-primary">
          <Label className="text-sm text-muted-foreground">VPL do projeto</Label>
          <div className="mt-2 text-3xl font-bold text-primary">
            {formatCurrency(vplDoProjeto)}
          </div>
        </Card>
      </div>

      {/* Alçada de Aprovação */}
      <Card className="p-6">
        <h2 className="text-lg font-semibold mb-4">Alçada de Aprovação</h2>
        <div className="grid gap-4 md:grid-cols-2">
          <div>
            <Label htmlFor="alcada">Alçada de Aprovação</Label>
            <Input
              id="alcada"
              type="text"
              value={alcada}
              placeholder="Não definida"
              className="mt-2 bg-muted font-semibold"
              readOnly
            />
          </div>
        </div>
      </Card>

      {/* Botões de Ação */}
      <div className="flex gap-4">
        <Button onClick={handleDownloadExcel} className="flex-1">
          <Download className="mr-2 h-4 w-4" />
          Baixar Fluxo de Caixa
        </Button>
        <Button onClick={handleViewProposal} className="flex-1">
          <FileText className="mr-2 h-4 w-4" />
          Visualizar Proposta (PDF)
        </Button>
        <Button onClick={handleViewProposalWord} variant="secondary" className="flex-1">
          <FileText className="mr-2 h-4 w-4" />
          Visualizar Proposta (Word)
        </Button>
      </div>
    </div>
  );
}
