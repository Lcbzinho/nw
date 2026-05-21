const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, PageBreak, VerticalAlign
} = require('docx');
const fs = require('fs');

// Colors
const COLOR = {
  primary: '1A1A2E',
  accent: 'C9A84C',
  danger: 'C0392B',
  warning: 'E67E22',
  success: '27AE60',
  info: '2980B9',
  light: 'F5F5F5',
  border: 'CCCCCC',
  white: 'FFFFFF',
  gray: '666666',
  darkgray: '333333',
};

const border = { style: BorderStyle.SINGLE, size: 1, color: COLOR.border };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, font: 'Arial', size: 36, bold: true, color: COLOR.primary })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 120 },
    children: [new TextRun({ text, font: 'Arial', size: 28, bold: true, color: COLOR.primary })],
  });
}

function heading3(text, color = COLOR.primary) {
  return new Paragraph({
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, font: 'Arial', size: 24, bold: true, color })],
  });
}

function body(text, options = {}) {
  return new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text, font: 'Arial', size: 22, color: COLOR.darkgray, ...options })],
  });
}

function spacer(size = 120) {
  return new Paragraph({ spacing: { after: size }, children: [new TextRun('')] });
}

function dividerLine(color = COLOR.accent) {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 1 } },
    spacing: { before: 60, after: 180 },
    children: [new TextRun('')],
  });
}

function bullet(text, ref = 'bullets') {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, font: 'Arial', size: 22, color: COLOR.darkgray })],
  });
}

function colorBadge(label, bgColor, textColor = COLOR.white) {
  return new Table({
    width: { size: 2500, type: WidthType.DXA },
    columnWidths: [2500],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noBorders,
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 120, right: 120 },
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: label, font: 'Arial', size: 18, bold: true, color: textColor })]
            })]
          })
        ]
      })
    ]
  });
}

function problemRow(severity, area, problem, impact, recommendation, severityColor) {
  const severityBg = { 'CRÍTICO': COLOR.danger, 'ALTO': COLOR.warning, 'MÉDIO': COLOR.info, 'BAIXO': COLOR.success };
  const bg = severityBg[severity] || COLOR.gray;
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 900, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: severity, font: 'Arial', size: 18, bold: true, color: COLOR.white })]
        })]
      }),
      new TableCell({
        borders,
        width: { size: 1200, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [new Paragraph({ children: [new TextRun({ text: area, font: 'Arial', size: 20, bold: true, color: COLOR.primary })] })]
      }),
      new TableCell({
        borders,
        width: { size: 2200, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [new Paragraph({ children: [new TextRun({ text: problem, font: 'Arial', size: 20, color: COLOR.darkgray })] })]
      }),
      new TableCell({
        borders,
        width: { size: 1700, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [new Paragraph({ children: [new TextRun({ text: impact, font: 'Arial', size: 20, color: COLOR.darkgray })] })]
      }),
      new TableCell({
        borders,
        width: { size: 2360, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [new Paragraph({ children: [new TextRun({ text: recommendation, font: 'Arial', size: 20, color: COLOR.darkgray })] })]
      }),
    ]
  });
}

function tableHeader(cols, widths, bgColor = COLOR.primary) {
  return new TableRow({
    tableHeader: true,
    children: cols.map((col, i) =>
      new TableCell({
        borders,
        width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: col, font: 'Arial', size: 20, bold: true, color: COLOR.white })]
        })]
      })
    )
  });
}

function infoBox(title, lines, bgColor = 'EBF5FB', titleColor = COLOR.info) {
  const children = [
    new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: title, font: 'Arial', size: 22, bold: true, color: titleColor })]
    }),
    ...lines.map(l => new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: l, font: 'Arial', size: 21, color: COLOR.darkgray })]
    }))
  ];
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({
      children: [new TableCell({
        borders: noBorders,
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children
      })]
    })]
  });
}

// ===================== DOCUMENT =====================

const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } }
        }]
      },
      {
        reference: 'check',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '✓', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: 'Arial', color: COLOR.primary },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: COLOR.primary },
        paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 1 }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 }
      }
    },
    children: [

      // =========== CAPA ===========
      new Paragraph({ spacing: { before: 1440, after: 600 }, children: [new TextRun('')] }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
        children: [new TextRun({ text: 'RELATÓRIO DE AUDITORIA DIGITAL', font: 'Arial', size: 52, bold: true, color: COLOR.primary })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: 'NW EXECUTIVE', font: 'Arial', size: 48, bold: true, color: COLOR.accent })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: COLOR.accent, space: 4 } },
        spacing: { after: 400 },
        children: [new TextRun('')]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: 'Análise completa: Website + Instagram', font: 'Arial', size: 26, color: COLOR.gray })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: 'Elaborado por: Silencode.com.br', font: 'Arial', size: 24, bold: true, color: COLOR.primary })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 1440 },
        children: [new TextRun({ text: 'Data: Maio / 2026', font: 'Arial', size: 22, color: COLOR.gray })]
      }),

      // Box resumo executivo na capa
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [9026],
        rows: [new TableRow({
          children: [new TableCell({
            borders: noBorders,
            shading: { fill: COLOR.primary, type: ShadingType.CLEAR },
            margins: { top: 240, bottom: 240, left: 360, right: 360 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [new TextRun({ text: 'RESUMO EXECUTIVO', font: 'Arial', size: 28, bold: true, color: COLOR.accent })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 80 },
                children: [new TextRun({ text: 'Foram identificados ', font: 'Arial', size: 22, color: COLOR.white }),
                  new TextRun({ text: '24 problemas', font: 'Arial', size: 22, bold: true, color: COLOR.accent }),
                  new TextRun({ text: ' críticos e melhorias prioritárias no ecossistema digital da NW Executive,', font: 'Arial', size: 22, color: COLOR.white })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: 'distribuídos entre SEO, performance, design, UX, conversão e presença digital.', font: 'Arial', size: 22, color: COLOR.white })]
              }),
            ]
          })]
        })]
      }),

      // Page break
      new Paragraph({ children: [new PageBreak()] }),

      // =========== SEÇÃO 1: VISÃO GERAL ===========
      heading1('1. VISÃO GERAL DA ANÁLISE'),
      dividerLine(),

      body('Este relatório apresenta uma auditoria completa do ecossistema digital da NW Executive, cobrindo o website executivenw.com e o perfil Instagram @nwexecutive. O objetivo é identificar todas as oportunidades de melhoria para que a Silencode.com.br execute um projeto de reconstrução digital de alto impacto.'),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [4513, 4513],
        rows: [
          tableHeader(['ITEM ANALISADO', 'DETALHES'], [4513, 4513]),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              shading: { fill: COLOR.light, type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [new TextRun({ text: 'Website', font: 'Arial', size: 21, bold: true })] })] }),
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: 'executivenw.com — plataforma Wix, 2 páginas', font: 'Arial', size: 21 })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              shading: { fill: COLOR.light, type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [new TextRun({ text: 'Instagram', font: 'Arial', size: 21, bold: true })] })] }),
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: '@nwexecutive — perfil de transporte executivo com reels', font: 'Arial', size: 21 })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              shading: { fill: COLOR.light, type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [new TextRun({ text: 'Segmento', font: 'Arial', size: 21, bold: true })] })] }),
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: 'Transporte Executivo Premium / Mobilidade VIP', font: 'Arial', size: 21 })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              shading: { fill: COLOR.light, type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [new TextRun({ text: 'Problemas Identificados', font: 'Arial', size: 21, bold: true })] })] }),
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: '24 (sendo 8 críticos, 7 altos, 6 médios, 3 baixos)', font: 'Arial', size: 21 })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              shading: { fill: COLOR.light, type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [new TextRun({ text: 'Executora do projeto', font: 'Arial', size: 21, bold: true })] })] }),
            new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: 'Silencode.com.br', font: 'Arial', size: 21, color: COLOR.accent, bold: true })] })] }),
          ]}),
        ]
      }),

      spacer(300),

      // =========== SEÇÃO 2: PROBLEMAS CRÍTICOS ===========
      new Paragraph({ children: [new PageBreak()] }),
      heading1('2. PROBLEMAS IDENTIFICADOS — WEBSITE'),
      dividerLine(),

      body('A seguir estão listados todos os problemas encontrados no website executivenw.com, organizados por categoria e nível de criticidade.'),
      spacer(),

      // Tabela grande de problemas
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [900, 1200, 2200, 1700, 2360, 666],
        rows: [
          // Header
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({ borders, width: { size: 900, type: WidthType.DXA }, shading: { fill: COLOR.primary, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 100, right: 100 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'NÍVEL', font: 'Arial', size: 18, bold: true, color: COLOR.white })] })] }),
              new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, shading: { fill: COLOR.primary, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: 'ÁREA', font: 'Arial', size: 18, bold: true, color: COLOR.white })] })] }),
              new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, shading: { fill: COLOR.primary, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: 'PROBLEMA', font: 'Arial', size: 18, bold: true, color: COLOR.white })] })] }),
              new TableCell({ borders, width: { size: 1700, type: WidthType.DXA }, shading: { fill: COLOR.primary, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: 'IMPACTO', font: 'Arial', size: 18, bold: true, color: COLOR.white })] })] }),
              new TableCell({ borders, width: { size: 2360, type: WidthType.DXA }, shading: { fill: COLOR.primary, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: 'RECOMENDAÇÃO', font: 'Arial', size: 18, bold: true, color: COLOR.white })] })] }),
              new TableCell({ borders, width: { size: 466, type: WidthType.DXA }, shading: { fill: COLOR.primary, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: '#', font: 'Arial', size: 18, bold: true, color: COLOR.white })] })] }),
            ]
          }),

          // Linha helper
          ...([
            ['CRÍTICO', 'SEO', 'Meta tag "noindex" ativa — o site está bloqueado para indexação no Google. Nenhuma página aparece nos buscadores.', 'Zero tráfego orgânico. Clientes que pesquisam no Google nunca encontrarão o site.', 'Remover imediatamente a tag noindex. Configurar sitemap.xml e robots.txt corretos.', '01'],
            ['CRÍTICO', 'Links', 'Botão "Facebook" aponta para o perfil WixPortugues (template padrão Wix), não para a página da empresa.', 'Visitante clica e vai para perfil desconhecido. Perda total de credibilidade.', 'Atualizar link para a página oficial da NW Executive no Facebook ou remover o ícone.', '02'],
            ['CRÍTICO', 'Links', 'Botão "TripAdvisor" aponta para tripadvisor.com genérico, sem página da NW Executive.', 'Credibilidade comprometida. Parece site abandonado ou template inacabado.', 'Criar perfil no TripAdvisor ou substituir por Instagram, WhatsApp ou LinkedIn.', '03'],
            ['CRÍTICO', 'Conversão', 'Ausência total de botão WhatsApp. Serviço de transporte VIP sem WhatsApp é crítico para conversão.', 'Clientes prontos para contratar não têm canal imediato. Perda direta de vendas.', 'Implementar botão flutuante de WhatsApp em todas as páginas com mensagem pré-definida.', '04'],
            ['CRÍTICO', 'SEO', 'Sem Open Graph Image configurada. Compartilhamentos no WhatsApp e redes sociais não geram preview visual.', 'Links compartilhados parecem spam. Redução drástica de cliques em compartilhamentos.', 'Criar OG Image (1200x630px) com branding da marca e configurar meta tags.', '05'],
            ['CRÍTICO', 'Plataforma', 'Site construído no Wix — plataforma com SEO limitado, velocidade baixa e customização restrita para premium.', 'Impossível criar experiência de marca premium. Concorre com limitações técnicas permanentes.', 'Migrar para Next.js ou WordPress com tema customizado hospedado em infra própria.', '06'],
            ['CRÍTICO', 'Rastreamento', 'Sem Google Analytics, sem Meta Pixel, sem Google Tag Manager. Zero dados sobre visitantes.', 'Impossível medir ROI de campanhas pagas, entender comportamento ou otimizar conversão.', 'Instalar GA4, Meta Pixel, GTM e configurar eventos de conversão (clique WhatsApp, formulário).', '07'],
            ['CRÍTICO', 'Instagram', 'Perfil @nwexecutive não está linkado no website. Principal canal social da marca completamente desconectado.', 'Visitante do site não consegue chegar ao Instagram onde há prova social (reels, fotos).', 'Adicionar link do Instagram no header, footer e em seção dedicada com feed embed.', '08'],

            ['ALTO', 'Design', 'Identidade visual fraca. Sem paleta de cores definida, sem tipografia premium, sem logo institucional visível no site.', 'Serviço que se propõe premium precisa ter estética premium. Site não transmite luxo.', 'Criar manual de identidade visual com paleta, tipografia, ícones e tom de voz. Implementar no site.', '09'],
            ['ALTO', 'UX/CTA', 'Ausência de CTAs (Calls to Action) claros em cada seção. O visitante não é direcionado a tomar uma ação.', 'Taxa de conversão próxima de zero. Visitante navega e sai sem entrar em contato.', 'Adicionar CTAs estratégicos: "Solicitar Cotação", "Falar no WhatsApp", "Reservar Agora" em cada bloco.', '10'],
            ['ALTO', 'Conteúdo', 'Nenhum depoimento, avaliação ou prova social no site. Sem cases, fotos reais de clientes ou parceiros.', 'Para serviços premium, prova social é essencial. Visitante não tem motivo para confiar.', 'Adicionar seção de depoimentos com nome, cargo e empresa do cliente. Integrar avaliações do Google.', '11'],
            ['ALTO', 'Navegação', 'Menu com apenas 2 itens (Início e Serviços) + "Mais" vazio. Estrutura extremamente pobre para um site premium.', 'Visitante não encontra informações sobre frota, equipe, cobertura geográfica ou FAQ.', 'Criar menu robusto: Início, Serviços, Frota, Sobre Nós, Cobertura, Blog, Contato.', '12'],
            ['ALTO', 'SEO', 'Sem meta descriptions personalizadas. Sem estrutura de heading correta (H1, H2, H3) para SEO.', 'Google não sabe do que o site trata. Ranqueamento prejudicado mesmo após corrigir o noindex.', 'Criar meta description única por página, estruturar headings semanticamente, criar sitemap.', '13'],
            ['ALTO', 'Performance', 'Imagens carregando com parâmetro blur_2 — imagens borradas são servidas no carregamento inicial.', 'Experiência visual degradada. Visitante vê site "borrado" antes do carregamento completo.', 'Usar formato WebP, lazy loading nativo, CDN de imagens e eliminar blur artificial.', '14'],
            ['ALTO', 'Formulário', 'Formulário de contato básico (nome, e-mail, mensagem) sem agendamento, sem seleção de serviço, sem cotação online.', 'Formulário genérico não qualifica o lead. Gerencia expectativas erradas.', 'Criar formulário inteligente com tipo de serviço, data/hora, origem/destino e número de passageiros.', '15'],

            ['MÉDIO', 'Conteúdo', 'Página de Serviços com textos incompletos — "Transfer aeroporto" e "Concierge" sem descrição completa.', 'Cliente não entende o que está comprando. Gera dúvidas que impedem a conversão.', 'Escrever descrições completas para cada serviço, com benefícios, diferenciais e como funciona.', '16'],
            ['MÉDIO', 'Design', 'Sem seção de frota de veículos com fotos dos carros disponíveis. Serviço de transporte sem mostrar os veículos.', 'Cliente premium quer saber exatamente em qual veículo será transportado.', 'Criar galeria/seção de frota com fotos profissionais, nome do modelo e especificações.', '17'],
            ['MÉDIO', 'SEO Local', 'Sem Google Maps embed, sem schema markup de LocalBusiness, sem endereço estruturado.', 'Não aparece em pesquisas locais no Google Maps. Perde clientes que buscam localmente.', 'Adicionar embed do Google Maps, estruturar schema LocalBusiness e criar perfil no Google Meu Negócio.', '18'],
            ['MÉDIO', 'Conteúdo', 'Sem área de cobertura geográfica clara. Site não especifica cidades, regiões ou países atendidos.', 'Cliente não sabe se o serviço atende sua localidade. Abandona sem entrar em contato.', 'Criar seção "Áreas de Atendimento" com mapa interativo e lista de cidades/aeroportos cobertos.', '19'],
            ['MÉDIO', 'Design Mobile', 'Wix frequentemente apresenta problemas de responsividade em telas menores. Não testado em múltiplos dispositivos.', 'Mais de 70% dos acessos são mobile. Experiência ruim = cliente perdido imediatamente.', 'Testar em 5+ dispositivos. Reconstruir mobile-first com breakpoints profissionais.', '20'],
            ['MÉDIO', 'Conversão', 'Sem chat online (livechat, Tawk.to, Tidio). Para serviço de agendamento, chat é essencial.', 'Clientes com dúvidas imediatas não têm canal rápido além do telefone.', 'Implementar chat online integrado ao WhatsApp Business para capturar leads em tempo real.', '21'],

            ['BAIXO', 'Conteúdo', 'Sem blog ou seção de conteúdo. Sem artigos sobre destinos, viagens de negócios ou dicas VIP.', 'Perde oportunidade de atrair tráfego orgânico via conteúdo relevante.', 'Criar blog com artigos estratégicos para SEO: "Transfer aeroporto SP", "Transporte executivo GRU" etc.', '22'],
            ['BAIXO', 'Segurança', 'Política de cookies implementada apenas via link, sem banner de consentimento LGPD adequado.', 'Possível não-conformidade com LGPD. Risco jurídico menor mas presente.', 'Implementar banner de cookies com aceite/rejeição conforme LGPD e GDPR.', '23'],
            ['BAIXO', 'Acessibilidade', 'Sem atributos alt text nas imagens, sem landmarks ARIA, sem suporte a leitores de tela.', 'Site inacessível para portadores de deficiência visual. Também afeta SEO negativamente.', 'Adicionar alt text descritivo em todas as imagens. Implementar estrutura semântica HTML5.', '24'],
          ]).map(([nivel, area, problema, impacto, rec, num]) =>
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  width: { size: 900, type: WidthType.DXA },
                  shading: { fill: nivel === 'CRÍTICO' ? COLOR.danger : nivel === 'ALTO' ? COLOR.warning : nivel === 'MÉDIO' ? COLOR.info : COLOR.success, type: ShadingType.CLEAR },
                  margins: { top: 80, bottom: 80, left: 80, right: 80 },
                  verticalAlign: VerticalAlign.CENTER,
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: nivel, font: 'Arial', size: 17, bold: true, color: COLOR.white })] })]
                }),
                new TableCell({
                  borders,
                  width: { size: 1200, type: WidthType.DXA },
                  margins: { top: 80, bottom: 80, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: area, font: 'Arial', size: 19, bold: true, color: COLOR.primary })] })]
                }),
                new TableCell({
                  borders,
                  width: { size: 2200, type: WidthType.DXA },
                  margins: { top: 80, bottom: 80, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: problema, font: 'Arial', size: 19, color: COLOR.darkgray })] })]
                }),
                new TableCell({
                  borders,
                  width: { size: 1700, type: WidthType.DXA },
                  margins: { top: 80, bottom: 80, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: impacto, font: 'Arial', size: 19, color: COLOR.darkgray })] })]
                }),
                new TableCell({
                  borders,
                  width: { size: 2360, type: WidthType.DXA },
                  margins: { top: 80, bottom: 80, left: 100, right: 100 },
                  children: [new Paragraph({ children: [new TextRun({ text: rec, font: 'Arial', size: 19, color: COLOR.darkgray })] })]
                }),
                new TableCell({
                  borders,
                  width: { size: 466, type: WidthType.DXA },
                  shading: { fill: 'F5F5F5', type: ShadingType.CLEAR },
                  margins: { top: 80, bottom: 80, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: num, font: 'Arial', size: 19, bold: true, color: COLOR.gray })] })]
                }),
              ]
            })
          )
        ]
      }),

      spacer(300),

      // =========== SEÇÃO 3: ANÁLISE INSTAGRAM ===========
      new Paragraph({ children: [new PageBreak()] }),
      heading1('3. ANÁLISE DO INSTAGRAM @nwexecutive'),
      dividerLine(),

      heading2('3.1 Pontos Positivos'),
      bullet('Perfil ativo com produção de Reels — formato de maior alcance orgânico no Instagram em 2025/26'),
      bullet('Nome do perfil "TRANSPORTE EXECUTIVO" é direto e descritivo para SEO de perfil'),
      bullet('Conteúdo visual alinhado ao nicho premium (veículos, motoristas, serviços)'),
      bullet('Presença no Instagram cria prova social que o site poderia aproveitar'),
      spacer(),

      heading2('3.2 Problemas e Oportunidades no Instagram'),
      spacer(60),

      infoBox('❌ Problema 1 — Desconexão total com o site', [
        'O site executivenw.com não exibe nenhum link, menção ou embed do Instagram. As duas plataformas operam em silos.',
        'Solução: Integrar feed do Instagram no site via widget. Criar UTM links para rastrear tráfego do bio.',
      ], 'FDECEA', COLOR.danger),
      spacer(120),

      infoBox('❌ Problema 2 — Link na Bio sem landing page otimizada', [
        'O link na bio aponta direto para o site Wix com parâmetros UTM, mas o site não tem landing page específica para tráfego do Instagram.',
        'Solução: Criar landing page dedicada para tráfego social com CTA de WhatsApp imediato e formulário de cotação.',
      ], 'FDECEA', COLOR.danger),
      spacer(120),

      infoBox('⚠️ Problema 3 — Sem estratégia de conteúdo documentada', [
        'Sem calendário editorial visível, os reels parecem publicados de forma ad hoc, sem funil de conteúdo (topo, meio, fundo).',
        'Solução: Criar calendário editorial mensal com 3 pilares: institucional, prova social e educativo.',
      ], 'FEF9E7', COLOR.warning),
      spacer(120),

      infoBox('⚠️ Problema 4 — Sem Highlights organizados', [
        'Destaque (Highlights) são o currículo visual do perfil. Sem eles organizados (Frota, Clientes, Serviços, Depoimentos), o perfil perde oportunidade de converter novos visitantes.',
        'Solução: Criar mínimo 5 highlights com capas personalizadas: Frota, VIPs, Aeroportos, Avaliações, Como Funciona.',
      ], 'FEF9E7', COLOR.warning),
      spacer(120),

      infoBox('ℹ️ Oportunidade — Reels com baixo alcance potencial sem hashtags estratégicas', [
        'Reels de nicho premium precisam de hashtags específicas para atingir tomadores de decisão corporativos e viajantes VIP.',
        'Sugestão de hashtags: #transporteexecutivo #viptransfer #motoristaprivado #executivetransfer #transferaeroporto #mobilidadepremium',
      ], 'EBF5FB', COLOR.info),
      spacer(300),

      // =========== SEÇÃO 4: ANÁLISE DE DESIGN ===========
      new Paragraph({ children: [new PageBreak()] }),
      heading1('4. ANÁLISE DE DESIGN E EXPERIÊNCIA DO USUÁRIO'),
      dividerLine(),

      heading2('4.1 Diagnóstico Visual Atual'),
      body('O site atual não reflete o posicionamento premium que a NW Executive comunica em seu texto. Há uma dissonância grave entre o que a marca promete ("sofisticação", "VIP", "mobilidade premium") e o que o site entrega visualmente.'),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3000, 3013, 3013],
        rows: [
          tableHeader(['ELEMENTO', 'SITUAÇÃO ATUAL', 'PADRÃO PREMIUM'], [3000, 3013, 3013]),
          ...([
            ['Paleta de Cores', 'Indefinida / Wix padrão', 'Preto, dourado e branco (luxo)'],
            ['Tipografia', 'Fonte genérica do Wix', 'Playfair Display + Inter (premium)'],
            ['Logo', 'Texto simples sem peso visual', 'Logotipo vetorial com brasão/símbolo'],
            ['Imagens', 'Fotos borradas, baixa resolução', 'Fotografia profissional de veículos e equipe'],
            ['Espaçamento', 'Parágrafos colados, sem ritmo', 'Espaçamento generoso, estilo editorial'],
            ['Animações', 'Nenhuma / Wix básico', 'Transições suaves, scroll parallax'],
            ['Hero Section', 'Imagem de fundo sem impacto', 'Vídeo ou foto full-screen de carro em movimento'],
            ['Botões', 'Estilo padrão do Wix', 'Botões dourado com borda, hover elegante'],
          ]).map(([el, atual, premium]) =>
            new TableRow({ children: [
              new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, shading: { fill: COLOR.light, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: el, font: 'Arial', size: 21, bold: true })] })] }),
              new TableCell({ borders, width: { size: 3013, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: atual, font: 'Arial', size: 21, color: COLOR.danger })] })] }),
              new TableCell({ borders, width: { size: 3013, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: premium, font: 'Arial', size: 21, color: COLOR.success })] })] }),
            ]})
          )
        ]
      }),

      spacer(300),

      heading2('4.2 Estrutura de Páginas Recomendada'),
      body('A Silencode deve propor um site com no mínimo 8 páginas estratégicas:'),
      spacer(60),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 1500, 5526],
        rows: [
          tableHeader(['PÁGINA', 'PRIORIDADE', 'OBJETIVO ESTRATÉGICO'], [2000, 1500, 5526]),
          ...([
            ['Home', 'CRÍTICA', 'Hero com vídeo, proposta de valor, serviços em destaque, prova social, CTA WhatsApp'],
            ['Serviços', 'CRÍTICA', 'Cada serviço em página própria com descrição, benefícios, fotos e formulário'],
            ['Frota', 'ALTA', 'Galeria dos veículos com specs técnicas, capacidade e fotos profissionais'],
            ['Sobre Nós', 'ALTA', 'História da empresa, equipe, diferenciais, certificações e valores'],
            ['Cobertura', 'ALTA', 'Mapa interativo com cidades, aeroportos e rotas atendidas'],
            ['Depoimentos', 'MÉDIA', 'Avaliações de clientes, logos de empresas parceiras, cases de sucesso'],
            ['Blog', 'MÉDIA', 'Conteúdo SEO para atrair tráfego orgânico de long-tail keywords'],
            ['Contato', 'CRÍTICA', 'Formulário inteligente, WhatsApp, telefone, Google Maps e chat online'],
          ]).map(([pag, pri, obj]) => {
            const priColor = pri === 'CRÍTICA' ? COLOR.danger : pri === 'ALTA' ? COLOR.warning : COLOR.info;
            return new TableRow({ children: [
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: COLOR.light, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: pag, font: 'Arial', size: 21, bold: true })] })] }),
              new TableCell({ borders, width: { size: 1500, type: WidthType.DXA }, shading: { fill: pri === 'CRÍTICA' ? 'FADBD8' : pri === 'ALTA' ? 'FDECEA' : 'D6EAF8', type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: pri, font: 'Arial', size: 20, bold: true, color: priColor })] })] }),
              new TableCell({ borders, width: { size: 5526, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: obj, font: 'Arial', size: 21, color: COLOR.darkgray })] })] }),
            ]});
          })
        ]
      }),

      spacer(300),

      // =========== SEÇÃO 5: PLANO DE AÇÃO ===========
      new Paragraph({ children: [new PageBreak()] }),
      heading1('5. PLANO DE IMPLEMENTAÇÃO — SILENCODE.COM.BR'),
      dividerLine(),

      heading2('FASE 1 — Quick Wins (Semanas 1-2)'),
      body('Ações imediatas de alto impacto que podem ser executadas rapidamente:'),
      spacer(60),

      bullet('Remover meta tag noindex de todas as páginas do Wix atual'),
      bullet('Corrigir links quebrados: Facebook (→ página da NW) e TripAdvisor (→ remover ou substituir)'),
      bullet('Adicionar botão flutuante de WhatsApp com mensagem pré-configurada'),
      bullet('Instalar Google Analytics 4 e Meta Pixel no site atual'),
      bullet('Linkar Instagram @nwexecutive no menu e footer'),
      bullet('Criar perfil no Google Meu Negócio e solicitar avaliações de clientes'),
      spacer(120),

      heading2('FASE 2 — Reconstrução (Semanas 3-8)'),
      body('Desenvolvimento do novo site na plataforma escolhida:'),
      spacer(60),

      bullet('Definição de identidade visual (paleta, tipografia, logo, iconografia)'),
      bullet('Criação de wireframes para todas as 8 páginas'),
      bullet('Desenvolvimento do site em Next.js ou WordPress premium'),
      bullet('Sessão fotográfica profissional de frota e equipe'),
      bullet('Redação de copy estratégico para cada seção e página'),
      bullet('Implementação de formulário inteligente de cotação'),
      bullet('Configuração completa de SEO técnico + on-page'),
      bullet('Integração com WhatsApp Business API'),
      bullet('Embed de feed do Instagram na Home'),
      bullet('Configuração de chat online integrado ao WhatsApp'),
      spacer(120),

      heading2('FASE 3 — Otimização (Semanas 9-12)'),
      body('Refinamento e crescimento orgânico:'),
      spacer(60),

      bullet('Publicação das primeiras 4 páginas do blog com keywords estratégicas'),
      bullet('Configuração de campanhas Google Ads para keywords de transporte executivo'),
      bullet('Auditoria completa de performance (Core Web Vitals, PageSpeed)'),
      bullet('Implementação de schema markup LocalBusiness e Service'),
      bullet('Criação de calendário editorial para Instagram'),
      bullet('Configuração de remarketing no Meta Ads'),
      bullet('Relatório mensal de KPIs e evolução do ranqueamento'),
      spacer(300),

      // =========== SEÇÃO 6: STACK TECNOLÓGICO ===========
      new Paragraph({ children: [new PageBreak()] }),
      heading1('6. STACK TECNOLÓGICO RECOMENDADO'),
      dividerLine(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 2200, 4626],
        rows: [
          tableHeader(['CATEGORIA', 'TECNOLOGIA', 'JUSTIFICATIVA'], [2200, 2200, 4626]),
          ...([
            ['Frontend', 'Next.js 14 + TailwindCSS', 'Performance máxima, SSR/SSG, excelente para SEO, fácil de customizar com design premium'],
            ['CMS', 'Sanity.io ou Contentful', 'Permite cliente editar conteúdo sem mexer no código. Blog fácil de gerenciar.'],
            ['Hospedagem', 'Vercel ou AWS', 'CDN global, zero-config deploy, SSL automático, tempo de resposta < 100ms'],
            ['Formulário', 'Formspree + Zapier', 'Formulário com ação automática: lead → WhatsApp + CRM + e-mail'],
            ['Analytics', 'Google Analytics 4 + Meta Pixel', 'Rastreamento completo de conversões e comportamento do usuário'],
            ['SEO', 'Next SEO + Schema.org', 'Meta tags dinâmicas, sitemap automático, structured data para rich results'],
            ['WhatsApp', 'WhatsApp Business API', 'Botão flutuante + mensagem pré-configurada por tipo de serviço'],
            ['Chat', 'Tidio ou JivoChat', 'Chat online integrado ao WhatsApp para captura de leads em tempo real'],
            ['Imagens', 'Cloudinary + WebP', 'Otimização automática de imagens, lazy loading, formatos modernos'],
            ['Mapas', 'Google Maps Embed + Places API', 'Mapa interativo de cobertura geográfica e localização'],
          ]).map(([cat, tech, just]) =>
            new TableRow({ children: [
              new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, shading: { fill: COLOR.light, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: cat, font: 'Arial', size: 21, bold: true })] })] }),
              new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: tech, font: 'Arial', size: 21, color: COLOR.accent, bold: true })] })] }),
              new TableCell({ borders, width: { size: 4626, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: just, font: 'Arial', size: 21, color: COLOR.darkgray })] })] }),
            ]})
          )
        ]
      }),

      spacer(300),

      // =========== SEÇÃO 7: KPIs ===========
      new Paragraph({ children: [new PageBreak()] }),
      heading1('7. KPIs E MÉTRICAS DE SUCESSO'),
      dividerLine(),

      body('As seguintes métricas devem ser acompanhadas mensalmente após a implementação:'),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 2000, 2000, 2526],
        rows: [
          tableHeader(['MÉTRICA', 'VALOR ATUAL', 'META 3 MESES', 'META 6 MESES'], [2500, 2000, 2000, 2526]),
          ...([
            ['Tráfego orgânico (Google)', '~0 (noindex)', '+200 visitas/mês', '+800 visitas/mês'],
            ['Posição Google "transporte executivo SP"', 'Não indexado', 'Top 20', 'Top 10'],
            ['Taxa de conversão (contato)', '< 0.5% (est.)', '> 3%', '> 5%'],
            ['Cliques WhatsApp por semana', 'Desconhecido', '20+ cliques', '50+ cliques'],
            ['PageSpeed Score (Mobile)', 'Estimado: 40-55', '> 75', '> 90'],
            ['Seguidores Instagram / mês', 'Crescimento orgânico', '+150 seguidores', '+500 seguidores'],
            ['Leads qualificados / mês', 'Não rastreado', '15+ leads', '40+ leads'],
            ['Avaliações Google Meu Negócio', '0 (não configurado)', '10+ avaliações', '30+ avaliações'],
          ]).map(([metrica, atual, m3, m6]) =>
            new TableRow({ children: [
              new TableCell({ borders, width: { size: 2500, type: WidthType.DXA }, shading: { fill: COLOR.light, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: metrica, font: 'Arial', size: 21, bold: true })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: atual, font: 'Arial', size: 21, color: COLOR.danger })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: m3, font: 'Arial', size: 21, color: COLOR.warning })] })] }),
              new TableCell({ borders, width: { size: 2526, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: m6, font: 'Arial', size: 21, color: COLOR.success })] })] }),
            ]})
          )
        ]
      }),

      spacer(300),

      // =========== RODAPÉ FINAL ===========
      new Paragraph({ children: [new PageBreak()] }),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [9026],
        rows: [new TableRow({
          children: [new TableCell({
            borders: noBorders,
            shading: { fill: COLOR.primary, type: ShadingType.CLEAR },
            margins: { top: 360, bottom: 360, left: 480, right: 480 },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: 'SILENCODE.COM.BR', font: 'Arial', size: 36, bold: true, color: COLOR.accent })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: 'Desenvolvimento de Sites e Estratégias Digitais', font: 'Arial', size: 24, color: COLOR.white })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', font: 'Arial', size: 22, color: COLOR.accent })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Este relatório é confidencial e destinado exclusivamente à NW Executive.', font: 'Arial', size: 20, color: '999999' })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '© 2026 Silencode.com.br — Todos os direitos reservados.', font: 'Arial', size: 20, color: '999999' })] }),
            ]
          })]
        })]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Auditoria_NW_Executive_Silencode.docx', buffer);
  console.log('Relatório gerado com sucesso!');
});