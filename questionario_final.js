
function onFormSubmit(e) {
  if (!e || !e.range) {
    throw new Error(
      "Esta função deve ser executada pelo acionador de envio do formulário."
    );
  }

  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  // Lê todas as respostas da linha enviada.
  const responses = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  // CONFIGURAÇÃO DA PLANILHA
  // A = carimbo de data/hora
  // B = e-mail
  // C:O = 13 perguntas
  // V = pontuação; W = nível; X = percentual
  const email = String(responses[1] || "").trim();

  if (!email) {
    console.log("Nenhum e-mail encontrado na coluna B.");
    return;
  }

  // Cada pergunta vale:
  // correta = 2 pontos
  // parcial = 1 ponto
  // errada = 0 pontos
  const perguntas = [
    {
      tema: "Senha do Wi-Fi",
      dica: "Utilize uma senha longa, exclusiva e difícil de adivinhar.",
      correta: [
        "É longa (> 10 dígitos) e usa vários tipos de carácter."
      ],
      parcial: [
        "É curta e tem diferentes tipos de carácter (ex: letras minúsculas com letras maiúsculas e números)"
      ],
      errada: [
        "Não tem senha. Deixo minha rede aberta para todos que quiserem usar.",
        "Minha senha é uma curta (< 10 dígitos) sequência simples de apenas um tipo de caractere, fácil de lembrar."
      ]
    },
    {
      tema: "Nome da rede Wi-Fi",
      dica: "Evite nomes de rede que revelem seu nome, endereço ou informações pessoais.",
      correta: [
        "Sim, o nome que escolhi não cabe nas categorias descritas acima"
      ],
      parcial: [
        "Não, deixei o nome padrão.",
        "Sim, o nome inclui dados sobre o meu equipamento."
      ],
      errada: [
        "Sim, inclui informações pessoais como meu nome e número da casa."
      ]
    },
    {
      tema: "Proteção da rede Wi-Fi",
      dica: "Utilize criptografia moderna, como WPA2-AES ou WPA3, e mantenha o firewall ativo quando aplicável.",
      correta: [
        "Ativei a criptografia mais atual na minha rede.",
        "Tenho as duas primeiras opções ativas na minha rede."
      ],
      parcial: [
        "Ativei o firewall da rede."
      ],
      errada: [
        "Não conheço essas medidas de segurança."
      ]
    },
    {
      tema: "Senha do Wi-Fi descoberta",
      dica: "Troque a senha imediatamente, revise os dispositivos conectados e investigue possíveis acessos indevidos.",
      correta: [
        "Mudar a senha da minha rede e procurar onde ocorreu essa brecha na segurança."
      ],
      parcial: [
        "Apenas mudar a sua senha é o suficiente."
      ],
      errada: [
        "Não há necessidade de tomar nenhuma ação.",
        "Não há necessidade de tomar nenhuma ação imediata, apenas se houverem problemas relacionados a rede"
      ]
    },
    {
      tema: "Uso de software pago",
      dica: "Prefira alternativas gratuitas legítimas, software de código aberto ou licenças oficiais mais acessíveis.",
      correta: [
        "Procuro uma versão com licença mais barata",
        "Procuro software alternativos mais baratos ou de graça"
      ],
      parcial: [
        "Espero a versão do software que preciso estar disponível"
      ],
      errada: [
        "Procuro e baixo a primeira versão não oficial do software que encontro na internet."
      ]
    },
    {
      tema: "Atualização de software",
      dica: "Mantenha o sistema operacional e os aplicativos atualizados, priorizando correções de segurança.",
      correta: [
        "Tenho atualizações automáticas ativadas em todos meus dispositivos.",
        "Faço questão de verificar se há atualizações disponíveis para softwares importantes, como meu sistema operacional."
      ],
      parcial: [
        "Atualizo apenas quando sou obrigado."
      ],
      errada: [
        "Tento evitar atualizar meu sistema operacional."
      ]
    },
    {
      tema: "Backup de dados",
      dica: "Faça cópias de segurança regularmente, mantenha mais de uma cópia e teste a restauração dos arquivos.",
      correta: [
        "Faço back-up em um dispositivo físico (pen-drive, SSD ou HDD) e na nuvem."
      ],
      parcial: [
        "Fiz um back-up quando primeiro tive problemas com perda de dados e continuo fazendo.",
        "Fiz quando um técnico de TI me pede para fazer um back-up."
      ],
      errada: [
        "Desconheço o termo, talvez nunca fiz \"back-up\" dos meus dados."
      ]
    },
    {
      tema: "Proteção dos dados no computador",
      dica: "Combine criptografia de disco, firewall, proteção contra malware, atualizações e boas práticas de acesso.",
      correta: [
        "Adoto as duas primeiras opções."
      ],
      parcial: [
        "Mantenho meu disco rígido cifrado.",
        "Acredito que as funções de segurança nativas do meu sistema operacional , como antivirus e firewall, são o suficiente."
      ],
      errada: [
        "Não sou responsável pela proteção dos meus dados."
      ]
    },
    {
      tema: "Ransomware",
      dica: "Mantenha backups separados, atualize seus sistemas e desconfie de anexos e links inesperados.",
      correta: [
        "Ransomware"
      ],
      parcial: [],
      errada: [
        "Adware",
        "Rootkit",
        "Worm"
      ]
    },
    {
      tema: "Keylogger",
      dica: "Utilize proteção contra malware, mantenha os dispositivos atualizados e ative autenticação multifator.",
      correta: [
        "Keylogger"
      ],
      parcial: [],
      errada: [
        "Stalkerware",
        "Screenlogger",
        "Adware"
      ]
    },
    {
      tema: "Instalação de aplicativos",
      dica: "Baixe aplicativos de lojas oficiais, confira o desenvolvedor e mantenha as atualizações ativadas.",
      correta: [
        "Baixar aplicativos somente de lojas oficiais"
      ],
      parcial: [
        "Salvar as senhas de acesso no navegador"
      ],
      errada: [
        "Desativar as atualizações automáticas",
        "Instalar aplicativos recebidos por links em mensagens"
      ]
    },
    {
      tema: "Botnet",
      dica: "Mantenha seus dispositivos atualizados, proteja suas credenciais e monitore comportamentos suspeitos.",
      correta: [
        "Rede formada por inúmeros dispositivos zumbis controlados remotamente"
      ],
      parcial: [],
      errada: [
        "Um único vírus que se autorreplica em arquivos",
        "Um firewall pessoal instalado no dispositivo",
        "Um mecanismo automático de backup de dados"
      ]
    },
    {
      tema: "Vazamento de dados",
      dica: "Proteja seus equipamentos contra furto, bloqueie o acesso e mantenha os dados sensíveis criptografados.",
      correta: [
        "Furto de equipamentos"
      ],
      parcial: [],
      errada: [
        "Uso de firewall pessoal",
        "Verificação em duas etapas",
        "Backup automático"
      ]
    }
  ];

  // Corrige as 13 perguntas.
  let pontuacao = 0;
  const pontosFracos = [];

  perguntas.forEach((pergunta, index) => {
    // A primeira pergunta está na coluna C (índice 2).
    const resposta = String(responses[index + 2] || "").trim();

    let pontos = 0;

    if (pergunta.correta.includes(resposta)) {
      pontos = 2;
    } else if (pergunta.parcial.includes(resposta)) {
      pontos = 1;
    } else {
      pontos = 0;
    }

    pontuacao += pontos;

    // Recomendações para respostas parciais ou erradas.
    if (pontos < 2) {
      pontosFracos.push(pergunta);
    }
  });

  const pontuacaoMax = perguntas.length * 2;
  const percentual = Math.round(
    (pontuacao / pontuacaoMax) * 100
  );

  // Classificação: máximo de 26 pontos.
  let nivel;
  let dicaGeral;
  let emoji;

  if (pontuacao >= 22) {
    nivel = "ALTO";
    emoji = "🏆";
    dicaGeral = "Você demonstra bons hábitos de cibersegurança. Continue se atualizando e revise suas práticas periodicamente.";
  } else if (pontuacao >= 14) {
    nivel = "MÉDIO";
    emoji = "📈";
    dicaGeral = "Você já conhece algumas práticas importantes, mas ainda existem pontos que merecem atenção.";
  } else {
    nivel = "BAIXO";
    emoji = "🚨";
    dicaGeral = "Seu resultado indica que é importante reforçar suas práticas de segurança digital. Comece pelas recomendações abaixo.";
  }

  // Registra os resultados nas colunas V, W e X.
  sheet.getRange(row, 22, 1, 3).setValues([[
    `${pontuacao}/${pontuacaoMax}`,
    nivel,
    `${percentual}%`
  ]]);

  // Monta a lista de recomendações personalizadas.
  const dicasHtml = pontosFracos.length
    ? `<h3>🎯 Pontos para melhorar (${pontosFracos.length})</h3>
       <ul>${pontosFracos.map(p => 
         `<li><b>${p.tema}:</b> ${p.dica}</li>`
       ).join("")}</ul>`
    : "<h3>✅ Excelente! Todas as respostas receberam pontuação máxima.</h3>";

  const assunto = "Projeto Integrador - Seu resultado de Cibersegurança";

  const corNivel =
    nivel === "ALTO" ? "green" :
    nivel === "MÉDIO" ? "orange" : "red";

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;color:#222">
      <h1>${emoji} Seu resultado de cibersegurança</h1>

      <div style="background:#f0f8ff;padding:20px;border-radius:10px">
        <h2>Pontuação: ${pontuacao}/${pontuacaoMax}</h2>
        <h3>Percentual: ${percentual}%</h3>
        <h3>Nível:
          <span style="color:${corNivel}">${nivel}</span>
        </h3>
        <p>${dicaGeral}</p>
      </div>

      ${dicasHtml}

      <hr>

      <p><b>Próximo passo:</b> escolha uma das recomendações acima e aplique-a aos seus dispositivos.</p>

      <p style="color:#666;font-size:12px">
        Projeto Cibersegurança Interativa | Engenharia Mackenzie
      </p>
    </div>
  `;

  GmailApp.sendEmail(email, assunto, dicaGeral, {
    htmlBody: htmlBody,
    name: "Projeto Cibersegurança"
  });

  console.log(
    `${email}: ${pontuacao}/${pontuacaoMax} (${percentual}%) - ${nivel}`
  );
}
