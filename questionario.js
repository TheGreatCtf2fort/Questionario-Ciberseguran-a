function onFormSubmit(e) {
  if (!e || !e.source) {
    console.log('Execute via formulário, não manual!');
    return;
  }

  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const responses = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const emailCol = 2; // Coluna B: E-mail
  const email = responses[emailCol - 1]?.toString().trim();
  if (!email) return;

  // === CONFIGURE AQUI SUAS 10 PERGUNTAS ===
  const respostasCorretas = [  // Coluna C=0, D=1, ..., L=9
    { // Pergunta 1 (coluna C)
      id: 1,
      tema: 'Certificados de segurança',
      dica: 'Os certificados de segurança são essenciais em todos os ambientes',
      correta: ['Devem ser observados em sites de qualquer natureza'],
      parcial: ['Servem para métodos de assinatura eletrônica', 'em algumas contas'],
      errada: ['Ainda que esteja expirado, não há risco ao usuário', 'Não conheco o termo']
    },
    { // Pergunta 2 (coluna D)
      id: 2,
      tema: 'Senhas fortes',
      dica: 'Utilize senhas longas e com caracteres especiais',
      correta: ['Uma combinação pequena de letras, números e caracteres especiais'],
      parcial: ['Uma senha com mais de 8 caracteres'],
      errada: ['Uma senha fácil de lembrar','Não sei']
    },
    { // Pergunta 3
      id: 3,
      tema: 'Troca de senha',
      dica: 'Troque de senhas regularmente',
      correta: ['A cada 3 meses','Regularmente (mensalmente ou mais frequente)'],
      parcial: ['Somente quando solicitado'],
      errada: ['Nunca']
    },
    {
      id: 4,
      tema: 'Gerenciador de Senhas',
      dica: 'Se for possível utilizar, vale a pena!',
      correta: ['É importante para concentrar senhas, mas aumenta os risos'],
      parcial: ['Não deve ser usado, pois não é recomendado','É um meio efetivo, por si só, para proteção de senhas'],
      errada: ['Não conheço'] },
    {
      id: 5,
      tema: 'Segundo Fator de Autenticação',
      dica: 'Sempre use por aplicativo',
      correta: ['É um elemento que incrementa a segurança de acessos','Pode ser utilizado em conjunto com outros métodos de segurança','Dependendo do método escolhido, pode oferecer riscos'],
      parcial: ['Não conheço'],
      errada: [''] },
    {
      id: 6,
      tema: 'SPAM',
      dica: 'Fique sempre atento as mensagens recebidas',
      correta: ['É caracterizado por e-mails com conteúdo falso'],
      parcial: ['Está associado a mensagens que mascaram o originador'],
      errada: ['É perigoso quando está atrelado a algum link','Não conheço'] },

    {
      id: 7,
      tema: 'Atualização de sistemas',
      dica: 'Faça a atualização regularmente, quando solicitado',
      correta: ['Há necessidade de atualização semanal','É necessário atualizar, assim que solicitado'],
      parcial: ['Apenas atualizaçõe de grande porte são obrigatórias'],
      errada: ['compartilhÉ necessário atualizar quando torna-se obrigatórioo'] },


    {
      id: 8,
      tema: 'Antivírus',
      dica: 'Adquira um ou utilize, pelo menos, um gratuito',
      correta: ['A contratação de um bom aplicativo é uma medida infalível contra vírus'],
      parcial: ['O próprio Windows Defender já é suficiente','Qualquer antivírus disponível já é suficiente para uma proteção básica'],
      errada: ['Não há necesside de instalação de antivírus em estações pessoais'] },

    {
      id: 9,
      tema: 'Formulários on-line',
      dica: 'De preferência, nunca preencha!',
      correta: [''],
      parcial: ['O preenchimento só deve ser feito caso a LGPD esteja respeitada','Nunca se deve realizar preenchimentos em formulários on-line'],
      errada: ['Quando demandado, é necessário preenchê-losnunca leio','Sempre devem ser preenchidos formulários on-line'] },

    {
      id: 10,
      tema: 'Sites Seguros',
      dica: 'Sempre verifique os selos de segurança',
      correta: ['A avaliação é complexa e demanda conhecimento técnico'],
      parcial: ['São fáceis de identificar','Se há selos de qualidade, não há risco'],
      errada: ['Se são recomendados, o uso é livre'] },

    {
      id: 11,
      tema: 'Wi-Fi Públicos',
      dica: 'Nunca utilize!',
      correta: ['Não devem ser utilizados em hipótese alguma'],
      parcial: [''],
      errada: ['Podem ser utilizados, desde que requeiram cadastro','Podem ser utilizados, desde que não requeiram cadastro','Não há restrição ao uso'] },

    {
      id: 12,
      tema: 'Backup',
      dica: 'Quem tem 1, não tem nenhum. Quem tem 2, talvez tenha 1!',

      correta: ['Redundâncias são sempre necessárias'],
      parcial: ['Somente a nuvem é segura','Somente mídia física é segura'],
      errada: ['Não há necessidade de bakups'] },

    {
      id: 13,
      tema: 'LGPD',
      dica: 'Vale a pena estudar! São seus direitos.',

      correta: ['Conheço o conteúdo e seus objetivos'],
      parcial: ['Desconheço seu conteúdo, mas sei sobre o que se trata','É utilizada exclusivamente para proteção de dados de empresas'],
      errada: ['Não tenho informações'] },

    {
      id: 14,
      tema: 'Repetição de senha',
      dica: 'É fazer e ser kackeado!',
      correta: [''],
      parcial: ['Não deve ocorrer, embora o risco de exposição seja baixo','Pode ser usada, desde que ela seja grande','É comum, mas deve ser evitada'],
      errada: ['Pode ser usada, desde que um carctere seja trocado'] },

    {
      id: 15,
      tema: 'Sobre criptografia',
      dica: 'Sem criptografia, seus dados estão expostos!',

      correta: ['Deve ser usada em qualquer situação de envio de dados','É imprescindível quando há senhas envolvidas nos dados'],
      parcial: ['É imprescindível apenas em transações relacionadas a cartões de crédito','É importante, mas pode afetar o processamento de dados'],
      errada: [''] },


    {
      id: 16,
      tema: 'Caiu em golpe? Não se preocupe',
      dica: 'Revise sua proteção digital!',

      correta: ['Não, nunca fui e não conheço',],
      parcial: ['Sim, eu conheço alguém'],
      errada: ['Sim, eu já fui','Não conheço o termo "golpe cibernético"'] },
    {
      id: 17,
      tema: 'Engenharia Social',
      dica: 'Desconfie de tudo!',
      correta: ['Está relacionado com a obtenção de dados privados de terceiros',],
      parcial: ['Guarda relação com a exposição de dados públicos, por exemplo','É uma técnica hacker de difícil detecção'],
      errada: ['Não conheço o termo'] },

    {
      id: 18,
      tema: 'Man-In-The-Middle',
      dica: 'Pode haver alguém interceptando seus dados',
      correta: ['É uma ação hacker que pode afetar empresas',],
      parcial: ['É uma técnica de ataque cibernético direcionada a pessoas','Não necessariamente se trata de uma ação hacker'],
      errada: ['Desconheço o termo'] },
  ];

  // Calcula pontuação EXATA
  let pontuacao = 0, erros=[];
  respostasCorretas.forEach((ref, index) => {
    if (index >= responses.length - 2) return; // Para se faltarem colunas

    const resposta = responses[index + 3]; // +2 = C=2

    if (ref.correta.some(c => c === resposta)) pontuacao += 2;
    else if (ref.parcial.some(p => p === resposta)) pontuacao += 1;
    else erros.push(ref)
  });

  const pontuacaoMax = 36;
  const percentual = Math.round((pontuacao / pontuacaoMax) * 100);

  // Nível
  let nivel, emoji, dica;
  if (pontuacao >= 30) { nivel = 'ALTO'; emoji = '&#127942;'; dica = 'Parabéns! Nível Alto.'; }
  else if (pontuacao >= 12) { nivel = 'MÉDIO'; emoji = '📈'; dica = 'Aceitável, mas reforçe hábitos.'; }
  else { nivel = 'BAIXO'; emoji = '🚨'; dica = 'ALERTA: Aja imediatamente nas dicas!'; }

  // Grava planilha
  sheet.getRange(row, 22).setValue(pontuacao + '/' + pontuacaoMax);
  sheet.getRange(row, 23).setValue(nivel);
  sheet.getRange(row, 24).setValue(percentual + '%');


// Dicas PERSONALIZADAS (só erros) - EMOJIS FIXOS
let dicasHtml = erros.length ?
  `<h3>&#127919; Foco nestes ${erros.length} pontos:</h3>
   ${erros.map(e => `<li><b>${e.tema}:</b> ${e.dica}</li>`).join('')}` :
  '<h3>&#9989; Perfeito! Sem pontos fracos detectados.</h3>';

// Assunto com emoji seguro
const assunto = `Projeto Integrador - Veja seu Nível de Cibersegurança!`;

// E-mail ÚNICO por nível - TODOS EMOJIS CORRIGIDOS
const htmlBody = `
<div style="font-family:Arial,sans-serif;max-width:600px">
  <h1>${nivel === 'ALTO' ? '&#127942;' : nivel === 'MÉDIO' ? '&#128200;' : '&#128680;'} </h1>
  <div style="background:#f0f8ff;padding:20px;border-radius:10px">
    <h2>&#128200; Pontuação: ${pontuacao}/${pontuacaoMax} (${percentual}%)</h2>
    <h3>Nível: <span style="color:${nivel==='ALTO'?'green':nivel==='MÉDIO'?'orange':'red'}">${nivel}</span></h3>
    <p style="font-style:italic"> </p>
  </div>
  
  ${dicasHtml}
  
  <div style="background:#e8f5e8;padding:15px;margin:20px 0;border-left:5px solid green">
    <h4>&#128161; Dica Extra para nível ${nivel}:</h4>
    <ul>
      ${nivel==='ALTO' ? '<li>&#9989; Monitore logs de acesso (Google Account > Segurança).</li>' :
        nivel==='MÉDIO' ? '<li>&#128200; Atualize seu sistema operacional e antivírus.</li>' :
        '<li>&#10060; Primeiro passo HOJE: Mude a senha do seu e-mail principal.</li>'}
    </ul>
  </div>
  
  <hr>
  <p><small>Projeto Cibersegurança Interativa | &#128153; Engenharia Mackenzie </small></p>
</div>
`;

  GmailApp.sendEmail(email, assunto, '', { htmlBody });
  console.log(`✅ ${email}: ${pontuacao}/${pontuacaoMax} (${nivel})`);
}