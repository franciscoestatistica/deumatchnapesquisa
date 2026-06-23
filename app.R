library(shiny)
library(DT)
library(tm)
library(openxlsx)
library(shinycustomloader)
library(markdown)
library(rsconnect)

# Create directory for HTML files if it doesn't exist
dir.create("www", showWarnings = FALSE)

# Create HTML files directly instead of markdown
pesquisador_html <- '
<div style="text-align: center;">
  <img src="https://i.ibb.co/4N9y2mX/logo.png" width="200px" alt="DeuMatchNaPesquisa"/>
  <br/><br/>
  <img src="https://www.gov.br/ebserh/pt-br/hospitais-universitarios/regiao-sudeste/hc-ufu/logos/hc-ufu-assinatura.png" width="400px" alt="HC-UFU/EBSERH"/>
</div>

<div class="card" style="border-radius: 10px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); padding: 25px; margin: 30px 0; border-top: 4px solid #e83e8c;">
  <h3 style="color: #e83e8c; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px; margin-bottom: 20px;">👩‍🔬👨‍🔬 Para Pesquisadores</h3>
  
  <div class="alert alert-info" style="border-radius: 8px; padding: 20px; background-color: #e8f4f8; border-color: #17a2b8;">
    <h4 style="color: #17a2b8;"><i class="fas fa-info-circle"></i> Cadastramento Temporariamente Indisponível</h4>
    <p style="margin-top: 15px;">Estamos atualizando nossa base de dados! Em breve, abriremos novamente para recadastramento de pesquisadores atuais e cadastramento de novos interessados.</p>
  </div>
  
  <p style="font-size: 1.1em; margin: 20px 0;">
    Nosso objetivo é conectar <strong>estudantes</strong> e <strong>mentores/orientadores</strong>. Se você é um profissional com interesse em mentorar ou orientar projetos, fique atento às nossas atualizações!
  </p>
  
  <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; margin: 25px 0;">
    <h4 style="color: #17a2b8; margin-bottom: 15px;"><i class="fas fa-envelope-open-text"></i> Entre em Contato</h4>
    <p>Enquanto estamos em processo de atualização, você pode manifestar seu interesse entrando em contato diretamente com a UGITS:</p>
    
    <div style="display: flex; margin: 20px 0; align-items: center;">
      <div style="min-width: 50px; text-align: center;">
        <i class="fa fa-whatsapp" style="font-size: 24px; color: #25D366;"></i>
      </div>
      <div>
        <strong>WhatsApp:</strong><br>
        <a href="https://wa.me/553432182323?text=Estou%20interessado%20em%20ser%20pesquisador%20no%20DEUMATCHNAPESQUISA" target="_blank" style="color: #17a2b8; text-decoration: none;">(34) 3218-2323</a>
      </div>
    </div>
    
<div style="display: flex; align-items: center;">
  <div style="min-width: 50px; text-align: center;">
    <i class="fa fa-envelope" style="font-size: 24px; color: #e83e8c;"></i>
  </div>
  <div>
    <strong>Email:</strong><br>
    <a href="mailto:ugits.hc-ufu@ebserh.gov.br?subject=Interesse%20em%20cadastro%20como%20pesquisador%20-%20DeuMatchNaPesquisa" style="color: #17a2b8; text-decoration: none;">ugits.hc-ufu@ebserh.gov.br</a>
  </div>
</div>
  </div>
  
  <p style="text-align: center; font-style: italic; color: #6c757d; margin-top: 20px;">
    A <strong style="color: #e83e8c;">Unidade de Gestão da Inovação Tecnológica em Saúde</strong> - UGITS do HC-UFU/EBSERH receberá seu contato e responderá em breve! 🔍
  </p>
</div>
'

sobre_html <- '
<div style="text-align: center;">
  <img src="https://i.ibb.co/4N9y2mX/logo.png" width="200px" alt="DeuMatchNaPesquisa"/><br /><br />
  <img src="https://www.gov.br/ebserh/pt-br/hospitais-universitarios/regiao-sudeste/hc-ufu/logos/hc-ufu-assinatura.png" width="400px" alt="HC-UFU/EBSERH" /><br><br>
</div>

<div class="card" style="border-radius: 10px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 30px; border-top: 4px solid #17a2b8;">
  <h2 style="color: #17a2b8; text-align: center; margin-bottom: 30px;">🚀 Sobre o DeuMatchNaPesquisa</h2>
  
  <div style="background-color: #f0f9fa; border-radius: 10px; padding: 20px; margin-bottom: 30px; border-left: 5px solid #17a2b8;">
    <p style="font-size: 1.1em; font-style: italic; color: #17a2b8;">
      "Conectando ideias promissoras a pesquisadores experientes através da tecnologia"
    </p>
  </div>
  
<div class="row" style="display: flex; flex-wrap: wrap; margin: 0 -15px;">
  <div class="col" style="flex: 0 0 100%; max-width: 100%; padding: 0 15px; margin-bottom: 30px;">
    <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; height: 100%; border-bottom: 3px solid #17a2b8; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
      <h4 style="color: #17a2b8; margin-bottom: 15px; font-weight: 600;"><i class="fas fa-lightbulb"></i> Origem</h4>
      <p style="font-size: 1.05em; line-height: 1.6;">O projeto DeuMatchNaPesquisa.com surgiu a partir da observação de que muitos estudantes tinham ideias de pesquisa, porém não possuíam expertise para desenvolvê-las. Esta plataforma nasceu para resolver esse desafio, criando um ambiente de conexão entre ideias e mentores!</p>
    </div>
  </div>
  
  <div class="col" style="flex: 0 0 100%; max-width: 100%; padding: 0 15px; margin-bottom: 30px;">
    <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; height: 100%; border-bottom: 3px solid #e83e8c; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
      <h4 style="color: #e83e8c; margin-bottom: 15px; font-weight: 600;"><i class="fas fa-flask"></i> Missão Institucional</h4>
      <p style="font-size: 1.05em; line-height: 1.6;">Nossa missão é <strong>fomentar a pesquisa e inovação</strong> no HC-UFU/EBSERH, elevando tanto a quantidade quanto a qualidade dos projetos desenvolvidos na instituição. Buscamos colaborar ativamente para a formação de uma <strong>cultura de pesquisa científica</strong> que posicione o HC-UFU como um polo de referência na geração de conhecimento dentro da rede EBSERH.</p>
    </div>
  </div>
  
  <div class="col" style="flex: 0 0 100%; max-width: 100%; padding: 0 15px; margin-bottom: 30px;">
    <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; height: 100%; border-bottom: 3px solid #17a2b8; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
      <h4 style="color: #17a2b8; margin-bottom: 15px; font-weight: 600;"><i class="fas fa-graduation-cap"></i> Apoio à Formação Acadêmica</h4>
      <p style="font-size: 1.05em; line-height: 1.6;">A plataforma foi especialmente desenvolvida para auxiliar residentes médicos e multiprofissionais do HC-UFU/EBSERH na busca por mentores e orientadores qualificados para o desenvolvimento de seus Trabalhos de Conclusão de Residência (TCRs). Reconhecemos a importância desta etapa obrigatória em diversos programas de residência e buscamos facilitar este processo, garantindo maior qualidade na produção científica institucional.</p>
    </div>
  </div>
  
  <div class="col" style="flex: 0 0 100%; max-width: 100%; padding: 0 15px; margin-bottom: 30px;">
    <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; height: 100%; border-bottom: 3px solid #e83e8c; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
      <h4 style="color: #e83e8c; margin-bottom: 15px; font-weight: 600;"><i class="fas fa-link"></i> Conexão e Colaboração</h4>
      <p style="font-size: 1.05em; line-height: 1.6;">Criamos uma ponte entre ideias promissoras e expertise científica, conectando estudantes, residentes e outros profissionais da saúde a mentores experientes. Esta colaboração potencializa a qualidade das pesquisas desenvolvidas e contribui significativamente para avanços na área da saúde, beneficiando diretamente a assistência aos pacientes e o desenvolvimento institucional.</p>
    </div>
  </div>
</div>
  
  <div class="row" style="display: flex; flex-wrap: wrap; margin: 0 -15px;">
    <div class="col" style="flex: 0 0 50%; max-width: 50%; padding: 0 15px; margin-bottom: 30px;">
      <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; height: 100%; border-bottom: 3px solid #17a2b8;">
        <h4 style="color: #17a2b8; margin-bottom: 15px;"><i class="fas fa-cogs"></i> Funcionamento</h4>
        <p>O DeuMatchNaPesquisa.com utiliza técnicas de "text mining" entre os termos inseridos na caixa de "objetivos do projeto" e o conteúdo registrado no currículo Lattes dos pesquisadores.</p>
      </div>
    </div>
    
    <div class="col" style="flex: 0 0 50%; max-width: 50%; padding: 0 15px; margin-bottom: 30px;">
      <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; height: 100%; border-bottom: 3px solid #e83e8c;">
        <h4 style="color: #e83e8c; margin-bottom: 15px;"><i class="fas fa-chart-bar"></i> Compatibilidade</h4>
        <p>O resultado do cruzamento é expresso na coluna "COMPATIBILIDADE", que é um valor percentual. Quanto maior a "COMPATIBILIDADE", maior a afinidade do pesquisador com os "objetivos do projeto".</p>
      </div>
    </div>
  </div>
  
  <div class="row" style="display: flex; flex-wrap: wrap; margin: 0 -15px;">
    <div class="col" style="flex: 0 0 50%; max-width: 50%; padding: 0 15px; margin-bottom: 30px;">
      <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; height: 100%; border-bottom: 3px solid #17a2b8;">
        <h4 style="color: #17a2b8; margin-bottom: 15px;"><i class="fas fa-calculator"></i> Algoritmo</h4>
        <p>O algoritmo realiza a leitura de diversos campos do Lattes, incluindo: artigos publicados, capítulos e livros publicados, palestras proferidas, apresentação de trabalhos e inclusive o resumo do currículo. Todos esses dados são padronizados, traduzidos para o português e então, a partir dessa base, é calculada a "COMPATIBILIDADE".</p>
      </div>
    </div>
    
    <div class="col" style="flex: 0 0 50%; max-width: 50%; padding: 0 15px; margin-bottom: 30px;">
      <div style="background-color: #f8f9fa; border-radius: 10px; padding: 20px; height: 100%; border-bottom: 3px solid #e83e8c;">
        <h4 style="color: #e83e8c; margin-bottom: 15px;"><i class="fas fa-laptop-code"></i> Tecnologia</h4>
        <p>Essa aplicação foi desenvolvida em linguagem R, utilizando o framework Shiny para criar a interface web interativa. A análise de texto e mineração de dados foi implementada com uso das bibliotecas tm, openxlsx e outras ferramentas específicas para processamento de linguagem natural.</p>
      </div>
    </div>
  </div>
  
  <div style="background-color: #f8f9fa; border-radius: 10px; padding: 25px; margin-top: 15px; border-left: 5px solid #17a2b8; text-align: center;">
    <h4 style="color: #17a2b8; margin-bottom: 15px;"><i class="fas fa-question-circle"></i> Dúvidas?</h4>
    <p style="font-size: 1.1em;">A <strong style="color: #e83e8c;">Unidade de Gestão da Inovação Tecnológica em Saúde</strong> - UGITS do HC-UFU/EBSERH está à disposição para esclarecer qualquer dúvida ou fornecer mais informações.</p>
    
    <div style="display: flex; justify-content: center; gap: 30px; margin-top: 20px;">
      <a href="https://wa.me/553432182323?text=Tenho%20dúvidas%20sobre%20o%20DEUMATCHNAPESQUISA" target="_blank" class="btn btn-lg" style="background-color: #25D366; color: white; text-decoration: none; padding: 10px 20px; border-radius: 5px; display: inline-flex; align-items: center;">
        <i class="fab fa-whatsapp" style="margin-right: 8px;"></i> WhatsApp
      </a>
      
      <a href="mailto:ugits.hc-ufu@ebserh.gov.br?subject=Dúvida%20sobre%20o%20DeuMatchNaPesquisa" class="btn btn-lg" style="background-color: #e83e8c; color: white; text-decoration: none; padding: 10px 20px; border-radius: 5px; display: inline-flex; align-items: center;">
        <i class="fas fa-envelope" style="margin-right: 8px;"></i> Email
      </a>
    </div>
  </div>
</div>
'

contato_html <- '
<div style="text-align: center;">
  <img src="https://i.ibb.co/4N9y2mX/logo.png" width="200px" alt="DeuMatchNaPesquisa"/><br /><br />
  <img src="https://www.gov.br/ebserh/pt-br/hospitais-universitarios/regiao-sudeste/hc-ufu/logos/hc-ufu-assinatura.png" width="400px" alt="HC-UFU / EBSERH" /><br><br>
</div>

<div class="container" style="max-width: 900px; margin: 0 auto;">
  <div class="card" style="border-radius: 15px; box-shadow: 0 8px 16px rgba(0,0,0,0.1); padding: 0; margin-bottom: 40px; overflow: hidden; border: none;">
    <div style="background: linear-gradient(135deg, #17a2b8 0%, #e83e8c 100%); padding: 30px; text-align: center;">
      <h2 style="color: white; margin-bottom: 10px; font-weight: 700;"><i class="fas fa-headset"></i> Entre em contato conosco</h2>
      <p style="color: rgba(255,255,255,0.9); font-size: 1.1em;">
        A <strong>Unidade de Gestão da Inovação Tecnológica em Saúde</strong> - UGITS do HC-UFU/EBSERH está à disposição para esclarecer dúvidas, receber sugestões ou auxiliar no uso da plataforma.
      </p>
    </div>
    
    <div style="padding: 30px;">
      <div class="row" style="display: flex; flex-wrap: wrap; margin: 0 -15px;">
        <div class="col" style="flex: 0 0 50%; max-width: 50%; padding: 0 15px; margin-bottom: 30px;">
          <div class="hoverable" style="background-color: #f8f9fa; border-radius: 10px; padding: 25px; height: 100%; text-align: center; transition: transform 0.3s; cursor: pointer;">
            <i class="fas fa-phone-alt" style="font-size: 48px; color: #17a2b8; margin-bottom: 20px;"></i>
            <h4 style="color: #17a2b8; margin-bottom: 15px;">Telefone</h4>
            <p style="font-size: 1.2em; font-weight: 500;">(34) 3218-2323</p>
            <p style="color: #6c757d; margin-top: 10px;">Segunda a sexta-feira, 8h às 16h</p>
          </div>
        </div>
        
        <div class="col" style="flex: 0 0 50%; max-width: 50%; padding: 0 15px; margin-bottom: 30px;">
          <div class="hoverable" style="background-color: #f8f9fa; border-radius: 10px; padding: 25px; height: 100%; text-align: center; transition: transform 0.3s; cursor: pointer;">
            <i class="fas fa-envelope" style="font-size: 48px; color: #e83e8c; margin-bottom: 20px;"></i>
            <h4 style="color: #e83e8c; margin-bottom: 15px;">Email</h4>
            <a href="mailto:ugits.hc-ufu@ebserh.gov.br" style="font-size: 1.2em; color: #e83e8c; text-decoration: none; font-weight: 500;">ugits.hc-ufu@ebserh.gov.br</a>
            <p style="color: #6c757d; margin-top: 10px;">Respondemos em até 48 horas úteis</p>
          </div>
        </div>
      </div>
      
      <div class="row" style="display: flex; flex-wrap: wrap; margin: 0 -15px;">
        <div class="col" style="flex: 0 0 50%; max-width: 50%; padding: 0 15px; margin-bottom: 30px;">
          <div class="hoverable" style="background-color: #f8f9fa; border-radius: 10px; padding: 25px; height: 100%; text-align: center; transition: transform 0.3s; cursor: pointer;">
            <i class="fab fa-whatsapp" style="font-size: 48px; color: #25D366; margin-bottom: 20px;"></i>
            <h4 style="color: #25D366; margin-bottom: 15px;">WhatsApp</h4>
            <a href="https://wa.me/553432182323?text=Estou%20escrevendo%20para%20falar%20sobre%20o%20DEUMATCHNAPESQUISA.COM" target="_blank" style="font-size: 1.2em; color: #25D366; text-decoration: none; font-weight: 500;">(34) 3218-2323</a>
            <p style="color: #6c757d; margin-top: 10px;">Mensagens respondidas em horário comercial</p>
          </div>
        </div>
<div class="col" style="flex: 0 0 50%; max-width: 50%; padding: 0 15px; margin-bottom: 30px;">
  <a href="https://maps.app.goo.gl/EYYmrS2ZRCSsbvLc7" target="_blank" style="text-decoration: none;">
    <div class="hoverable" style="background-color: #f8f9fa; border-radius: 10px; padding: 25px; height: 100%; text-align: center; transition: transform 0.3s; cursor: pointer;">
      <i class="fas fa-map-marker-alt" style="font-size: 48px; color: #dc3545; margin-bottom: 20px;"></i>
      <h4 style="color: #dc3545; margin-bottom: 15px;">Endereço</h4>
      <p style="font-size: 1.2em; font-weight: 500; color: #333;">BLOCO 8F - R. República do Piratini, 1418</p>
      <p style="color: #6c757d; margin-top: 10px;">Umuarama, Uberlândia - MG, 38402-028 • Brasil</p>
    </div>
  </a>
</div>
      </div>
      
      <div style="background-color: #f0f9fa; border-radius: 10px; padding: 25px; margin-top: 20px; text-align: center; border-left: 5px solid #17a2b8;">
        <h4 style="color: #17a2b8; margin-bottom: 15px;"><i class="fas fa-calendar-check"></i> Horário de Atendimento</h4>
        <p style="font-size: 1.1em;">
          Segunda a sexta-feira, das <strong>8h00 às 16h00</strong>
        </p>
        <p style="color: #6c757d; margin-top: 10px;">
          Respondemos mensagens e emails durante o horário comercial. Em caso de urgência, favor entrar em contato por telefone.
        </p>
      </div>
    </div>
    
    <div style="background-color: #f8f9fa; padding: 25px; text-align: center; border-top: 1px solid #eaeaea;">
      <p style="font-size: 1.2em; font-weight: bold; color: #17a2b8;">Ficaremos felizes em responder todas as suas dúvidas! 😊</p>
    </div>
  </div>
</div>
'

limitacoes_html <- '
<div style="text-align: center;">
  <img src="https://i.ibb.co/4N9y2mX/logo.png" width="200px" alt="DeuMatchNaPesquisa"/><br /><br />
  <img src="https://www.gov.br/ebserh/pt-br/hospitais-universitarios/regiao-sudeste/hc-ufu/logos/hc-ufu-assinatura.png" width="400px" alt="HC-UFU/EBSERH" /><br><br>
</div>

<div class="container">
  <div class="alert alert-info" style="border-radius: 8px; margin-bottom: 25px;">
    <h3>💡 Sobre este MVP (Minimum Viable Product)</h3>
  </div>
  
  <p>O DeuMatchNaPesquisa.com que você está utilizando neste momento é uma <strong>versão experimental</strong>, desenvolvida em linguagem R, que visa demonstrar a viabilidade e potencial da ideia. Este MVP foi criado como uma prova de conceito para validar a abordagem de conectar estudantes e pesquisadores através de mineração de dados do Currículo Lattes.</p>
  
  <div class="alert alert-warning" style="border-radius: 8px; margin: 25px 0; background-color: #fff3cd; border-color: #ffeeba;">
    <h4>⚠️ Limitações importantes desta versão:</h4>
    <ul style="padding-left: 20px; margin-top: 15px;">
      <li>Os dados do Lattes apresentados foram minerados no início de 2023 e podem estar desatualizados</li>
      <li>A base de dados atual utiliza informações públicas do Lattes, sem consentimento específico dos pesquisadores para esta aplicação</li>
      <li>O algoritmo de compatibilidade é uma versão simplificada, que ainda pode ser aprimorado</li>
      <li>A interface e recursos de comunicação são limitados nesta versão de demonstração</li>
      <li>Os resultados podem não representar com precisão todas as possibilidades de orientação disponíveis</li>
    </ul>
  </div>
  
  <h4 style="color: #17a2b8; margin-top: 30px; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px;">🚀 Planos para a versão oficial</h4>
  <p>Para evolução deste projeto para uma versão oficial e completa, pretendemos:</p>
  <ul style="padding-left: 20px; margin-top: 15px;">
    <li>✅ Realizar a recoleta formal de consentimento de todos os pesquisadores incluídos na base</li>
    <li>✅ Explicar detalhadamente os objetivos e funcionamento do sistema aos participantes</li>
    <li>✅ Atualizar os dados do Lattes para refletir as informações mais recentes</li>
    <li>✅ Aprimorar o algoritmo de compatibilidade com técnicas mais avançadas de processamento de linguagem natural</li>
    <li>✅ Desenvolver uma interface mais robusta e com mais recursos</li>
    <li>✅ Implementar um sistema de feedback para melhorar continuamente a experiência</li>
    <li>✅ Estabelecer uma política clara de privacidade e uso de dados</li>
  </ul>
  
  <h4 style="color: #17a2b8; margin-top: 30px; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px;">🔍 Propósito deste MVP</h4>
  <p>Este projeto experimental visa demonstrar como a tecnologia pode facilitar a conexão entre alunos com ideias de pesquisa e potenciais orientadores com expertise nas áreas de interesse. A versão atual serve como uma demonstração funcional do conceito, permitindo que as partes interessadas visualizem o potencial da ferramenta e contribuam com feedback para seu desenvolvimento.</p>
  
  <p>Agradecemos sua compreensão quanto às limitações desta versão e valorizamos qualquer feedback que possa nos ajudar a aprimorar este projeto.</p>
  
<div class="alert alert-success" style="border-radius: 8px; margin-top: 25px; background-color: #d4edda; border-color: #c3e6cb;">
  <p>Para dúvidas, sugestões ou mais informações, entre em contato com a UGITS do HC-UFU/EBSERH pelos canais disponíveis na aba "Contato".</p>
</div>
</div>
'

# Função para limpar texto
limpa_texto <- function(texto) {
  if (is.null(texto) || length(texto) == 0 || all(nchar(texto) == 0)) {
    return(list())
  }
  
  # Converter para minúsculas
  texto <- tolower(texto)
  # Remover pontuação
  texto <- removePunctuation(texto)
  # Remover números
  texto <- removeNumbers(texto)
  # Remover espaços extras
  texto <- stripWhitespace(texto)
  # Remover stopwords
  texto <- removeWords(texto, c(stopwords("pt-br"), "fez", "grande"))
  
  # Dividir em palavras e obter palavras únicas
  palavras <- strsplit(texto, " ")[[1]]
  palavras <- unique(palavras[palavras != ""])
  
  return(palavras)
}

# Função para criar botão
botao <- function(index) {
  actionButton(
    inputId = paste0("botao_", index),
    label = HTML('<i class="fab fa-whatsapp" style="margin-right: 6px;"></i> Contato via WhatsApp'),
    onclick = sprintf("Shiny.setInputValue('botao_clicado', %s, {priority: 'event'})", index),
    class = "btn-success",
    style = "white-space: nowrap;"
  )
}

# Define UI
ui <- fluidPage(
  tags$head(
    # Add this to fix JavaScript issues with hover effects
    tags$script("
      $(document).on('mouseenter', '.hoverable', function() {
        $(this).css('transform', 'translateY(-10px)');
      });
      $(document).on('mouseleave', '.hoverable', function() {
        $(this).css('transform', 'translateY(0)');
      });
    "),
    tags$style(HTML("
      /* Estilos gerais */
      body {
        font-family: 'Roboto', sans-serif;
        color: #333;
        line-height: 1.6;
      }
      
      /* Estilização do título */
      .nav-tabs > li > a {
        font-weight: bold;
        color: #17a2b8;
        border-radius: 5px 5px 0 0;
      }
      
      .nav-tabs > li.active > a, 
      .nav-tabs > li.active > a:hover, 
      .nav-tabs > li.active > a:focus {
        color: #e83e8c;
        border-top: 3px solid #e83e8c;
      }
      
      /* Alertas e mensagens */
      .alert {
        margin-top: 20px;
        border-radius: 8px;
      }
      
      /* Botões */
      .btn-primary {
        background-color: #17a2b8;
        border-color: #17a2b8;
        transition: all 0.3s ease;
      }
      
      .btn-primary:hover {
        background-color: #138496;
        border-color: #138496;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
      }
      
      /* Cards e containers */
      .card {
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        margin-bottom: 20px;
        transition: transform 0.3s;
      }
      
      .card:hover {
        transform: translateY(-5px);
      }
      
      /* Tabela de resultados */
      .dataTable {
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
      }
      
      .dataTable thead th {
        background-color: #f8f9fa;
        color: #17a2b8;
        font-weight: bold;
      }
      
/* Footer */
.footer {
  margin-top: 50px;
  padding: 20px 0;
  background-color: #f8f9fa;
  border-top: 1px solid #eaeaea;
  font-size: 1.2rem; /* Alterado de 0.9rem para 1rem */
}
      
      /* Hoverable elements */
      .hoverable {
        transition: transform 0.3s ease;
      }
   ")), 
    tags$link(rel = "stylesheet", href = "https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap"),
    tags$link(rel = "stylesheet", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css")
  ),
  
  # Título da aplicação
  titlePanel(
    title = div(
      style = "text-align: center; margin-bottom: 10px;",
      span("DeuMatchNaPesquisa", style = "color: #17a2b8; font-weight: bold;"),
      span(" - Conectando ideias e mentores 🔍", style = "color: #6c757d;")
    ),
    windowTitle = "DeuMatchNaPesquisa.com"
  ),
  
  tabsetPanel(
    tabPanel("Pesquisar",
             tags$div(style = "text-align: center; margin: 20px 0;", 
                      tags$img(src = "https://i.ibb.co/4N9y2mX/logo.png", 
                               width = 200, alt = "DeuMatchNaPesquisa",
                               style = "margin-bottom: 00px;"),br(),
                      tags$img(src = "https://www.gov.br/ebserh/pt-br/hospitais-universitarios/regiao-sudeste/hc-ufu/logos/hc-ufu-assinatura.png", 
                               width = 400, alt = "HC-UFU/EBSERH",
                               style = "margin-bottom: 00px;")
                      ),
             

             # Aviso MVP
             tags$div(
               class = "alert alert-warning",
               style = "text-align: center; max-width: 800px; margin: 1em auto; border-radius: 10px;",
               tags$h4(style = "color: #856404; font-weight: bold;", "⚠️ Versão Experimental (MVP)"),
               tags$p("Esta é uma versão de demonstração com dados do Lattes de 2023. Para saber mais sobre as limitações desta versão, ",
                      tags$a(href = "#", onclick = "$(\'a[data-value=\\\'Limitações do MVP\\\']\').tab(\'show\'); return false;", 
                             style = "color: #e83e8c; font-weight: bold;", "clique aqui."))
             ),
             
             # Campo de entrada principal
             tags$div(style = "text-align: center; max-width: 600px; margin: 2em auto; background-color: #f8f9fa; padding: 20px; border-radius: 10px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);",
                      textInput(inputId = "caption", 
                                label = tags$span(style="font-weight: bold; color: #17a2b8; font-size: 16px;",
                                                  '✨ Digite os objetivos do seu projeto de pesquisa...'), 
                                "", width="100%"),
                      verbatimTextOutput("value")),
             
             # Opções de filtro
             tags$div(style = "text-align: center; background-color: #f8f9fa; margin: 20px auto; max-width: 600px; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);",
                      tags$div(style = "display: inline-block; width: 220px; margin-right: 20px;", 
                               checkboxInput("resumo", tags$span(style = "color: #17a2b8;", "Mostrar Resumo do Lattes 📄"), value = FALSE)),
                      tags$div(style = "display: inline-block; width: 220px; margin-left: 20px;", 
                               checkboxInput("regiao", tags$span(style = "color: #17a2b8;", "Triângulo Mineiro 📍"), value = TRUE))),
             
             # Tabela de resultados
             tags$div(style = "text-align: center; background-color: none; margin: auto; padding: 20px;",
                      withLoader(DT::dataTableOutput('table', width = "100%"), 
                                 type = "html", 
                                 loader = "loader5"), 
                      div(id = "input_botao_clicado", style = "display: none;")),
             
             # Rodapé personalizado
             tags$div(class = "footer", style = "text-align: center; margin-top: 40px; padding: 30px; background-color: #f8f9fa; border-top: 1px solid #eaeaea;",
                      tags$div(style = "max-width: 800px; margin: 0 auto;",
                               tags$h3(style = "color: #17a2b8; margin-bottom: 20px;", "🔍 DeuMatchNaPesquisa", 
                                       tags$small(style = "color: #6c757d; display: block; margin-top: 8px; font-size: 0.8em;", "Conectando ideias e pesquisadores")),
                               tags$p(style = "font-size: 1.1em; margin-bottom: 20px;", 
                                      "Desenvolvido pela ", 
                                      tags$strong(style = "color: #e83e8c;", "🔬 UGITS"), 
                                      " (Unidade de Gestão da Inovação Tecnológica em Saúde) do HC-UFU/EBSERH"),
                               
                               tags$hr(style = "margin: 20px 0; border-color: #eaeaea;"),
                               
                               tags$div(style = "display: flex; justify-content: center; flex-wrap: wrap; gap: 40px; margin: 30px 0;",
                                        tags$div(
                                          tags$p(style = "font-size: 1.1em;", 
                                                 tags$i(class = "fas fa-map-marker-alt", style = "color: #dc3545; margin-right: 8px;"), 
                                                 tags$strong("Endereço:"), 
                                                 tags$a(href = "https://maps.app.goo.gl/EYYmrS2ZRCSsbvLc7", target = "_blank", 
                                                        style = "color: #17a2b8;",
                                                        "R. República do Piratini, 1418 - Bloco 8F")),
                                          tags$p(style = "color: #6c757d; font-size: 1em; margin-top: -5px;", 
                                                 "Umuarama, Uberlândia - MG, 38402-028")
                                        ),
                                        
                                        tags$div(
                                          tags$p(style = "font-size: 1.1em;", 
                                                 tags$i(class = "fas fa-clock", style = "color: #ffc107; margin-right: 8px;"), 
                                                 tags$strong("Horário:"), 
                                                 "Segunda a sexta-feira, 8h às 16h")
                                        ),
                                        
                                        tags$div(
                                          tags$p(style = "font-size: 1.1em;", 
                                                 tags$i(class = "fas fa-envelope", style = "color: #e83e8c; margin-right: 8px;"), 
                                                 tags$strong("Email:"), 
                                                 tags$a(href = "mailto:ugits.hc-ufu@ebserh.gov.br", 
                                                        style = "color: #17a2b8;", 
                                                        "ugits.hc-ufu@ebserh.gov.br"))
                                        ),
                                        
                                        tags$div(
                                          tags$p(style = "font-size: 1.1em;", 
                                                 tags$i(class = "fas fa-phone", style = "color: #28a745; margin-right: 8px;"), 
                                                 tags$strong("WhatsApp:"), 
                                                 tags$a(href = "https://wa.me/553432182323", 
                                                        style = "color: #17a2b8;", 
                                                        "(34) 3218-2323"))
                                        )
                               ),
                               
                               tags$hr(style = "margin: 20px 0; border-color: #eaeaea;"),
                               
                               tags$p(style = "color: #333; font-size: 1em;", 
                                      "© 2025 HC-UFU/EBSERH. Todos os direitos reservados.", 
                                      tags$br(),
                                      "Desenvolvido pela ", 
                                      tags$a(href = "mailto:ugits.hc-ufu@ebserh.gov.br", 
                                             style = "color: #17a2b8; font-weight: bold;", 
                                             "Unidade de Gestão da INovação Tecnológica em Saúde"), 
                                      " - UGITS HC-UFU/EBSERH")
                      ))
    ),
    tabPanel("É Pesquisador? Cadastre-se",
             # Use HTML() instead of includeMarkdown()
             HTML(pesquisador_html)
    ),
    tabPanel("Sobre o Projeto",
             # Use HTML() instead of includeMarkdown()
             HTML(sobre_html)
    ),
    tabPanel("Contato",
             # Use HTML() instead of includeMarkdown()
             HTML(contato_html)
    ),
    tabPanel("Limitações do MVP",
             # Use HTML() instead of includeMarkdown()
             HTML(limitacoes_html)
    )
  )
)
      
# Define server
server <- function(input, output, session) {

  # Pop-up de boas-vindas / como usar (exibido ao abrir o site)
  showModal(modalDialog(
    title = div(style = "color: #17a2b8; font-weight: 700; text-align: center;",
                HTML('<i class="fas fa-hand-sparkles" style="margin-right: 8px;"></i>'),
                "Bem-vindo(a) ao DeuMatchNaPesquisa.com"),

    div(
      style = "max-width: 100%;",

      p(style = "text-align: center; color: #6c757d; font-size: 1.05em; margin-bottom: 20px;",
        "Encontre pesquisadores(as) compatíveis com a sua ideia de projeto em ", strong("3 passos"), ":"),

      # Passo 1
      div(style = "display: flex; align-items: flex-start; gap: 14px; background-color: #f8f9fa; border-radius: 10px; padding: 14px; margin-bottom: 12px;",
          span(style = "flex: 0 0 38px; height: 38px; width: 38px; background-color: #17a2b8; color: #fff; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 1.1em;", "1"),
          div(style = "text-align: left;",
              strong(style = "color: #17a2b8;", "Digite a área que deseja estudar"),
              p(style = "margin: 4px 0 0; color: #555;",
                "Descreva os objetivos do seu projeto. O sistema filtra apenas os pesquisadores que desenvolvem estudos ou têm know-how naquela área."))),

      # Passo 2
      div(style = "display: flex; align-items: flex-start; gap: 14px; background-color: #f8f9fa; border-radius: 10px; padding: 14px; margin-bottom: 12px;",
          span(style = "flex: 0 0 38px; height: 38px; width: 38px; background-color: #e83e8c; color: #fff; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 1.1em;", "2"),
          div(style = "text-align: left;",
              strong(style = "color: #e83e8c;", "Escolha um(a) pesquisador(a) e clique no contato"),
              p(style = "margin: 4px 0 0; color: #555;",
                "Veja a lista de compatibilidade e clique no botão verde ",
                HTML('<i class="fab fa-whatsapp" style="color: #25D366;"></i>'),
                " ", strong("Contato via WhatsApp"), " do(a) escolhido(a)."))),

      # Passo 3
      div(style = "display: flex; align-items: flex-start; gap: 14px; background-color: #f8f9fa; border-radius: 10px; padding: 14px; margin-bottom: 16px;",
          span(style = "flex: 0 0 38px; height: 38px; width: 38px; background-color: #25D366; color: #fff; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 1.1em;", "3"),
          div(style = "text-align: left;",
              strong(style = "color: #25D366;", "Acione a Equipe da UGPESQ"),
              p(style = "margin: 4px 0 0; color: #555;",
                "No pop-up, clique no botão verde. A Equipe da ", strong("UGPESQ"),
                " tentará intermediar o contato com o(a) pesquisador(a)."))),

      # Aviso importante
      div(style = "background-color: #fff3cd; border: 2px solid #ffc107; border-radius: 10px; padding: 14px; text-align: left;",
          p(style = "color: #856404; margin: 0;",
            HTML('<i class="fas fa-exclamation-triangle" style="margin-right: 6px;"></i>'),
            strong("Importante: "),
            "a UGPESQ envia um e-mail e aguarda a resposta, mas ", strong("o contato não é garantido"),
            ". Se for importante para você, ", strong("busque o contato também por conta própria pelo Lattes"), "."))
    ),

    footer = modalButton(div(HTML('<i class="fas fa-rocket" style="margin-right: 6px;"></i>'), "Começar!")),
    size = "l",
    easyClose = TRUE
  ))

  # Carregar dados apenas uma vez na inicialização
  base_path <- "base_dados/base_expertise.xlsx"
  base_comp_path <- "base_dados/base_expertise_complemento.xlsx"
  
  # Verificar se os arquivos existem
  if (!file.exists(base_path)) {
    stop("Arquivo base_expertise.xlsx não encontrado! 📁❌")
  }
  
  if (!file.exists(base_comp_path)) {
    stop("Arquivo base_expertise_complemento.xlsx não encontrado! 📁❌")
  }
  
  # Carregar os dados
  base <- read.xlsx(base_path)
  base_comp <- read.xlsx(base_comp_path)
  
  # Processar dados base
  fotos <- base$LINK.FOTO
  base$FOTO_HTML <- paste0('<img src="', fotos, '" height="52" style="border-radius: 50%; object-fit: cover;">')
  
  # Armazenar dados
  valores <- reactiveValues(
    base_completa = base,
    base_comp = base_comp,
    resultados = NULL,
    pesquisador_selecionado = NULL
  )
  
  # Funções de filtro
  showTriangulo <- reactive({
    input$regiao
  })
  
  showResumo <- reactive({
    input$resumo
  })
  
  # Exibir texto digitado
  output$value <- renderText({
    input$caption
  })
  
  # Calcular compatibilidade quando o usuário digita
  observe({
    req(input$caption)
    
    if (nchar(input$caption) == 0) {
      valores$resultados <- NULL
      return(NULL)
    }
    
    # Limpar entrada do usuário
    texto_usuario <- limpa_texto(input$caption)
    
    if (length(texto_usuario) == 0) {
      valores$resultados <- NULL
      return(NULL)
    }
    
    # Criar dataframe de resultados
    resultados <- data.frame(
      ID = 1:nrow(valores$base_completa),
      COMPATIBILIDADE = numeric(nrow(valores$base_completa))
    )
    
    # Calcular compatibilidade para cada pesquisador
    for (i in 1:nrow(resultados)) {
      # Combinar resumo e dados concatenados
      texto_pesquisador <- paste(
        valores$base_completa$RESUMO.LATTES[i],
        valores$base_completa$CONCATENADO[i]
      )
      
      # Limpar texto do pesquisador
      palavras_pesquisador <- limpa_texto(texto_pesquisador)
      
      # Calcular compatibilidade
      if (length(texto_usuario) > 0 && length(palavras_pesquisador) > 0) {
        matches <- sum(texto_usuario %in% palavras_pesquisador)
        resultados$COMPATIBILIDADE[i] <- round((matches / length(texto_usuario)) * 100, 1)
      } else {
        resultados$COMPATIBILIDADE[i] <- 0
      }
    }
    
    # Ordenar resultados por compatibilidade
    resultados <- resultados[order(resultados$COMPATIBILIDADE, decreasing = TRUE), ]
    
    # Filtrar compatibilidade zero
    resultados <- resultados[resultados$COMPATIBILIDADE > 0, ]
    
    if (nrow(resultados) > 0) {
      # Criar dataframe de resultados finais
      resultados_finais <- data.frame(
        FOTO = valores$base_completa$FOTO_HTML[resultados$ID],
        COMPATIBILIDADE = paste0('<span style="font-weight: bold; color: ', 
                                 ifelse(resultados$COMPATIBILIDADE >= 70, '#28a745', 
                                        ifelse(resultados$COMPATIBILIDADE >= 40, '#17a2b8', '#6c757d')),
                                 ';">', resultados$COMPATIBILIDADE, '%</span>'),
        NOME = paste0('<span style="font-weight: bold; color: #e83e8c;">', 
                      valores$base_completa$NOME[resultados$ID], '</span>'),
        LINK = paste0('<a href="', valores$base_completa$LINK.LATTES[resultados$ID], 
                      '" target="_blank" class="btn btn-sm btn-outline-primary" style="white-space: nowrap;">',
                      '<i class="fas fa-external-link-alt"></i> Lattes</a>'),
        FORMACAO = valores$base_comp$formacao_pesq[resultados$ID],
        AREA = valores$base_comp$area_e_subarea_pesq[resultados$ID],
        LOCALIDADE = valores$base_comp$regiao[resultados$ID],
        stringsAsFactors = FALSE
      )
      
      # Filtrar por região se necessário
      if (showTriangulo()) {
        resultados_finais <- resultados_finais[resultados_finais$LOCALIDADE == "Triângulo Mineiro", ]
      }
      
      if (nrow(resultados_finais) > 0) {
        # Adicionar resumo se solicitado
        if (showResumo()) {
          resultados_finais$RESUMO <- paste0('<div style="max-height: 100px; overflow-y: auto; font-size: 0.9em; color: #6c757d;">', 
                                             valores$base_completa$RESUMO.LATTES[resultados$ID], '</div>')
        }
        
        # Adicionar botão de contato
        resultados_finais$CONTATO <- sapply(1:nrow(resultados_finais), function(i) {
          as.character(botao(i))
        })
        
        # Armazenar IDs para contato
        valores$resultados_ids <- resultados$ID[resultados_finais$LOCALIDADE == resultados_finais$LOCALIDADE]
        valores$resultados <- resultados_finais
      } else {
        valores$resultados <- data.frame(
          Mensagem = paste0('<div class="alert alert-info" style="text-align: center; margin: 30px;">',
                            '<i class="fas fa-search" style="font-size: 2em; color: #17a2b8; margin-bottom: 15px;"></i><br>',
                            '<strong>Nenhum resultado encontrado na região selecionada.</strong><br>',
                            '🔍 Tente outros termos ou desmarque o filtro de região.',
                            '</div>'),
          stringsAsFactors = FALSE
        )
      }
    } else {
      valores$resultados <- data.frame(
        Mensagem = paste0('<div class="alert alert-warning" style="text-align: center; margin: 30px;">',
                          '<i class="fas fa-exclamation-triangle" style="font-size: 2em; color: #ffc107; margin-bottom: 15px;"></i><br>',
                          '<strong>Nenhum resultado compatível encontrado.</strong><br>',
                          '🤔 Experimente usar termos mais gerais ou diferentes palavras-chave.',
                          '</div>'),
        stringsAsFactors = FALSE
      )
    }
  })
  
  # Renderizar tabela de resultados
  output$table <- DT::renderDataTable({
    if (is.null(valores$resultados)) {
      return(NULL)
    }
    
    if ("Mensagem" %in% colnames(valores$resultados)) {
      return(DT::datatable(
        valores$resultados,
        options = list(
          dom = 't',
          language = list(
            emptyTable = paste0('<div class="alert alert-primary" style="text-align: center; padding: 20px;">',
                                '<i class="fas fa-lightbulb" style="font-size: 2em; color: #17a2b8; margin-bottom: 15px;"></i><br>',
                                '<strong>Digite os objetivos do seu projeto</strong><br>',
                                'Para encontrar pesquisadores compatíveis com seu interesse de pesquisa.',
                                '</div>')
          )
        ),
        escape = FALSE,
        rownames = FALSE
      ))
    }
    
    DT::datatable(
      valores$resultados,
      escape = FALSE,
      rownames = FALSE,
      options = list(
        dom = 't', 
        pageLength = 150,
        language = list(
          emptyTable = paste0('<div class="alert alert-primary" style="text-align: center; padding: 20px;">',
                              '<i class="fas fa-lightbulb" style="font-size: 2em; color: #17a2b8; margin-bottom: 15px;"></i><br>',
                              '<strong>Digite os objetivos do seu projeto</strong><br>',
                              'Para encontrar pesquisadores compatíveis com seu interesse de pesquisa.',
                              '</div>')
        ),
        columnDefs = list(
          list(className = 'dt-center', targets = c(0, 1, 3, 6))
        )
      )
    ) %>%
      DT::formatStyle(
        columns = c('COMPATIBILIDADE', 'NOME', 'FORMACAO', 'AREA'),
        fontSize = '14px'
      )
  })
  
  # Tratar clique no botão de contato
  observeEvent(input$botao_clicado, {
    index <- as.numeric(input$botao_clicado)
    
    if (!is.null(valores$resultados) && nrow(valores$resultados) >= index) {
      valores$pesquisador_selecionado <- index

      # Extrair o nome do pesquisador (removendo as tags HTML)
      nome_pesquisador <- gsub("<.*?>", "", valores$resultados$NOME[index])

      # Extrair o link do Lattes do HTML da coluna LINK
      link_lattes <- sub('.*href="([^"]+)".*', '\\1', valores$resultados$LINK[index])
      if (identical(link_lattes, valores$resultados$LINK[index])) link_lattes <- ""

      # Montar a mensagem pré-formatada para o WhatsApp
      mensagem_wpp <- sprintf(
        "Olá, gostaria que a UGPESQ me ajudasse a entrar em contato com o(a) profissional %s (%s). Consegue fazer isso?",
        nome_pesquisador, link_lattes
      )
      url_whatsapp <- paste0("https://wa.me/553432182323?text=",
                             URLencode(mensagem_wpp, reserved = TRUE))

      showModal(modalDialog(
        title = div(style = "color: #25D366; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px; font-weight: 700;",
                    HTML('<i class="fab fa-whatsapp" style="margin-right: 8px;"></i>'),
                    paste0("Contato com ", nome_pesquisador)),

        div(
          style = "text-align: center; padding: 5px 0;",

          # Passo 1 - Acionar a UGPESQ via WhatsApp
          div(
            style = "background-color: #f0fff4; border: 1px solid #c3e6cb; border-radius: 10px; padding: 20px; margin-bottom: 20px;",
            HTML('<i class="fab fa-whatsapp" style="font-size: 42px; color: #25D366; margin-bottom: 10px;"></i>'),
            h4(style = "color: #25D366; margin-bottom: 8px; font-weight: 700;",
               "Precisa entrar em contato?"),
            p(style = "font-size: 1.05em; color: #333; margin-bottom: 18px;",
              "Envie um WhatsApp para ", strong("(34) 3218-2323"),
              " e a Equipe da ", strong("UGPESQ"), " tentará intermediar o contato com o(a) pesquisador(a)."),
            a(href = url_whatsapp, target = "_blank",
              class = "btn btn-success btn-lg",
              style = "padding: 12px 28px; font-weight: 700;",
              HTML('<i class="fab fa-whatsapp" style="margin-right: 8px;"></i>'),
              "Abrir WhatsApp")
          ),

          # AVISO IMPORTANTE - contato não garantido
          div(
            style = "background-color: #fff3cd; border: 2px solid #ffc107; border-radius: 10px; padding: 18px; margin-bottom: 20px; text-align: left;",
            h4(style = "color: #856404; font-weight: 700; text-align: center; margin-top: 0;",
               HTML('<i class="fas fa-exclamation-triangle" style="margin-right: 8px;"></i>'),
               "Importante: o contato NÃO é garantido"),
            p(style = "color: #856404; font-size: 1.0em; margin-bottom: 8px;",
              "A UGPESQ irá ", strong("enviar um e-mail ao pesquisador e aguardar a resposta"),
              ", mas, por causa do prazo curto, ", strong("nem sempre conseguimos retorno a tempo"), "."),
            p(style = "color: #856404; font-size: 1.0em; margin-bottom: 0;",
              HTML('<i class="fas fa-lightbulb" style="margin-right: 6px;"></i>'),
              strong("Se este contato é importante para você, busque-o também por conta própria"),
              " — pelos canais abaixo.")
          ),

          # Lattes e contatos do pesquisador - destacado
          div(
            style = "background-color: #e8f4f8; border: 1px solid #bee5eb; border-radius: 10px; padding: 18px;",
            h5(style = "color: #17a2b8; font-weight: 700; margin-top: 0;",
               HTML('<i class="fas fa-id-card" style="margin-right: 8px;"></i>'),
               "Contato direto pelo Lattes"),
            p(style = "color: #0c5460; font-size: 1.0em; margin-bottom: 14px;",
              "No currículo Lattes do(a) pesquisador(a) você encontra e-mail, instituição e outras formas de contato."),
            if (nzchar(link_lattes)) {
              a(href = link_lattes, target = "_blank",
                class = "btn btn-info btn-lg",
                style = "padding: 10px 24px; font-weight: 700; color: #fff;",
                HTML('<i class="fas fa-external-link-alt" style="margin-right: 8px;"></i>'),
                paste0("Abrir o Lattes de ", nome_pesquisador))
            } else {
              p(style = "color: #0c5460; font-style: italic;",
                "Busque o nome do(a) pesquisador(a) na ",
                a(href = "http://buscatextual.cnpq.br/buscatextual/busca.do", target = "_blank",
                  style = "font-weight: 700;", "Plataforma Lattes"), ".")
            }
          )
        ),

        footer = modalButton(div(icon("times"), "Fechar")),
        size = "m",
        easyClose = TRUE
      ))
    }
  })
}

# Executar a aplicação
shinyApp(ui, server)