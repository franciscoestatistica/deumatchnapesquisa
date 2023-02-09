library(fs)
#fs::dir_copy("/home/diego/deumatchnapesquisa/base_dados/", "/srv/shiny-server/")
#fs::file_delete("/srv/shiny-server/app.R")
#fs::file_copy("/home/diego/deumatchnapesquisa/app.R", "/srv/shiny-server/")
#sudo chmod -R 777 /srv/shiny-server
#sudo service nginx start
#sudo systemctl restart shiny-server

#NA PASTA SHINY-SERVER, FAZER OS LINKS SIMBÓLICOS
#sudo ln -s /home/diego/deumatchnapesquisa/app.R app.R
#sudo ln -s /home/diego/deumatchnapesquisa/base_dados base_dados

library(readr)
library(tm)
library(shiny)
library(DT)
library(shinycustomloader)
library(shinythemes)
library(markdown)
library(XML)
library(methods)
library(GetLattesData)
library(stringr)
library(readxl)
library(openxlsx)
library(rstudioapi) 
library(here)

atualiza_tudo <- function(tenho_que_atualizar) {
  if (tenho_que_atualizar == TRUE) {
    local <- c(as.String(paste0(here(),'/base_dados/lattes_zip/XML/0155959835062402.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0702503776284410.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0781143912494425.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0853119429254827.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0853955010849504.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0876204917775050.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0905836893898631.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/1989772308502986.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2102993403919035.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2317284338521807.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2325354616581027.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2522602422184941.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2551720593062355.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2926473300969422.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/3027146613276970.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/3085568175388061.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/3119879857641234.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/3383567271020768.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/3445472713137548.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/3684666180428347.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4247658527800602.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4506503088340718.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4569169833604734.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4761289879698286.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5069063202614497.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5202499767006928.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5426268075449547.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5481509221004874.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5507813606549292.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5641001295200928.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5697765739821467.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5717713170043491.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/6294763034488617.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/6371375983884303.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/7670714110574321.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/7919901048891187.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8042687654173704.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8066578384803071.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8135197639370701.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8144127951159683.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8154680248919577.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8188329050349666.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8217806694441815.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8393669280266372.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8618933240073994.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8806072933197265.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8931806861246338.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9335137326600884.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9554425330898629.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9899075789126981.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9865330615540205.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0027511648392947.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0170810924772574.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0292834654421299.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/0661209395751558.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/1437087614639311.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/1668285987305521.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2020742498984612.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2227139061867181.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2266112154431321.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2585330510002189.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2688436056239402.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2848812662814026.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2860179178643956.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/2864843574433034.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4026715152001626.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4063985480147162.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4083960867428854.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4130787400995997.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4247276247749935.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4482870742031147.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4600154820027807.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4671444108825248.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5502129635422495.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5984672774577439.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/6096599677518899.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/6872717162309834.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/6953578838978396.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/7181115937433821.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/7482509894855417.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/7703927782935094.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/7751742904886118.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8130370622150343.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8417190478657621.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8646633997201802.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8810916520087214.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8869553181466970.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/8924508898024445.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9070111459228197.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9092296814446576.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9108499010275081.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9416668271579542.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9652541311039940.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/3564706302936545.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/7340105957340705.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/1229329519982110.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9850033081267802.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/9047298948861003.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5452734249403936.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/5680475427208141.zip')),
               as.String(paste0(here(),'/base_dados/lattes_zip/XML/4980687835840176.zip')))
    
    
    l.out <- gld_get_lattes_data_from_zip(local)
    nomes_pesq <- l.out$tpesq[,1]$name
    #link_pesq <- paste0("http://lattes.cnpq.br/",D:/DEU_MATCH_NA_PESQUISA/base_dados/lattes_zip/XML/8806072933197265.zip(local,93,108))
    link_pesq <- paste0("http://lattes.cnpq.br/",substr(local,52,67))
    
    formacao_pesq <- l.out$tpesq[,6][[1]]
    area_pesq <- gsub("_", " ", l.out$tpesq[,11][[1]]) 
    subarea_pesq <- l.out$tpesq[,12][[1]]
    area_e_subarea_pesq<- paste0(area_pesq," (",subarea_pesq,")")
    regiao <- rep("Triângulo Mineiro", length(nomes_pesq))
    
    #write.csv2(as.data.frame(l.out$tpesq), file="tpesq.csv", sep=";", fileEncoding = "iso-8859-1")
    #write.csv2(as.data.frame(l.out$tpublic.published), file="tpublic.published.csv", sep=";", fileEncoding = "iso-8859-1") A
    #write.csv2(as.data.frame(l.out$tpublic.accepted), file="tpublic.accepted.csv", sep=";", fileEncoding = "iso-8859-1") B
    #write.csv2(as.data.frame(l.out$tsupervisions), file="tsupervisions.csv", sep=";", fileEncoding = "iso-8859-1") -- NCO USAREI --
    #write.csv2(as.data.frame(l.out$tbooks), file="tbooks.csv", sep=";", fileEncoding = "iso-8859-1") C
    #write.csv2(as.data.frame(l.out$tconferences), file="tconferences.csv", sep=";", fileEncoding = "iso-8859-1") D
    
    concatenado_pesquisador <- matrix("",length(link_pesq),6)
    #concatenado_pesquisador[,1] <- substr(local,93,108)
    concatenado_pesquisador[,1] <- substr(local,52,67)
    concatenado_pesquisador[,2] <- nomes_pesq
    #montando_coluna_de_artigos_publicados A
    artigos_publicados_pesq <- cbind(substr(l.out$tpublic.published[,1][[1]],1,20), l.out$tpublic.published[,2][[1]], l.out$tpublic.published[,3][[1]])
    for (i in 1:length(unique(concatenado_pesquisador[,1]))){
      #i=1
      for (j in 1:length(artigos_publicados_pesq[artigos_publicados_pesq[,1] %in% paste0(concatenado_pesquisador[i,1],".zip"),3]) ){
        #j=2
        concatenado_pesquisador[i,3] <- paste0(concatenado_pesquisador[i,3]," ", artigos_publicados_pesq[artigos_publicados_pesq[,1] %in% paste0(concatenado_pesquisador[i,1],".zip"),3][j])
      } }
    
    #montando_coluna_de_artigos_aceitos B
    artigos_aceitos_pesq <- cbind(substr(l.out$tpublic.accepted[,1][[1]],1,20), l.out$tpublic.accepted[,2][[1]], l.out$tpublic.accepted[,3][[1]])
    
    for (i in 1:length(unique(concatenado_pesquisador[,1]))){
      #i=2
      for (j in 1:length(artigos_aceitos_pesq[artigos_aceitos_pesq[,1] %in% paste0(concatenado_pesquisador[i,1],".zip"),3]) ){
        #j=2
        concatenado_pesquisador[i,4] <- paste0(concatenado_pesquisador[i,4]," ", artigos_aceitos_pesq[artigos_aceitos_pesq[,1] %in% paste0(concatenado_pesquisador[i,1],".zip"),3][j])
      } }
    
    #montando_coluna_de_book C
    books_pesq <- cbind(substr(l.out$tbooks[,1][[1]],1,20), l.out$tbooks[,2][[1]], paste0(toupper(as.data.frame(l.out$tbooks['book.title'])[,1])," (", casefold(as.data.frame(l.out$tbooks['book.chapter'])[,1], upper = FALSE),")  | "))
    for (i in 1:length(unique(concatenado_pesquisador[,1]))){
      #i=35
      for (j in 1:length(books_pesq[books_pesq[,1] %in% paste0(concatenado_pesquisador[i,1],".zip"),3]) ){
        #j=2
        concatenado_pesquisador[i,5] <- paste0(concatenado_pesquisador[i,5]," ", books_pesq[books_pesq[,1] %in% paste0(concatenado_pesquisador[i,1],".zip"),3][j] )
      } }
    #montando_coluna_de_conferences D
    conferences_pesq <- cbind(substr(l.out$tconferences[,1][[1]],1,20), l.out$tconferences[,2][[1]], paste0(toupper(as.data.frame(l.out$tconferences['article.title'])[,1])," | ") )
    for (i in 1:length(unique(concatenado_pesquisador[,1]))){
      #i=35
      for (j in 1:length(conferences_pesq[conferences_pesq[,1] %in% paste0(concatenado_pesquisador[i,1],".zip"),3]) ){
        #j=2
        concatenado_pesquisador[i,6] <- paste0(" ",concatenado_pesquisador[i,6]," ", conferences_pesq[conferences_pesq[,1] %in% paste0(concatenado_pesquisador[i,1],".zip"),3][j] )
      } }
    
    concatenado_pesquisador <- cbind(concatenado_pesquisador, c(paste0(" ",concatenado_pesquisador[,3]," ",concatenado_pesquisador[,4]," ",concatenado_pesquisador[,5]," ",formacao_pesq," ", area_e_subarea_pesq," ", regiao," " ) ) )
    concatenado_pesquisador[,7] <- gsub(" NA ", "", concatenado_pesquisador[,7]) 
    texto_concatenado <- concatenado_pesquisador[,7]
    link_fotos <- read.csv2(paste0(here(),'/base_dados/lattes_zip/XML/links_fotos.csv'), sep=";")[,2]
    
    resumo_lattes <- matrix("", length(link_pesq), 2)
    #resumo_lattes[,1] <- substr(local,93,108)
    resumo_lattes[,1] <- substr(local,52,67)
    
    for (i in 1:length(resumo_lattes[,1])){
      #i=35
      texto <- ""
      texto <- xmlToList(unzip(local[i]))
      texto <- texto$`DADOS-GERAIS`$`RESUMO-CV`[[1]]
      if(is.null(texto)){resumo_lattes[i,2] <-"SEM RESUMO DO LATTES"}else {resumo_lattes[i,2] <- texto[1]}
    }
    
    
    base_expertise <- cbind(seq(1:length(local)), nomes_pesq, as.vector(resumo_lattes[,2]), as.vector(texto_concatenado), as.vector(link_pesq), as.vector(link_fotos) )
    base_expertise <- as.data.frame(base_expertise)
    names(base_expertise) <- c("ID", "NOME", "RESUMO LATTES", "CONCATENADO", "LINK LATTES", "LINK FOTO")
    #write.xlsx(base_expertise, paste0(here(),'/base_dados/base_expertise.xlsx'), rowNames = FALSE)
    
    base_expertise_complemento <- cbind(seq(1:length(local)), as.vector(regiao),  as.vector(formacao_pesq), as.vector(area_e_subarea_pesq))
    base_expertise_complemento <- as.data.frame(base_expertise_complemento)
    names(base_expertise_complemento) <- c("ID","regiao", "formacao_pesq", "area_e_subarea_pesq")
    #write.xlsx(base_expertise_complemento, paste0(here(),'/base_dados/base_expertise_complemento.xlsx'), rowNames = FALSE)
    
  }
}

#atualiza_tudo(TRUE)
base <<- data.frame(read.xlsx("base_dados/base_expertise.xlsx"))
base_complemento <<- data.frame(read.xlsx("base_dados/base_expertise_complemento.xlsx"))
fotos <- base[,6]
base <<- base[,-6]
fotos <- paste0('<img src=',fotos,' height=',52,' >', '</img>')

base <<- cbind(base, fotos, base_complemento[,2], base_complemento[,3], base_complemento[,4])
base <<- base[sample(1:nrow(base)), ]  


limpa_texto <- function(vetor_txt_bruto) {
  #tudo em letras minC:sculas
  vetor_txt_bruto <- tolower(vetor_txt_bruto)
  #removendo pontuaC'C#o
  vetor_txt_bruto <- removePunctuation(vetor_txt_bruto)
  #removendo nC:meros
  vetor_txt_bruto <- removeNumbers(vetor_txt_bruto)
  #removendo excesso de espaC'os
  vetor_txt_bruto <- stripWhitespace(vetor_txt_bruto)
  #removendo palavras de ligaC'C#o
  vetor_txt_bruto <-  removeWords(vetor_txt_bruto, c(stopwords("pt-br"), "fez", "grande")) 
  
  vetor_final <- as.list(NULL)
  for (i in 1:length(vetor_txt_bruto)) {
    vetor_final[[i]] <- unique(unlist(strsplit(vetor_txt_bruto[i], split = " "))) }
  vetor_final }

ui <- navbarPage( theme = shinytheme("spacelab"),
                  title = span(id="logo", tags$a(href="/#", span(id="logo", "DeuMatchNaPesquisa"))), 
                  windowTitle="DeuMatchNaPesquisa", responsive = NULL, collapsible = TRUE,
                  
                  tabPanel("Pesquisar",icon = icon("search-plus"),
                           
                           sidebarLayout(
                             sidebarPanel(
                               textInput("caption", "Objetivos do projeto", "digite aqui..."),
                               verbatimTextOutput("value"),
                               checkboxInput("resumo","Mostrar Resumo do Lattes",value=FALSE),
                               checkboxInput("regiao","Triângulo Mineiro",value=TRUE),
                             ),
                             
                             mainPanel(
                               tags$figure(
                                 align = "center",
                                 tags$img(src = "https://i.ibb.co/4N9y2mX/logo.png", width = 200, alt = "DeuMatchNaPesquisa"),
                               ),
                               
                               withLoader( DT::dataTableOutput('table'), type = "html", "loader5") ),
                           ) ),
                  
                  tabPanel("Como funciona?", icon = icon("question")),
                  tabPanel("Sobre", icon = icon("users")),
                  tabPanel("Contato", icon = icon("headset"),
                           mainPanel(
                             tags$figure(
                               align = "center",
                               tags$img(src = "https://i.ibb.co/4N9y2mX/logo.png", width="200px", alt = "DeuMatchNaPesquisa", align="center"),br(),br(),
                               tags$img(src = "https://www.gov.br/ebserh/pt-br/hospitais-universitarios/regiao-sudeste/hc-ufu/logos/hc-ufu-assinatura.png", width = "400px", alt = "HC-UFU / EBSERH", align="center") ),
                                br(),br(),h4("Entre em contato com a Unidade de Gestão da Inovação Tecnológica em Saúde - UGITS do HC-UFU/EBSERH"),

                             tags$table(
                               style = "width:100%",
                               tags$tr(
                                 tags$th("Whatsapp"),
                                 tags$td( tags$a(href=" https://wa.me/553432182323?text=Estou%20escrevendo%20para%20falar%20sobre%20o%20DEUMATCHNAPESQUISA.COM", "(34)3218-2323")) ),
                               tags$tr(
                                 tags$th("Email"),
                                 tags$td( tags$a(href="mailto:francisco.negrao@ebserh.gov.br", "francisco.negrao@ebserh.gov.br")) ),br(),
                             tags$tr(
                               tags$th("Endereço"),
                               tags$td(span("BLOCO 8F - R. República do Piratini, 1418"),br(),
                                       span("Umuarama, Uberlândia - MG, 38402-028 • Brasil"))  ) ),br(),
                             
                             span("Estamos guardando seu contato!"),
                             
                             
                             ) ) )


server <- function(input, output) {
  
  output$value <- renderText({ input$caption })
  
  df <- eventReactive(input$caption, {
    
    #Criando um vetor C:nico
    vetor_compat <- paste0(base[,3], base[,4], sep=" ")
    vetor_compat <- limpa_texto(vetor_compat)
    #vetor_compat
    
    vetor_frase_aluno <- input$caption
    #vetor_frase_aluno <- "pediatria"
    vetor_frase_aluno <- limpa_texto(vetor_frase_aluno)
    
    resultado <- matrix(NA, length(vetor_compat),2)
    colnames(resultado) <- c("ID","compatibilidade")
    resultado <- data.frame(resultado)
    for (i in 1:length(vetor_compat)) {
      resultado[i,1] <-i
      resultado[i,2] <-
        round(
          length(vetor_frase_aluno[[1]][vetor_frase_aluno[[1]] %in% vetor_compat[[i]]==TRUE]) 
          / length(vetor_frase_aluno[[1]]),3)*100
    }
    
    resultado <- data.frame(resultado[order(resultado$compatibilidade, decreasing = TRUE),])
    names(base)
    
    resultado <- cbind(base[resultado$ID,6], resultado, base[resultado$ID,2],base[resultado$ID,5],base[resultado$ID,8], base[resultado$ID,9], base[resultado$ID,7],base[resultado$ID,3]) 
    names(resultado) <- c("FOTO","ID", "COMPATIBILIDADE","NOME", "LINK", "FORMAÇÃO", "ÁREA", "LOCALIDADE", "RESUMO LATTES")
    resultado[,5] <- paste0('<a href=',resultado[,5],' target="_blank">',"Lattes", '</a>')
    
    resultado[,3] <- paste0(resultado[,3] , "%")
    
    resultado[if(f_tr()){resultado[,8] %in% "Triângulo Mineiro"} else {rep(TRUE, length(resultado[,2]))} , ]
    
  })
  
  f_tr <- eventReactive(input$regiao, {
    input$regiao
  })
  
  mostrar_resumo <- eventReactive(input$resumo, {
    input$resumo == TRUE
  })
  
  #if(f_tr()){data.frame(df())[,8] %in% "Triângulo Mineiro"}else{rep(TRUE, length(data.frame(df())[,8]))}
  
  output$table <- DT::renderDataTable({
    if ( mostrar_resumo() ) {
      DT::datatable(df()[,c(-2)], escape = FALSE, rownames= FALSE, options = list(dom = 't', pageLength = 150))
    } else { DT::datatable(df()[, c(-2,-9)], escape = FALSE, rownames= FALSE, options = list(dom = 't', pageLength = 150)) }
  } , 
  rownames= FALSE ) 
}

shinyApp(ui, server)




