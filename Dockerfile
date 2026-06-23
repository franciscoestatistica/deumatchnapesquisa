# Use a imagem base do R com Shiny
FROM rocker/shiny:latest

# Instalar dependências do sistema
RUN apt-get update && apt-get install -y \
    libcurl4-gnutls-dev \
    libssl-dev \
    libxml2-dev \
    libv8-dev \
    && rm -rf /var/lib/apt/lists/*

# Instalar pacotes R necessários
RUN R -e "install.packages(c('shiny', 'DT', 'tm', 'openxlsx', 'shinycustomloader', 'markdown', 'rsconnect', 'httr'), repos='https://cran.rstudio.com/')"

# Criar diretório da aplicação
WORKDIR /srv/shiny-server

# Remover aplicação padrão
RUN rm -rf /srv/shiny-server/*

# Copiar arquivos da aplicação
COPY app.R /srv/shiny-server/
COPY base_dados/ /srv/shiny-server/base_dados/

# Criar diretório www se não existir
RUN mkdir -p /srv/shiny-server/www

# Definir permissões
RUN chown -R shiny:shiny /srv/shiny-server

# Expor porta
EXPOSE 3838

# Comando para iniciar o servidor
CMD ["/usr/bin/shiny-server"]